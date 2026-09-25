"""Leitura e extração dos relatórios de adquirentes (Excel/CSV).

Este módulo lê, para cada adquirente (ConectCar, Cielo, Greenpass,
SemParar e Veloe), o respectivo relatório baixado em
``ID01_1_ConcEst_Download/Arquivos_Baixados`` (formato .csv, .xlsx ou
.xls) e normaliza as colunas para o layout da tabela ``tbl_aux_dados``.

Este módulo só lê e trata os arquivos — não grava nada em banco.
A gravação fica a cargo do módulo
``classes/sqlite/gravador_tbl_aux_dados.py``.

Uso típico (dentro do process do projeto):

    from ID01_2_ConcEst_TratarDados.classes.excel.leitura_relat_adquirentes import (
        LeitorRelatoriosAdquirentes,
    )
    from ID01_2_ConcEst_TratarDados.classes.sqlite.gravador_tbl_aux_dados import (
        GravadorTblAuxDados,
    )

    leitor = LeitorRelatoriosAdquirentes()
    dados_consolidados = leitor.processar_todos()

    gravador = GravadorTblAuxDados()
    gravador.gravar_tbl_aux_dados(dados_consolidados)
"""

from __future__ import annotations

from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import (
    InitAllSettings,
)

import csv
import io
import logging
import re
import warnings
from datetime import date
from pathlib import Path
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------

DIRETORIO_ARQUIVOS_PADRAO = Path(InitAllSettings.config["arquivos_baixados"])

EXTENSOES_SUPORTADAS = (".csv", ".xlsx", ".xls")

# Ordem de tentativa de codificação para arquivos .csv. latin1 nunca
# falha (mapeia qualquer byte), por isso fica como última opção.
CODIFICACOES_CSV = ("utf-8", "utf-8-sig", "cp1252", "latin1")

# Separadores testados para arquivos .csv, na ordem de tentativa. O
# separador correto é confirmado batendo com o cabeçalho esperado do
# adquirente (ver "colunas_identificacao_cabecalho"), em vez de
# adivinhado — relatórios com texto de preâmbulo (datas, frases com
# vírgula) podem enganar uma detecção automática de separador.
SEPARADORES_CSV = (",", ";", "\t", "|")

# Quantidade máxima de linhas iniciais varridas em busca da linha de
# cabeçalho real do relatório (alguns adquirentes exportam linhas de
# preâmbulo — dados do usuário, filtros aplicados etc. — antes da
# tabela propriamente dita).
LIMITE_LINHAS_BUSCA_CABECALHO = 60

COLUNAS_TABELA_DESTINO = (
    "adquirente",
    "empresa",
    "data_processamento",
    "forma_pagto",
    "Bandeira",
    "valor_lancamento",
    "valor_taxa",
    "data_atualizacao",
)

# Mapeamento: chave usada para localizar o arquivo -> nome do
# adquirente gravado na coluna "adquirente" -> nomes das colunas do
# relatório original que correspondem a cada coluna da tbl_aux_dados.
# Um valor None indica que o relatório não possui aquela informação;
# nesse caso a coluna é gravada como NULL.
MAPEAMENTO_ADQUIRENTES: dict[str, dict] = {
    "conectcar": {
        "nome_adquirente": "ConectCar",
        "colunas_identificacao_cabecalho": (
            "Data do Processamento",
            "Hora do Processamento",
        ),
        "coluna_empresa": "Conveniado",
        "coluna_data_processamento": "Data do Processamento",
        "coluna_forma_pagto": None,
        "coluna_bandeira": None,
        "coluna_valor_lancamento": "Valor Cobrado",
        "coluna_valor_taxa": "Tarifa de interconexão",
    },
    "cielo": {
        "nome_adquirente": "Cielo",
        "colunas_identificacao_cabecalho": ("Data da venda", "Hora da venda"),
        "coluna_empresa": "CPF/CNPJ do estabelecimento",
        "coluna_data_processamento": "Data da venda",
        "coluna_forma_pagto": "Forma de pagamento",
        "coluna_bandeira": "Bandeira",
        "coluna_valor_lancamento": "Valor bruto",
        "coluna_valor_taxa": "Taxa/tarifa",
        # A Cielo grava o CPF/CNPJ formatado (ex.: "13.008.381/0002-88");
        # mantém-se só os dígitos.
        "limpar_documento_empresa": True,
    },
    "greenpass": {
        "nome_adquirente": "Greenpass",
        "colunas_identificacao_cabecalho": ("DATA ESTADIA", "DATA RECEBIMENTO"),
        "coluna_empresa": "CONVENIADO",
        "coluna_data_processamento": "DATA ESTADIA",
        "coluna_forma_pagto": None,
        "coluna_bandeira": None,
        "coluna_valor_lancamento": "VALOR TOTAL",
        "coluna_valor_taxa": None,
    },
    "semparar": {
        "nome_adquirente": "SemParar",
        "colunas_identificacao_cabecalho": ("credenciado", "valor_total"),
        "coluna_empresa": "credenciado",
        "coluna_data_processamento": "data_periodo",
        "coluna_forma_pagto": None,
        "coluna_bandeira": None,
        "coluna_valor_lancamento": "valor_total",
        "coluna_valor_taxa": None,
    },
    "veloe": {
        "nome_adquirente": "Veloe",
        "colunas_identificacao_cabecalho": ("CNPJ", "Estabelecimento"),
        "coluna_empresa": "Estabelecimento",
        "coluna_data_processamento": "Data Saída",
        "coluna_forma_pagto": "Tipo",
        "coluna_bandeira": None,
        "coluna_valor_lancamento": "Valor Transação",
        "coluna_valor_taxa": "Valor MDR",
    },
    "bradesco": {
        "nome_adquirente": "Bradesco",
        "colunas_identificacao_cabecalho": ("data_arquivo_ret", "data_movimento"),
        "coluna_empresa": "inscricao_empresa",
        "coluna_data_processamento": "data_lancamento",
        "coluna_forma_pagto": "historico",
        "coluna_bandeira": None,
        "coluna_valor_lancamento": "valor_assinado",
        "coluna_valor_taxa": None,
    },
}


class ArquivoAdquirenteNaoEncontrado(FileNotFoundError):
    """Levantada quando não há arquivo do adquirente no diretório."""


class LeitorRelatoriosAdquirentes:
    """Lê os relatórios de cada adquirente e consolida no layout padrão.

    Attributes:
        diretorio_arquivos: pasta onde os relatórios baixados ficam.
    """

    def __init__(
        self, diretorio_arquivos: Path = DIRETORIO_ARQUIVOS_PADRAO
    ) -> None:
        self.diretorio_arquivos = diretorio_arquivos

    # -----------------------------------------------------------------
    # Funções públicas — uma por adquirente
    # -----------------------------------------------------------------

    def conectcar(self) -> pd.DataFrame:
        """Lê e normaliza o relatório da ConectCar."""
        return self._processar_adquirente("conectcar")

    def cielo(self) -> pd.DataFrame:
        """Lê e normaliza o relatório da Cielo."""
        return self._processar_adquirente("cielo")

    def greenpass(self) -> pd.DataFrame:
        """Lê e normaliza o relatório da Greenpass."""
        return self._processar_adquirente("greenpass")

    def semparar(self) -> pd.DataFrame:
        """Lê e normaliza o relatório da SemParar."""
        return self._processar_adquirente("semparar")

    def veloe(self) -> pd.DataFrame:
        """Lê e normaliza o relatório da Veloe."""
        return self._processar_adquirente("veloe")

    def bradesco(self) -> pd.DataFrame:
        """Lê e normaliza o relatório da Veloe."""
        return self._processar_adquirente("bradesco")

    # -----------------------------------------------------------------
    # Orquestração da extração
    # -----------------------------------------------------------------

    def processar_todos(self) -> pd.DataFrame:
        """Processa todos os adquirentes cadastrados no mapeamento.

        Adquirentes cujo arquivo não é encontrado são ignorados (com
        aviso no log) para não interromper o processamento dos demais.

        Returns:
            DataFrame consolidado, já no layout da tbl_aux_dados.
        """
        funcoes_por_chave = {
            "conectcar": self.conectcar,
            "cielo": self.cielo,
            "greenpass": self.greenpass,
            "semparar": self.semparar,
            "veloe": self.veloe,
            "bradesco": self.bradesco,
        }

        dataframes_processados: list[pd.DataFrame] = []
        for chave_adquirente, funcao_extracao in funcoes_por_chave.items():
            try:
                dataframes_processados.append(funcao_extracao())
            except ArquivoAdquirenteNaoEncontrado:
                logger.warning(
                    "Arquivo do adquirente '%s' não encontrado em %s; "
                    "adquirente ignorado nesta execução.",
                    chave_adquirente,
                    self.diretorio_arquivos,
                )
            except (pd.errors.ParserError, ValueError, KeyError) as erro:
                logger.error(
                    "Falha ao processar o adquirente '%s': %s",
                    chave_adquirente,
                    erro,
                )

        if not dataframes_processados:
            logger.warning("Nenhum relatório de adquirente foi processado.")
            return pd.DataFrame(columns=COLUNAS_TABELA_DESTINO)

        # Colunas como forma_pagto/Bandeira/valor_taxa ficam inteiramente
        # None para alguns adquirentes (ex.: Greenpass). O pandas emite um
        # FutureWarning inofensivo ao concatenar esse tipo de coluna
        # totalmente vazia junto com outras que têm dados; o
        # comportamento atual (o que queremos) já é preservado.
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore",
                message=".*empty or all-NA entries.*",
                category=FutureWarning,
            )
            return pd.concat(dataframes_processados, ignore_index=True)

    # -----------------------------------------------------------------
    # Extração e normalização (uso interno)
    # -----------------------------------------------------------------

    def _processar_adquirente(self, chave_adquirente: str) -> pd.DataFrame:
        """Localiza, lê e normaliza o(s) relatório(s) de um adquirente.

        Um adquirente pode ter mais de um arquivo no diretório (ex.:
        Cielo, que exporta um .csv por período em
        "Vendas_cielo_historico_..."); todos os arquivos encontrados
        são lidos e consolidados antes da normalização.

        Args:
            chave_adquirente: chave em MAPEAMENTO_ADQUIRENTES
                (ex.: "cielo").

        Returns:
            DataFrame com as colunas de COLUNAS_TABELA_DESTINO.

        Raises:
            ArquivoAdquirenteNaoEncontrado: se nenhum arquivo do
                adquirente for encontrado no diretório configurado.
        """
        configuracao = MAPEAMENTO_ADQUIRENTES[chave_adquirente]
        caminhos_arquivos = self._localizar_arquivos(chave_adquirente)

        logger.info(
            "%d arquivo(s) encontrado(s) para o adquirente '%s': %s",
            len(caminhos_arquivos),
            configuracao["nome_adquirente"],
            ", ".join(caminho.name for caminho in caminhos_arquivos),
        )

        tabelas_extraidas = []
        for caminho_arquivo in caminhos_arquivos:
            dados_brutos = self._ler_arquivo(caminho_arquivo, configuracao)
            tabelas_extraidas.append(
                self._extrair_tabela(dados_brutos, configuracao)
            )

        if len(tabelas_extraidas) == 1:
            dados_originais = tabelas_extraidas[0]
        else:
            with warnings.catch_warnings():
                warnings.filterwarnings(
                    "ignore",
                    message=".*empty or all-NA entries.*",
                    category=FutureWarning,
                )
                dados_originais = pd.concat(tabelas_extraidas, ignore_index=True)

        return self._montar_dataframe_padrao(dados_originais, configuracao)

    def _localizar_arquivos(self, chave_adquirente: str) -> list[Path]:
        """Procura no diretório de arquivos todos os relatórios do adquirente.

        A identificação é feita pelo nome do arquivo conter a chave do
        adquirente, ignorando maiúsculas/minúsculas e caracteres não
        alfanuméricos (ex.: "SEM_PARAR.xlsx" e "Vendas_cielo_historico_1.csv"
        casam com as chaves "semparar" e "cielo", respectivamente).
        Um mesmo adquirente pode ter vários arquivos compatíveis — todos
        são retornados, em ordem alfabética.

        Args:
            chave_adquirente: trecho esperado no nome do arquivo.

        Returns:
            Lista (ordenada) dos caminhos dos arquivos compatíveis.

        Raises:
            ArquivoAdquirenteNaoEncontrado: se não houver nenhum
                arquivo compatível no diretório.
        """
        if not self.diretorio_arquivos.exists():
            raise ArquivoAdquirenteNaoEncontrado(
                f"Diretório não encontrado: {self.diretorio_arquivos}"
            )

        chave_normalizada = self._normalizar_nome(chave_adquirente)
        arquivos_encontrados = []
        for caminho in sorted(self.diretorio_arquivos.iterdir()):
            nome_normalizado = self._normalizar_nome(caminho.stem)
            extensao_suportada = caminho.suffix.lower() in EXTENSOES_SUPORTADAS
            if chave_normalizada in nome_normalizado and extensao_suportada:
                arquivos_encontrados.append(caminho)

        if not arquivos_encontrados:
            raise ArquivoAdquirenteNaoEncontrado(
                f"Nenhum arquivo de '{chave_adquirente}' encontrado em "
                f"{self.diretorio_arquivos}"
            )

        return arquivos_encontrados

    @staticmethod
    def _normalizar_nome(texto: str) -> str:
        """Normaliza um nome para comparação (minúsculas, só alfanum).

        Args:
            texto: texto original.

        Returns:
            Texto em minúsculas, sem espaços, "_", "-" ou outros
            caracteres não alfanuméricos.
        """
        return re.sub(r"[^a-z0-9]", "", texto.lower())

    def _ler_arquivo(
        self, caminho_arquivo: Path, configuracao: dict
    ) -> pd.DataFrame:
        """Lê um arquivo .csv, .xlsx ou .xls como DataFrame bruto.

        A leitura é feita sem assumir qual linha é o cabeçalho, pois
        alguns adquirentes exportam linhas de preâmbulo antes da
        tabela de dados; a localização da linha de cabeçalho real é
        feita depois, em ``_extrair_tabela``.

        Args:
            caminho_arquivo: caminho do arquivo a ser lido.
            configuracao: entrada correspondente em
                MAPEAMENTO_ADQUIRENTES (usada, no caso de .csv, para
                confirmar qual separador é o correto).

        Returns:
            DataFrame bruto (sem cabeçalho definido, tudo como texto).

        Raises:
            ValueError: se a extensão do arquivo não for suportada, ou
                se não for possível identificar codificação/separador
                do .csv.
        """
        extensao = caminho_arquivo.suffix.lower()

        if extensao == ".csv":
            return self._ler_csv(caminho_arquivo, configuracao)

        if extensao in (".xlsx", ".xls"):
            return pd.read_excel(caminho_arquivo, header=None, dtype=str)

        raise ValueError(f"Extensão não suportada: {extensao}")

    def _ler_csv(self, caminho_arquivo: Path, configuracao: dict) -> pd.DataFrame:
        """Lê um .csv testando codificação e separador até achar o certo.

        Em vez de deixar o separador ser "adivinhado" (o que falha em
        relatórios com texto de preâmbulo contendo vírgulas, datas
        etc.), cada combinação de codificação/separador é validada
        contra o cabeçalho esperado do adquirente
        (``colunas_identificacao_cabecalho``) antes de ser aceita. A
        leitura usa o módulo ``csv`` da biblioteca padrão, que não
        exige que todas as linhas tenham a mesma quantidade de
        campos — diferente do parser do pandas.

        Args:
            caminho_arquivo: caminho do arquivo .csv.
            configuracao: entrada correspondente em
                MAPEAMENTO_ADQUIRENTES.

        Returns:
            DataFrame bruto (sem cabeçalho definido, tudo como texto),
            com todas as linhas preenchidas até a maior quantidade de
            campos encontrada no arquivo.

        Raises:
            ValueError: se nenhuma combinação de codificação e
                separador testada encontrar o cabeçalho esperado.
        """
        coluna_1_esperada, coluna_2_esperada = configuracao[
            "colunas_identificacao_cabecalho"
        ]
        coluna_1_esperada = coluna_1_esperada.strip()
        coluna_2_esperada = coluna_2_esperada.strip()

        ultimo_erro: Optional[UnicodeDecodeError] = None
        for codificacao in CODIFICACOES_CSV:
            try:
                texto = caminho_arquivo.read_text(encoding=codificacao)
            except UnicodeDecodeError as erro:
                ultimo_erro = erro
                continue

            for separador in SEPARADORES_CSV:
                linhas = list(csv.reader(io.StringIO(texto), delimiter=separador))
                limite = min(LIMITE_LINHAS_BUSCA_CABECALHO, len(linhas))
                cabecalho_bate = any(
                    self._normalizar_celula_cabecalho(
                        linhas[indice][0] if len(linhas[indice]) > 0 else None
                    )
                    == coluna_1_esperada
                    and self._normalizar_celula_cabecalho(
                        linhas[indice][1] if len(linhas[indice]) > 1 else None
                    )
                    == coluna_2_esperada
                    for indice in range(limite)
                )
                if cabecalho_bate:
                    return self._construir_dataframe_bruto(linhas)

        raise ValueError(
            f"Não foi possível ler {caminho_arquivo.name}: nenhuma "
            "combinação de codificação/separador testada encontrou o "
            f"cabeçalho esperado ('{coluna_1_esperada}', "
            f"'{coluna_2_esperada}')."
        ) from ultimo_erro

    @staticmethod
    def _construir_dataframe_bruto(linhas: list[list[str]]) -> pd.DataFrame:
        """Monta um DataFrame a partir de linhas de tamanhos variados.

        Cada linha é completada com None até a maior quantidade de
        campos encontrada no arquivo, para todas terem o mesmo
        tamanho (exigência do pandas).

        Args:
            linhas: linhas já separadas em campos (via csv.reader).

        Returns:
            DataFrame bruto, sem cabeçalho definido.
        """
        maior_quantidade_campos = max((len(linha) for linha in linhas), default=0)
        linhas_preenchidas = [
            linha + [None] * (maior_quantidade_campos - len(linha))
            for linha in linhas
        ]
        return pd.DataFrame(linhas_preenchidas)

    def _extrair_tabela(
        self, dados_brutos: pd.DataFrame, configuracao: dict
    ) -> pd.DataFrame:
        """Localiza a linha de cabeçalho real e recorta a tabela.

        Alguns relatórios trazem linhas de preâmbulo (dados do
        usuário, filtros aplicados etc.) antes da tabela de dados, e
        cada adquirente pode começar em uma linha diferente. Esta
        função varre as primeiras linhas do arquivo em busca da que
        tem, nas duas primeiras posições, os nomes de coluna
        esperados (ver "colunas_identificacao_cabecalho" em
        MAPEAMENTO_ADQUIRENTES) e usa essa linha como cabeçalho.

        Args:
            dados_brutos: DataFrame lido sem cabeçalho definido.
            configuracao: entrada correspondente em
                MAPEAMENTO_ADQUIRENTES.

        Returns:
            DataFrame com o cabeçalho real já aplicado.

        Raises:
            ValueError: se a linha de cabeçalho não for encontrada
                dentro do limite de linhas varridas.
        """
        coluna_1_esperada, coluna_2_esperada = configuracao[
            "colunas_identificacao_cabecalho"
        ]

        limite = min(LIMITE_LINHAS_BUSCA_CABECALHO, len(dados_brutos))
        for indice_linha in range(limite):
            linha = dados_brutos.iloc[indice_linha]
            valor_coluna_1 = self._normalizar_celula_cabecalho(
                linha.iloc[0] if len(linha) > 0 else None
            )
            valor_coluna_2 = self._normalizar_celula_cabecalho(
                linha.iloc[1] if len(linha) > 1 else None
            )
            if (
                valor_coluna_1 == coluna_1_esperada.strip()
                and valor_coluna_2 == coluna_2_esperada.strip()
            ):
                tabela = dados_brutos.iloc[indice_linha + 1 :].copy()
                tabela.columns = self._deduplicar_nomes_coluna(
                    self._normalizar_celula_cabecalho(valor)
                    for valor in dados_brutos.iloc[indice_linha]
                )
                return tabela.reset_index(drop=True)

        raise ValueError(
            "Linha de cabeçalho não encontrada nas primeiras "
            f"{limite} linhas do arquivo do adquirente "
            f"'{configuracao['nome_adquirente']}'. Esperava as colunas "
            f"'{coluna_1_esperada}' e '{coluna_2_esperada}' nas duas "
            "primeiras posições."
        )

    @staticmethod
    def _normalizar_celula_cabecalho(valor: object) -> str:
        """Normaliza uma célula para comparação de cabeçalho.

        Remove espaços nas pontas e o caractere BOM ("\\ufeff") que
        alguns arquivos .csv trazem na primeira célula.

        Args:
            valor: valor bruto da célula (pode ser NaN).

        Returns:
            Texto normalizado, ou string vazia quando ausente.
        """
        if pd.isna(valor):
            return ""
        return str(valor).strip().lstrip("\ufeff")

    @staticmethod
    def _deduplicar_nomes_coluna(nomes) -> list[str]:
        """Garante nomes de coluna únicos, preservando a primeira ocorrência.

        Alguns relatórios (ex.: Cielo) repetem o mesmo nome de coluna
        mais de uma vez (ex.: duas colunas "Taxa/tarifa" com
        significados diferentes). A primeira ocorrência mantém o nome
        original — é a que o mapeamento em MAPEAMENTO_ADQUIRENTES
        espera — e as seguintes recebem um sufixo, para não colidir.

        Args:
            nomes: nomes de coluna na ordem em que aparecem no
                relatório.

        Returns:
            Lista de nomes únicos, na mesma ordem.
        """
        contagem: dict[str, int] = {}
        nomes_unicos: list[str] = []
        for nome in nomes:
            ocorrencias = contagem.get(nome, 0)
            contagem[nome] = ocorrencias + 1
            if ocorrencias == 0:
                nomes_unicos.append(nome)
            else:
                nomes_unicos.append(f"{nome}__duplicada_{ocorrencias}")
        return nomes_unicos

    def _montar_dataframe_padrao(
        self, dados_originais: pd.DataFrame, configuracao: dict
    ) -> pd.DataFrame:
        """Converte o DataFrame bruto para o layout da tbl_aux_dados.

        Args:
            dados_originais: DataFrame já com o cabeçalho real
                aplicado (ver ``_extrair_tabela``).
            configuracao: entrada correspondente em
                MAPEAMENTO_ADQUIRENTES.

        Returns:
            DataFrame com as colunas de COLUNAS_TABELA_DESTINO, sem
            linhas de rodapé (que não têm data de processamento
            válida).
        """
        quantidade_linhas = len(dados_originais)
        dados_normalizados = pd.DataFrame(index=range(quantidade_linhas))

        dados_normalizados["adquirente"] = configuracao["nome_adquirente"]

        coluna_empresa = self._obter_coluna(
            dados_originais, configuracao["coluna_empresa"]
        )
        if configuracao.get("limpar_documento_empresa"):
            coluna_empresa = self._limpar_documento(coluna_empresa)
        dados_normalizados["empresa"] = coluna_empresa

        dados_normalizados["data_processamento"] = self._converter_data(
            self._obter_coluna(
                dados_originais, configuracao["coluna_data_processamento"]
            )
        )
        dados_normalizados["forma_pagto"] = self._obter_coluna(
            dados_originais, configuracao["coluna_forma_pagto"]
        )
        dados_normalizados["Bandeira"] = self._obter_coluna(
            dados_originais, configuracao["coluna_bandeira"]
        )
        dados_normalizados["valor_lancamento"] = self._converter_valor(
            self._obter_coluna(
                dados_originais, configuracao["coluna_valor_lancamento"]
            )
        )
        dados_normalizados["valor_taxa"] = self._converter_valor(
            self._obter_coluna(dados_originais, configuracao["coluna_valor_taxa"])
        )
        dados_normalizados["data_atualizacao"] = date.today().isoformat()

        dados_normalizados = dados_normalizados[list(COLUNAS_TABELA_DESTINO)]

        # Descarta linhas sem data de processamento válida: alguns
        # relatórios trazem linhas de rodapé (totais, percentuais)
        # após a tabela de dados, que não representam transações.
        quantidade_antes = len(dados_normalizados)
        dados_normalizados = dados_normalizados[
            dados_normalizados["data_processamento"].notna()
        ].reset_index(drop=True)
        quantidade_descartada = quantidade_antes - len(dados_normalizados)
        if quantidade_descartada:
            logger.warning(
                "%d linha(s) do adquirente '%s' descartada(s) por não "
                "terem data de processamento válida (provável rodapé "
                "do relatório).",
                quantidade_descartada,
                configuracao["nome_adquirente"],
            )

        return dados_normalizados

    @staticmethod
    def _obter_coluna(
        dados: pd.DataFrame, nome_coluna: Optional[str]
    ) -> pd.Series:
        """Retorna a coluna pedida, ou uma série de None se ausente.

        Args:
            dados: DataFrame de origem.
            nome_coluna: nome da coluna no relatório original, ou
                None quando o relatório não possui essa informação.

        Returns:
            Série com os valores da coluna, ou série de None.
        """
        if nome_coluna is None or nome_coluna not in dados.columns:
            return pd.Series([None] * len(dados), index=dados.index)
        return dados[nome_coluna]

    @staticmethod
    def _limpar_documento(valores: pd.Series) -> pd.Series:
        """Mantém apenas os dígitos de um CPF/CNPJ formatado.

        Remove ".", "/", "-" e qualquer outro caractere não numérico
        (ex.: "13.008.381/0002-88" -> "13008381000288").

        Args:
            valores: série com o documento formatado.

        Returns:
            Série com apenas os dígitos, ou None quando ausente.
        """

        def _limpar_um_valor(valor: object) -> Optional[str]:
            if valor is None or (isinstance(valor, float) and pd.isna(valor)):
                return None
            apenas_digitos = re.sub(r"[^0-9]", "", str(valor).strip())
            return apenas_digitos or None

        return valores.apply(_limpar_um_valor)

    @staticmethod
    def _converter_valor(valores: pd.Series) -> pd.Series:
        """Converte valores monetários em formato brasileiro para float.

        Remove símbolos de moeda ("R$"), espaços e qualquer outro
        caractere que não seja dígito, separador ou sinal, depois
        trata separador de milhar "." e decimal "," (ex.:
        "R$ 1.234,56" -> 1234.56). Valores já numéricos ou ausentes
        são preservados.

        Args:
            valores: série com os valores monetários brutos.

        Returns:
            Série com valores convertidos para float (ou None).
        """

        def _converter_um_valor(valor: object) -> Optional[float]:
            if valor is None or (isinstance(valor, float) and pd.isna(valor)):
                return None
            if isinstance(valor, (int, float)):
                return float(valor)
            try:
                texto_bruto = str(valor).strip()
                texto_numerico = re.sub(r"[^0-9,.\-]", "", texto_bruto)
                texto_limpo = texto_numerico.replace(".", "").replace(",", ".")
                return float(texto_limpo)
            except ValueError:
                logger.warning("Valor monetário inválido ignorado: %r", valor)
                return None

        return valores.apply(_converter_um_valor)

    @staticmethod
    def _converter_data(valores: pd.Series) -> pd.Series:
        """Converte datas para o formato ISO (YYYY-MM-DD).

        Args:
            valores: série com as datas brutas do relatório.

        Returns:
            Série com datas em formato ISO (texto), ou None quando a
            conversão falhar.
        """
        # Alguns adquirentes (ex.: SemParar) gravam um intervalo
        # "14/09/2026 - 14/09/2026" em vez de uma data única; nesse
        # caso considera-se a primeira data do intervalo.
        valores_sem_intervalo = valores.astype("string").str.replace(
            r"\s+-\s+.*$", "", regex=True
        )

        datas_convertidas = pd.to_datetime(
            valores_sem_intervalo, dayfirst=True, errors="coerce", format="mixed"
        )
        return datas_convertidas.dt.strftime("%Y-%m-%d")
