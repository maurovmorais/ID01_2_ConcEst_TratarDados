"""Grava no banco SQLite os dados lidos do DEPARA_Estacionamento.xlsx.

Para cada planilha do DEPARA, cria (recriando, se já existir) a tabela
correspondente e insere os dados lidos por 'leitura_depara.py',
utilizando os nomes de tabela definidos em 'config.NOME_TABELAS'.
"""

from __future__ import annotations

import logging
from pathlib import Path

from ID01_2_ConcEst_TratarDados.classes.sqlite.sqlite import GerenciadorSQLite
from ID01_2_ConcEst_TratarDados.classes.sqlite.config import CAMINHO_BANCO_DADOS, NOME_TABELAS
from ID01_2_ConcEst_TratarDados.classes.excel.leitura_depara import ler_todas_planilhas

logger = logging.getLogger(__name__)

# Colunas (nome -> tipo SQL) de cada tabela, na mesma ordem das colunas
# da planilha correspondente no DEPARA.
ESQUEMA_TABELAS: dict[str, dict[str, str]] = {
    "Empresas": {
        "ID_UNICO": "TEXT",
        "Razao Social": "TEXT",
        "Nomenclatura SoftCase": "TEXT",
        "CNPJ": "TEXT",
        "SIGLA": "TEXT",
        "Cielo": "TEXT",
        "SemParar": "TEXT",
        "Greenpass": "TEXT",
        "ConectCar": "TEXT",
        "Veloe": "TEXT",
        "Bradesco": "TEXT",
        "Observação": "TEXT",
    },
    "FormaDePagamento": {
        "Forma de Pagamento": "TEXT",
        "Tipo": "TEXT",
        "Bandeira": "TEXT",
        "Valida": "TEXT",
        "Não atuar": "TEXT",
    },
    "CondPagamt": {
        "Tipo de Lançamento": "TEXT",
        "Dias Comp.": "INTEGER",
    },
    "Adquirentes": {
        "Lista de Adquirentes": "TEXT",
        "Extração via": "TEXT",
        "Observação": "TEXT",
    },
    "TaxaAdquirentes": {
        "ID_UNICO": "TEXT",
        "Sigla Empresa": "TEXT",
        "CONECTCAR": "REAL",
        "VELOE": "REAL",
        "GREENPASS": "REAL",
        "SEMPARAR": "REAL",
        "Coluna1": "TEXT",
    },
    "TaxaCielo": {
        "TAXAS CIELO": "TEXT",
        "Bandeira": "TEXT",
        # 'Taxa' mistura número (ex.: 0.01) e texto (PIX = 'Conforme
        # exportado'), por isso fica como TEXT.
        "Taxa": "TEXT",
    },
    "Dinheiro": {
        "Dia da Semana": "TEXT",
        "Dias Comp. Dinheiro": "INTEGER",
    },
}


def gravar_planilha_no_banco(
    db: GerenciadorSQLite,
    nome_planilha: str,
    linhas: list[dict[str, object]],
) -> None:
    """Cria a tabela de uma planilha (se necessário) e insere seus dados.

    Args:
        db: instância já conectada de 'GerenciadorSQLite'.
        nome_planilha: nome da planilha no DEPARA (chave em
            'config.NOME_TABELAS' e 'ESQUEMA_TABELAS').
        linhas: dados lidos da planilha (saída de 'leitura_depara').
    """
    if nome_planilha not in NOME_TABELAS or nome_planilha not in ESQUEMA_TABELAS:
        logger.warning(
            "Planilha '%s' sem tabela/esquema configurado; pulando gravação.",
            nome_planilha,
        )
        return

    nome_tabela = NOME_TABELAS[nome_planilha]
    esquema = ESQUEMA_TABELAS[nome_planilha]

    db.criar_tabela(nome_tabela, esquema)
    db.inserir_dados(nome_tabela, linhas)


def processar_depara(
    caminho_arquivo_depara: str | Path | None = None,
    caminho_banco_dados: str | Path = CAMINHO_BANCO_DADOS,
) -> None:
    """Lê todas as planilhas do DEPARA e grava cada uma em sua tabela.

    Args:
        caminho_arquivo_depara: caminho do DEPARA_Estacionamento.xlsx.
            Se None, usa o padrão definido em 'leitura_depara'/'config'.
        caminho_banco_dados: caminho do arquivo .db do SQLite. Por
            padrão usa 'config.CAMINHO_BANCO_DADOS'.
    """
    if caminho_arquivo_depara is None:
        dados_por_planilha = ler_todas_planilhas()
    else:
        dados_por_planilha = ler_todas_planilhas(caminho_arquivo_depara)

    with GerenciadorSQLite(caminho_banco_dados) as db:
        for nome_planilha, linhas in dados_por_planilha.items():
            gravar_planilha_no_banco(db, nome_planilha, linhas)


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    processar_depara()
