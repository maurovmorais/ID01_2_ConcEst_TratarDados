"""Monta registros de fila (um por adquirente) com totais agrupados por
forma de pagamento, a partir da tabela tbl_dados_estruturados.

Fluxo:
    1. Busca todas as combinações distintas de (adquirente, forma_pagto).
    2. Para cada combinação, calcula os totais agrupados por
       "Nomenclatura SoftCase", bandeira, valor_taxa e taxa_adquirente.
    3. Agrupa os resultados por adquirente, com uma entrada por forma_pagto
       dentro de info_adicionais.
    4. Monta um dicionário pronto para inserção em tbl_fila (um por
       adquirente). A inserção em si fica comentada, aguardando
       confirmação explícita antes de gravar no banco.
"""

import json
import sqlite3
from datetime import datetime
from pathlib import Path
#from montar_fila_adquirentes import main, main_individual

QUERY_ADQUIRENTES_FORMAS = """
    SELECT DISTINCT adquirente, forma_pagto
    FROM tbl_dados_estruturados
    WHERE adquirente IS NOT NULL
      AND forma_pagto IS NOT NULL;
"""

QUERY_TOTAL_POR_FORMA = """
 SELECT adquirente,
       "Nomenclatura SoftCase" AS nomenclatura,
       bandeira,
       ROUND(SUM(valor_taxa), 2) AS valor_taxa,
       taxa_adquirente,
       forma_pagto,
       ROUND(SUM(valor_lancamento), 2) AS total
FROM tbl_dados_estruturados
WHERE adquirente = :adquirente
AND forma_pagto = :forma_pagto
AND "Nomenclatura SoftCase" IS NOT NULL
AND TRIM("Nomenclatura SoftCase") <> ''
GROUP BY adquirente, "Nomenclatura SoftCase", bandeira, taxa_adquirente, forma_pagto
HAVING SUM(valor_lancamento) IS NOT NULL
AND ROUND(SUM(valor_lancamento), 2) <> 0
ORDER BY "Nomenclatura SoftCase", bandeira;

"""

INSERT_FILA = """
    INSERT INTO tbl_Fila_Proc_Performer (referencia, datahora_criado, nome_maquina,
                           info_adicionais, status, obs, ultima_atualizacao)
    VALUES (:referencia, :datahora_criado, :nome_maquina,
            :info_adicionais, :status, :obs, :ultima_atualizacao);
"""


def _formatar_decimal(valor) -> str:
    """Converte um valor numérico ou textual para string com vírgula decimal.

    Args:
        valor: Número (int/float), texto já formatado, ou None.

    Returns:
        Representação em texto com vírgula como separador decimal.
        Retorna string vazia se o valor for None.
    """
    if valor is None:
        return ""
    return str(valor).replace(".", ",")


def combinacoes_adquirente_forma(
    conexao: sqlite3.Connection,
) -> list[tuple[str, str]]:
    """Retorna todas as combinações distintas de adquirente e forma_pagto.

    Args:
        conexao: Conexão aberta com a base SQLite.

    Returns:
        Lista de tuplas (adquirente, forma_pagto).
    """
    cursor = conexao.execute(QUERY_ADQUIRENTES_FORMAS)
    return [(linha[0], linha[1]) for linha in cursor.fetchall()]


def itens_por_forma(
    conexao: sqlite3.Connection,
    adquirente: str,
    forma_pagto: str,
) -> list[dict[str, str]]:
    """Retorna os totais detalhados para um adquirente e forma de pagamento.

    Args:
        conexao: Conexão aberta com a base SQLite.
        adquirente: Nome do adquirente.
        forma_pagto: Forma de pagamento.

    Returns:
        Lista de dicionários com adquirente, nomenclatura (softacase),
        bandeira, valor_taxa, taxa_adquirente, forma_pagto e valor (total),
        todos os campos numéricos formatados com vírgula decimal.
    """
    conexao.row_factory = sqlite3.Row
    cursor = conexao.execute(
        QUERY_TOTAL_POR_FORMA,
        {"adquirente": adquirente, "forma_pagto": forma_pagto},
    )
    return [
        {
            "adquirente": linha["adquirente"],
            "softacase": linha["nomenclatura"],
            "bandeira": linha["bandeira"],
            "valor_taxa": _formatar_decimal(linha["valor_taxa"]),
            "taxa_adquirente": _formatar_decimal(linha["taxa_adquirente"]),
            "forma_pagto": linha["forma_pagto"],
            "valor": _formatar_decimal(linha["total"]),
        }
        for linha in cursor.fetchall()
    ]


def montar_info_adicionais_por_adquirente(
    conexao: sqlite3.Connection,
) -> dict[str, list[dict]]:
    """Monta a info_adicionais agrupada por adquirente, em lista plana de itens."""
    combinacoes = combinacoes_adquirente_forma(conexao)

    por_adquirente: dict[str, list[dict]] = {}
    for adquirente, forma_pagto in combinacoes:
        itens = itens_por_forma(conexao, adquirente, forma_pagto)
        if not itens:
            continue
        por_adquirente.setdefault(adquirente, []).extend(itens)

    return por_adquirente


def montar_registros_fila(
    conexao: sqlite3.Connection,
    nome_maquina: str = "",
) -> list[dict]:
    """Monta um registro de fila por adquirente, pronto para inserção.

    Args:
        conexao: Conexão aberta com a base SQLite.
        nome_maquina: Nome da máquina de origem (opcional).

    Returns:
        Lista de dicionários prontos para inserir em tbl_fila.
    """
    agora = datetime.now()
    por_adquirente = montar_info_adicionais_por_adquirente(conexao)

    return [
        {
            "referencia": adquirente,
            "datahora_criado": agora.strftime("%d/%m/%Y %H:%M:%S"),
            "nome_maquina": nome_maquina,
            "info_adicionais": json.dumps(grupos, ensure_ascii=False),
            "status": "NEW",
            "obs": "",
            "ultima_atualizacao": agora.strftime("%Y-%m-%d %H:%M:%S.%f"),
        }
        for adquirente, grupos in por_adquirente.items()
    ]

def montar_registros_fila_individual(
    conexao: sqlite3.Connection,
    nome_maquina: str = "",
) -> list[dict]:
    agora = datetime.now()
    combinacoes = combinacoes_adquirente_forma(conexao)

    registros: list[dict] = []
    for adquirente, forma_pagto in combinacoes:
        itens = itens_por_forma(conexao, adquirente, forma_pagto)
        referencia = f"{adquirente}/{forma_pagto}"
        for item in itens:
            registros.append(
                {
                    "referencia": referencia,
                    "datahora_criado": agora.strftime("%d/%m/%Y %H:%M:%S"),
                    "nome_maquina": nome_maquina,
                    "info_adicionais": json.dumps([item], ensure_ascii=False),
                    "status": "NEW",
                    "obs": "",
                    "ultima_atualizacao": agora.strftime("%Y-%m-%d %H:%M:%S.%f"),
                }
            )

    return registros


def criar_fila_performer(caminho_banco: Path, nome_maquina: str = "") -> list[dict]:
    """Executa o fluxo completo: conecta, monta os registros e retorna a lista.

    A inserção em tbl_fila fica comentada de propósito — nenhuma gravação é
    feita sem confirmação explícita.

    Args:
        caminho_banco: Caminho para o arquivo SQLite.
        nome_maquina: Nome da máquina de origem (opcional).

    Returns:
        Lista de registros prontos para tbl_fila (um por adquirente).
    """
    with sqlite3.connect(caminho_banco) as conexao:
        #registros = montar_registros_fila(conexao, nome_maquina)
        registros = montar_registros_fila_individual(conexao, nome_maquina)

        # Inserção comentada de propósito — aguardando confirmação.
        conexao.executemany(INSERT_FILA, registros)
        conexao.commit()

    return registros

