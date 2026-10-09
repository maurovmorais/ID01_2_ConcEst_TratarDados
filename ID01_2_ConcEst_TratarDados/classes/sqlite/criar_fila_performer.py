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
  WITH normalizado AS (
    SELECT
        dados.*,
        CASE
            WHEN dados.data_processamento LIKE '__/__/____'
                THEN substr(dados.data_processamento, 7, 4) || '-' ||
                     substr(dados.data_processamento, 4, 2) || '-' ||
                     substr(dados.data_processamento, 1, 2)
            ELSE substr(dados.data_processamento, 1, 10)
        END AS data_processamento_iso,
        CASE
            WHEN instr(dados."Nomenclatura SoftCase", '(') > 0
             AND instr(dados."Nomenclatura SoftCase", ')') >
                 instr(dados."Nomenclatura SoftCase", '(')
                THEN TRIM(substr(
                    dados."Nomenclatura SoftCase",
                    instr(dados."Nomenclatura SoftCase", '(') + 1,
                    instr(dados."Nomenclatura SoftCase", ')')
                        - instr(dados."Nomenclatura SoftCase", '(') - 1
                ))
        END AS origem_extraida
    FROM tbl_dados_estruturados AS dados
),
forma AS (
    SELECT
        UPPER(TRIM("Origem")) AS origem,
        CASE
            WHEN "Bandeira" IS NULL OR TRIM("Bandeira") = ''
                THEN UPPER(TRIM("Tipo"))
            ELSE UPPER(TRIM("Bandeira"))
        END AS bandeira,
        UPPER(TRIM("Tipo")) AS tipo,
        MIN("Forma de Pagamento") AS forma_pagamento
    FROM tbl_FormaDePagamento
    WHERE LOWER(TRIM("Forma de Pagamento Principal")) = 'true'
    GROUP BY
        UPPER(TRIM("Origem")),
        CASE
            WHEN "Bandeira" IS NULL OR TRIM("Bandeira") = ''
                THEN UPPER(TRIM("Tipo"))
            ELSE UPPER(TRIM("Bandeira"))
        END,
        UPPER(TRIM("Tipo"))
),
forma_dinheiro AS (
    SELECT
        UPPER(TRIM("Origem")) AS origem,
        MIN("Forma de Pagamento") AS forma_pagamento
    FROM tbl_FormaDePagamento
    WHERE LOWER(TRIM("Forma de Pagamento Principal")) = 'true'
      AND UPPER(TRIM("Tipo")) = 'DINHEIRO'
    GROUP BY UPPER(TRIM("Origem"))
)
SELECT
    normalizado.adquirente,
    normalizado."Nomenclatura SoftCase" AS nomenclatura,
    normalizado.bandeira,
    CASE
        WHEN UPPER(normalizado.forma_pagto) = 'DINHEIRO' THEN dinheiro."Dias Comp. Dinheiro"
        ELSE normalizado.Dia_Comp
    END AS Dia_Comp,
    CASE
        WHEN normalizado.adquirente IN ('Greenpass', 'SemParar')
            THEN ROUND(SUM(normalizado.valor_lancamento) * CAST(REPLACE(normalizado.taxa_adquirente, ',', '.') AS REAL) / 100, 2)
        ELSE ROUND(SUM(normalizado.valor_taxa), 2)
    END AS valor_taxa,
    normalizado.taxa_adquirente,
    normalizado.forma_pagto,
    CASE
        WHEN UPPER(TRIM(normalizado.forma_pagto)) = 'TAG'
            THEN normalizado.adquirente
        WHEN UPPER(TRIM(normalizado.forma_pagto)) = 'DINHEIRO'
         AND (normalizado.bandeira IS NULL OR TRIM(normalizado.bandeira) = '')
            THEN COALESCE(forma.forma_pagamento, forma_dinheiro.forma_pagamento)
        ELSE forma.forma_pagamento
    END AS forma_pagamento,
    ROUND(SUM(normalizado.valor_lancamento), 2) AS total,
    normalizado.data_processamento,
    normalizado.data_processamento_iso || ' - ' ||
    CASE CAST(strftime('%w', normalizado.data_processamento_iso) AS INTEGER)
        WHEN 0 THEN 'DOMINGO'
        WHEN 1 THEN 'SEGUNDA'
        WHEN 2 THEN 'TERÇA'
        WHEN 3 THEN 'QUARTA'
        WHEN 4 THEN 'QUINTA'
        WHEN 5 THEN 'SEXTA'
        WHEN 6 THEN 'SABADO'
    END AS data_ref
FROM normalizado
LEFT JOIN tbl_Dinheiro AS dinheiro
    ON UPPER(dinheiro."Dia da Semana") = CASE CAST(strftime('%w', normalizado.data_processamento_iso) AS INTEGER)
        WHEN 0 THEN 'DOMINGO'
        WHEN 1 THEN 'SEGUNDA'
        WHEN 2 THEN 'TERÇA'
        WHEN 3 THEN 'QUARTA'
        WHEN 4 THEN 'QUINTA'
        WHEN 5 THEN 'SEXTA'
        WHEN 6 THEN 'SABADO'
    END
LEFT JOIN forma
    ON forma.origem   = UPPER(normalizado.origem_extraida)
   AND forma.bandeira = UPPER(COALESCE(NULLIF(TRIM(normalizado.bandeira), ''),
                                       TRIM(normalizado.forma_pagto)))
   AND forma.tipo     = UPPER(TRIM(normalizado.forma_pagto))
LEFT JOIN forma_dinheiro
    ON forma_dinheiro.origem = UPPER(normalizado.origem_extraida)
WHERE normalizado.adquirente = :adquirente
  AND normalizado.forma_pagto = :forma_pagto
  AND normalizado."Nomenclatura SoftCase" IS NOT NULL
  AND TRIM(normalizado."Nomenclatura SoftCase") <> ''
GROUP BY
    normalizado.adquirente,
    normalizado."Nomenclatura SoftCase",
    normalizado.bandeira,
    normalizado.taxa_adquirente,
    normalizado.forma_pagto,
    normalizado.data_processamento,
    dinheiro."Dias Comp. Dinheiro",
    forma.forma_pagamento,
    forma_dinheiro.forma_pagamento
HAVING SUM(normalizado.valor_lancamento) IS NOT NULL
   AND ROUND(SUM(normalizado.valor_lancamento), 2) <> 0
ORDER BY normalizado."Nomenclatura SoftCase", normalizado.bandeira;

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
            "softcase": linha["nomenclatura"],
            "bandeira": linha["bandeira"],
            "dias_comp": linha["Dia_Comp"],
            "valor_taxa": _formatar_decimal(linha["valor_taxa"]),
            "taxa_adquirente": _formatar_decimal(linha["taxa_adquirente"]),
            "forma_pagto": linha["forma_pagto"],
            "forma_pagamento": linha["forma_pagamento"],
            "valor": _formatar_decimal(linha["total"]),
            "data_ref": linha["data_ref"],
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

