"""Gravação da tbl_aux_dados no banco SQLite.

Este módulo só grava dados já prontos (um DataFrame no layout da
tbl_aux_dados) — não lê nem trata arquivo nenhum. A leitura e
normalização dos relatórios ficam a cargo do módulo
``classes/excel/leitura_relat_adquirentes.py``.

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
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings

import logging
import sqlite3
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------

CAMINHO_BANCO_PADRAO = Path(InitAllSettings.config['CaminhoBancoSqlite'])

NOME_TABELA_DESTINO = "tbl_aux_dados"


class GravadorTblAuxDados:
    """Grava um DataFrame consolidado na tbl_aux_dados (SQLite).

    Attributes:
        caminho_banco: caminho do arquivo banco_dados.db (SQLite).
    """

    def __init__(self, caminho_banco: Path = CAMINHO_BANCO_PADRAO) -> None:
        self.caminho_banco = caminho_banco

    def gravar_tbl_aux_dados(self, dados: pd.DataFrame) -> None:
        """Esvazia a tbl_aux_dados e grava os dados consolidados.

        A tabela é sempre esvaziada (``DELETE FROM``) antes da nova
        carga, dentro da mesma transação — em caso de erro, nada é
        alterado (rollback).

        Args:
            dados: DataFrame já no layout da tbl_aux_dados.
        """
        with sqlite3.connect(self.caminho_banco) as conexao:
            try:
                conexao.execute(f"DELETE FROM {NOME_TABELA_DESTINO}")
                if not dados.empty:
                    dados.to_sql(
                        NOME_TABELA_DESTINO,
                        conexao,
                        if_exists="append",
                        index=False,
                    )
                conexao.commit()
            except sqlite3.DatabaseError:
                conexao.rollback()
                logger.exception(
                    "Erro ao gravar dados na %s; transação revertida.",
                    NOME_TABELA_DESTINO,
                )
                raise

        logger.info(
            "Carga da %s concluída com %d registro(s).",
            NOME_TABELA_DESTINO,
            len(dados),
        )
