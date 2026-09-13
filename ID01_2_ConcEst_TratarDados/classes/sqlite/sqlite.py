"""Camada de acesso ao banco SQLite da automação ID01_2_ConcEst_TratarDados.

Concentra a criação de tabelas e a inserção de dados, para que os
demais módulos (ex.: manipular_tabelas.py) não precisem lidar com
SQL diretamente.
"""

from __future__ import annotations

import logging
import sqlite3
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


class GerenciadorSQLite:
    """Encapsula conexão, criação de tabelas e inserção de dados em SQLite.

    Uso recomendado (garante commit/rollback e fechamento da conexão):

        with GerenciadorSQLite(caminho_banco) as db:
            db.criar_tabela("tbl_empresas", {"ID_UNICO": "TEXT"})
            db.inserir_dados("tbl_empresas", linhas)
    """

    def __init__(self, caminho_banco: str | Path) -> None:
        self.caminho_banco = Path(caminho_banco)
        self._conexao: Optional[sqlite3.Connection] = None

    def __enter__(self) -> "GerenciadorSQLite":
        self.conectar()
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.fechar()

    def conectar(self) -> None:
        """Cria a pasta do banco (se necessário) e abre a conexão."""
        self.caminho_banco.parent.mkdir(parents=True, exist_ok=True)
        logger.debug("Conectando ao banco SQLite: %s", self.caminho_banco)
        self._conexao = sqlite3.connect(self.caminho_banco)
        self._conexao.execute("PRAGMA foreign_keys = ON;")

    def fechar(self) -> None:
        """Fecha a conexão com o banco, se estiver aberta."""
        if self._conexao is not None:
            self._conexao.close()
            self._conexao = None

    @property
    def conexao(self) -> sqlite3.Connection:
        if self._conexao is None:
            raise RuntimeError(
                "Conexão não está aberta. Use 'with GerenciadorSQLite(...) "
                "as db:' ou chame 'conectar()' antes de operar no banco."
            )
        return self._conexao

    def criar_tabela(
        self,
        nome_tabela: str,
        colunas: dict[str, str],
        recriar: bool = True,
    ) -> None:
        """Cria uma tabela a partir de um mapeamento coluna -> tipo SQL.

        Args:
            nome_tabela: nome da tabela a ser criada.
            colunas: dict {nome_da_coluna: tipo_sql}, ex.: {"ID_UNICO":
                "TEXT", "Dias Comp.": "INTEGER"}.
            recriar: se True (padrão), remove a tabela existente antes
                de criar novamente, mantendo o banco sincronizado com o
                DEPARA a cada execução da automação. Use False para
                preservar dados já existentes na tabela.
        """
        if not colunas:
            raise ValueError(
                f"Não é possível criar a tabela '{nome_tabela}' sem colunas."
            )

        definicao_colunas = ", ".join(
            f'"{coluna}" {tipo}' for coluna, tipo in colunas.items()
        )

        if recriar:
            logger.debug(
                "Removendo tabela existente (se houver): %s", nome_tabela
            )
            self.conexao.execute(f'DROP TABLE IF EXISTS "{nome_tabela}";')

        logger.info("Criando tabela: %s", nome_tabela)
        self.conexao.execute(
            f'CREATE TABLE IF NOT EXISTS "{nome_tabela}" ({definicao_colunas});'
        )
        self.conexao.commit()

    def inserir_dados(
        self,
        nome_tabela: str,
        dados: list[dict[str, object]],
    ) -> int:
        """Insere uma lista de dicionários em uma tabela já existente.

        Args:
            nome_tabela: nome da tabela de destino.
            dados: lista de dicionários {nome_da_coluna: valor}. As
                chaves do primeiro dicionário definem as colunas usadas
                no INSERT.

        Returns:
            Quantidade de linhas inseridas.

        Raises:
            sqlite3.Error: se a inserção falhar (a transação é desfeita
                com rollback antes de a exceção ser repropagada).
        """
        if not dados:
            logger.warning("Nenhum dado para inserir em '%s'.", nome_tabela)
            return 0

        colunas = list(dados[0].keys())
        marcadores = ", ".join("?" for _ in colunas)
        colunas_sql = ", ".join(f'"{coluna}"' for coluna in colunas)
        comando = (
            f'INSERT INTO "{nome_tabela}" ({colunas_sql}) '
            f"VALUES ({marcadores});"
        )
        valores = [
            tuple(linha.get(coluna) for coluna in colunas) for linha in dados
        ]

        try:
            self.conexao.executemany(comando, valores)
            self.conexao.commit()
        except sqlite3.Error:
            self.conexao.rollback()
            logger.exception(
                "Erro ao inserir dados na tabela '%s'.", nome_tabela
            )
            raise

        logger.info(
            "Tabela '%s': %d linha(s) inserida(s).", nome_tabela, len(valores)
        )
        return len(valores)
