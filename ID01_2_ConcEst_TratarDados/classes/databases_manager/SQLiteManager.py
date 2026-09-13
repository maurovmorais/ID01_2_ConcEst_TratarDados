import sqlite3
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException

class SqliteManager:
    """
    Classe modelo para realizar SELECT, UPDATE e INSERT no SQLite.
    """

    _config = InitAllSettings.config

    @classmethod
    def _connect(cls) -> sqlite3.Connection:
        """
        Realiza a conexão com o SQLite.
        Retorna:
        - sqlite3.Connection: Conexão com o SQLite.
        """
        try:
            db_path = cls._config["SqliteDbPath"]
            sql_conn = sqlite3.connect(db_path)
            return sql_conn
        except Exception as err:
            Log.write_log(f"Erro ao conectar ao SQLite: {str(err)}")
            return None

    @classmethod
    def select_items(cls, status: str) -> list:
        """
        Executa um SELECT no banco de dados para itens com o status especificado.

        Parâmetros:
        - status (str): Status dos itens a serem selecionados.

        Retorna:
        - list: Lista de dicionários com os resultados.
        """
        
        try:
            results = []

            query = f"SELECT * FROM {cls._config['TabelaNome']} WHERE status = ?"
        
            with cls._connect() as sql_conexao:
                csr_cursor = sql_conexao.cursor()

                csr_cursor.execute(query, [status])
                name_columns = [column[0] for column in csr_cursor.description]
                results = [dict(zip(name_columns, row)) for row in csr_cursor.fetchall()]
        
        except Exception as err:
            Log.write_log(f"Erro ao realizar select na tabela: {str(err)}")
        finally:
            if csr_cursor:
                csr_cursor.close()
            
            return results

    @classmethod
    def insert_item(cls, values: list):
        """
        Executa um INSERT no banco de dados.

        Parâmetros:
        - values (list): Valores a serem inseridos na tabela.
        """
        try:
            query = f"INSERT INTO {cls._config['TabelaNome']} (col1, col2) VALUES (?, ?)"

            with cls._connect() as sql_conexao:
                csr_cursor = sql_conexao.cursor()

                csr_cursor.execute(query, values)
                sql_conexao.commit()
                    
        except Exception as err:
            Log.write_log(f"Erro ao inserir item: {str(err)}")
        finally:
            if csr_cursor:
                csr_cursor.close()

    @classmethod
    def update_item(cls, values: list):
        """
        Executa um UPDATE no banco de dados.

        Parâmetros:
        - values (list): Valores para atualização, incluindo o novo status e o ID do item.
        """
        try:
            query = f"UPDATE {cls._config['TabelaNome']} SET status = ? WHERE id = ?"

            with cls._connect() as sql_conexao:
                csr_cursor = sql_conexao.cursor()
                csr_cursor.execute(query, values)
                
                sql_conexao.commit()

        except Exception as err:
            Log.write_log(f"Erro ao atualizar item: {str(err)}")
        finally:
            if csr_cursor:
                csr_cursor.close()