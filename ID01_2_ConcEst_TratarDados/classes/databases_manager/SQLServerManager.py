# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.CredentialManager import CredentialManager
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException

import pyodbc

class SqlServerManager:
    """
    Classe modelo para realizar SELECT, UPDATE e INSERT no SQL Server.
    """

    _config = InitAllSettings.config
    
    _server = None
    _database = None
    _user = None
    _password = None

    @classmethod
    def _connect(cls) -> pyodbc.Connection:
        """
        Realiza a conexão com o SQL Server.
        Retorna:
        - pyodbc.Connection: Retorna a conexão feita com o SQLServer.
        """
        
        # Verificando se as credenciais já foram carregadas, para evitar buscar mais de uma vez
        if cls._server is None or cls._database is None or cls._user is None or cls._password is None:
            cls._server = CredentialManager.get_credential(label=cls._config["CRED_LABEL_BANCO_MANAGER"], key=cls._config["CRED_KEY_HOST_BANCO_MANAGER"])
            cls._database = CredentialManager.get_credential(label=cls._config["CRED_LABEL_BANCO_MANAGER"], key=cls._config["CRED_KEY_DATABASE_BANCO_MANAGER"])
            cls._user = CredentialManager.get_credential(label=cls._config["CRED_LABEL_BANCO_MANAGER"], key=cls._config["CRED_KEY_USER_BANCO_MANAGER"])
            cls._password = CredentialManager.get_credential(label=cls._config["CRED_LABEL_BANCO_MANAGER"], key=cls._config["CRED_KEY_PASSWORD_BANCO_MANAGER"])

        connection_string = (
            "DRIVER={SQL Server};"
            + "SERVER=" + cls._server + ";"
            + "DATABASE=" + cls._database + ";"
            + "UID=" + cls._user + ";"
            + "PWD=" + cls._password
        )

        try:
            sql_conn = pyodbc.connect(connection_string)
            return sql_conn
        except Exception as err:
            Log.write_log(f"Erro ao conectar ao SQL Server: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
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
                with sql_conexao.cursor() as csr_cursor:
                    csr_cursor.execute(query, [status])
                    name_columns = [column[0] for column in csr_cursor.description]
                    results = [dict(zip(name_columns, row)) for row in csr_cursor.fetchall()]
        except Exception as err:
            Log.write_log(f"Erro ao realizar select na tabela: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
        finally:
            if csr_cursor:
                csr_cursor.close()

            return results

    @classmethod
    def insert_item(cls, values: tuple):
        """
        Executa um INSERT no banco de dados.

        Parâmetros:
        - values (list): Valores a serem inseridos na tabela.
        """
        try:
            query = f"INSERT INTO {cls._config['TabelaNome']} (col1, col2) VALUES (?, ?)"

            with cls._connect() as sql_conexao:
                with sql_conexao.cursor() as csr_cursor:
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
                with sql_conexao.cursor() as csr_cursor:
                    csr_cursor.execute(query, values)
        except Exception as err:
            Log.write_log(f"Erro ao atualizar item: {str(err)}")
        finally:
            if csr_cursor:
                csr_cursor.close()