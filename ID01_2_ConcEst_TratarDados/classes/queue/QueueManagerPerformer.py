# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.ExecutionControl import ExecutionControl
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException

# Imports dos pacotes externos
import sqlite3
import json
from datetime import datetime
from pathlib import Path
from os import path
from typing import Any

"""
ESTRUTURA ESPERADA POR ESSA CLASSE:

Create table tbl_Fila_Processamento(
    id integer primary key,
    referencia varchar(200),
    datahora_criado varchar(50),
    nome_maquina varchar(200),
    info_adicionais text,
    status varchar(100),
    obs varchar(500),
    ultima_atualizacao datetime);


O ARQUIVO DO BANCO LOCALIZADO EM resources/sqlite/banco_dados.db JÁ POSSUI ESSA TABELA CRIADA E VAZIA
"""
class QueueManagerPerformer:
    """
    Classe responsável para manipulação do sqlite e controle de fila.
    
    Parâmetros:

    Retorna:
    """

    _config = InitAllSettings.config
    tabela_fila = _config["FilaProcessamentoPerformer"]
    path_to_db = InitAllSettings.sqlite_caminho_bd_analit_sint

    items_queue = 0
    _item_atual = None
    nome_maquina = ExecutionControl.runner_id

    @classmethod
    def _connect(cls):
        """
        Realiza a conexão com o banco de dados SQLite.

        Parâmetros:
        
        Retorna:
        - sqlite3.Connection: Retorna a conexão feita com o SQLite.
        """
        try:
            # O caminho do arquivo do banco de dados é obtido do config
            db_path = cls.path_to_db
            sql_conn = sqlite3.connect(db_path)

            # Habilitar o suporte a chaves estrangeiras
            sql_conn.execute("PRAGMA foreign_keys = ON;")

            return sql_conn
        
        except Exception as err:
            Log.write_log(f"Erro ao conectar ao SQLite: {str(err)}")
            return None

    @classmethod
    def refresh_total_items_new(cls):
        """
        Atualiza a própria classe, usado em vários pontos do projeto para atualizar a quantidade de itens na fila como new
        
        Parâmetros:

        Retorna:
        """
        try:

            with cls._connect() as sql_conn:
                csr_cursor= sql_conn.cursor()
                query = f"SELECT * FROM {cls.tabela_fila} WHERE status = 'NEW' OR status = 'ON QUEUE'"
                csr_cursor.execute(query)
                cls.items_queue = len(csr_cursor.fetchall())

        except Exception as err:
            Log.write_log(
                mensagem_log="Erro ao atualizar total de itens na fila: " + str(err),
                log_level=LogLevel.ERROR,
                error_type=ErrorType.APP_ERROR
            )
        finally:
            if csr_cursor:
                csr_cursor.close()

            

    @classmethod
    def insert_new_queue_item(cls, referencia: str, inf_adicional: dict = None):
        """
        Insere um item na tabela especificada, com a referência e com os valores adicionais
                
        Parâmetros:
        - referencia (str): referência do item.
        - inf_adicional (dict): informações adicionais (opcional, default=None).

        Retorna:
        """
        try:

            agora: datetime = datetime.now()
            agora_str = agora.strftime("%d/%m/%Y %H:%M:%S")
            nome_maquina = ''
            values = [
                referencia, agora_str, nome_maquina,
                json.dumps(inf_adicional, ensure_ascii=False), "NEW", "", agora
            ]
            insert = (
                f"INSERT INTO {cls.tabela_fila} (referencia, datahora_criado, nome_maquina, info_adicionais, status, obs, ultima_atualizacao) "
                "VALUES (?,?,?,?,?,?,?)"
            )
            
            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                csr_cursor.execute(insert, values)
                sql_conn.commit()

        except Exception as err:
            Log.write_log(
                mensagem_log="Erro ao inserir linhas: " + str(err),
                log_level=LogLevel.ERROR,
                error_type=ErrorType.APP_ERROR
            )
            raise err
        finally:
            if csr_cursor:
                csr_cursor.close()

            cls.refresh_total_items_new()

    @classmethod
    def get_specific_queue_item(cls, id: str) -> dict:
        """
        Retorna um item específico da fila.

        Parâmetros:
        - id (str): índice do item.

        Retorna:
        - dict: retorna um dicionario com os dados do item da fila
        """
        try:
            dados_capturados_fila = None

            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                query = f"SELECT * FROM {cls.tabela_fila} WHERE id = ?"
            
                csr_cursor.execute(query, [int(id)])
                    
                name_columns = list(map(lambda x: x[0], csr_cursor.description))
                linha_capturada_fila = csr_cursor.fetchone()
                
                
                if linha_capturada_fila is not None:
                    dados_capturados_fila = dict(zip(name_columns, linha_capturada_fila))
                    dados_capturados_fila['info_adicionais'] = json.loads(str(dados_capturados_fila['info_adicionais']))
        
        except Exception as err:
            Log.write_log(
                mensagem_log="Erro ao buscar item específico da fila: " + str(err),
                log_level=LogLevel.ERROR,
                error_type=ErrorType.APP_ERROR
            )
            raise err
        finally:
            if csr_cursor:
                csr_cursor.close()

            cls.refresh_total_items_new()
            cls._item_atual = dados_capturados_fila

            return dados_capturados_fila


    @classmethod
    def get_next_queue_item(cls) -> dict:
        """
        Retorna o próximo item da fila que não foi processado e não possui máquina alocada, None se não existe nenhum item assim.

        Parâmetros:

        Retorna:
        - dict: dicionario com as informações do próximo item da fila.
        """
        try:
            agora = datetime.now()
            dados_capturados_fila = None

            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                csr_cursor.execute(f"UPDATE {cls.tabela_fila} SET \
                                                                    ultima_atualizacao = ?, \
                                                                    nome_maquina = ?, \
                                                                    status = 'ON QUEUE' \
                                                                    WHERE id = (SELECT MIN(id) FROM {cls.tabela_fila} WHERE status = 'NEW')",
                                                                    [agora,
                                                                    InitAllSettings.guid_execucao])
                sql_conn.commit()

                csr_cursor = sqlite3.connect(cls.path_to_db)
                csr_cursor = csr_cursor.execute(f"SELECT * FROM {cls.tabela_fila} WHERE \
                                                                            nome_maquina = ? \
                                                                            and status = 'ON QUEUE' \
                                                                            ORDER BY id LIMIT 1",
                                                                            [InitAllSettings.guid_execucao])
                csr_cursor.connection.commit()

                name_columns = list(map(lambda x: x[0], csr_cursor.description))


                linha_capturada_fila:list = csr_cursor.fetchone()
                

                if(linha_capturada_fila is not None):
                    dados_capturados_fila:dict = dict(zip(name_columns, linha_capturada_fila))
                    dados_capturados_fila['info_adicionais'] = json.loads(str(dados_capturados_fila['info_adicionais']))

                    csr_cursor.execute(f"UPDATE {cls.tabela_fila} SET \
                                        status = 'RUNNING' \
                                        WHERE id = ?",
                                        [dados_capturados_fila['id']]).connection.commit()
                    
        except Exception as err:
            Log.write_log(
                mensagem_log="Erro ao buscar próximo item da fila: " + str(err),
                log_level=LogLevel.ERROR,
                error_type=ErrorType.APP_ERROR
            )
            raise err
        finally:
            if csr_cursor:
                csr_cursor.close()

            cls.refresh_total_items_new()
            cls._item_atual = dados_capturados_fila
            
            return dados_capturados_fila

    @classmethod
    def update_status_item(cls, excecao: Any = None, obs: str = ""):
        """
        Atualiza o status do item da fila que está sendo processado de acordo com o TIPO de erro que é enviado.

        Parâmetros:
        - excecao (Any): Argumento que contem o objeto da exceção ocorrida.
        - obs (str): observação para inserir no item da fila. (opcional, default= "").

        Retorna:
        """
        try:
            if excecao is None:
                # Caso venha None Considera Sucesso
                novo_status = 'SUCESSO'
            elif type(excecao) == BusinessRuleException:
                # Considera como Erro de Negócio
                novo_status = 'BUSINESS ERROR'
            else:
                # Qualquer outro tipo considera como Excecao
                novo_status = 'APP ERROR'


            #Tratando os casos onde obs vem com quotes, trocando por *
            obs = obs.replace('"', '*').replace("'", '*')
            agora = datetime.now() 
            agora_str = agora.strftime("%d/%m/%Y %H:%M:%S")
            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                csr_cursor.execute(f"UPDATE {cls.tabela_fila} SET \
                                                                            ultima_atualizacao = ?, \
                                                                            status = ?, \
                                                                            obs = ? \
                                                                            WHERE id = ?" ,
                                                                            [agora, 
                                                                            novo_status,
                                                                            obs,
                                                                            int(cls._item_atual['id'])])
                sql_conn.commit()
        except Exception as err:
            Log.write_log(
                mensagem_log="Erro ao atualizar status do item da fila: " + str(err),
                log_level=LogLevel.ERROR,
                error_type=ErrorType.APP_ERROR
            )
            raise err
        finally:
            if csr_cursor:
                csr_cursor.close()
        
            cls.refresh_total_items_new()

    @classmethod
    def abandon_queue(cls):
        """
        Marca todos os itens com status NEW como ABANDONED.
        
        Parâmetros:

        Retorna:
        """
        try:
            agora = datetime.now() 

            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()

                csr_cursor.execute(f"UPDATE {cls.tabela_fila} SET \
                                                                        status = 'ABANDONED', \
                                                                        ultima_atualizacao = ? \
                                                                        WHERE status = 'NEW'",[agora])
                
                sql_conn.commit()
        except Exception as err:
            Log.write_log(
                mensagem_log="Erro ao abandonar itens da fila: " + str(err),
                log_level=LogLevel.ERROR,
                error_type=ErrorType.APP_ERROR
            )
            raise err
        finally:
            if csr_cursor:
                csr_cursor.close()
        
            cls.refresh_total_items_new()

QueueManagerPerformer.refresh_total_items_new()