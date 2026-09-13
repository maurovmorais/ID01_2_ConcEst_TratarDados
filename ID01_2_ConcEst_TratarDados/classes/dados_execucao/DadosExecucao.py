# Imports dos módulos internos do projeto
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.ExecutionControl import ExecutionControl
from ID01_2_ConcEst_TratarDados.classes.utils.LogMethod import log_metodo

# Imports dos pacotes externos
import json
from typing import Any
from datetime import datetime
from win32api import GetSystemMetrics # type: ignore
import socket
import sqlite3
import pandas as pd

"""
---------------------------------
ESTRUTURAS ESPERADA PELA CLASSE (SQLITE):
---------------------------------

- TABELA DADOS EXECUÇÃO:

CREATE TABLE <SCHEMA>.<NOME TABELA> (
    id_execucao INTEGER PRIMARY KEY AUTOINCREMENT,
    guid_execucao TEXT NOT NULL UNIQUE,
    nome_processo TEXT NOT NULL,
    nome_maquina TEXT NOT NULL,
    resolucao_tela TEXT NOT NULL,
    id_runner TEXT,
    versao_runner TEXT,
    inicio_exec TEXT NOT NULL,
    fim_exec TEXT,
    tempo_execucao TEXT,
    qtd_itens_fila INTEGER NOT NULL,
    qtd_itens_sucesso INTEGER DEFAULT 0,
    qtd_itens_business INTEGER DEFAULT 0,
    qtd_itens_app INTEGER DEFAULT 0,
    excecao_inicializacao TEXT,
    tipo_excecao_inicializacao TEXT,
    screenshot_excecao TEXT,
    ultima_atualizacao DATETIME NOT NULL
);

-- TABELA DADOS ITENS FILA:

CREATE TABLE <SCHEMA>.<NOME TABELA> (
    id_item INTEGER PRIMARY KEY AUTOINCREMENT,
    id_execucao INTEGER NOT NULL,
    id_item_fila TEXT NOT NULL,
    data_hora_inicio TEXT NOT NULL,
    data_hora_fim TEXT,
    tempo_execucao TEXT,
    nome_fila TEXT NOT NULL,
    referencia TEXT,
    detalhes_item_fila TEXT,
    status TEXT,
    tipo_excecao TEXT,
    descricao_excecao TEXT,
    screenshot_excecao TEXT,
    ultima_atualizacao DATETIME NOT NULL,
    FOREIGN KEY (id_execucao) REFERENCES <SCHEMA>.<NOME TABELA EXECUCAO>(id_execucao)
);
""" 

@log_metodo
class DadosExecucao:
    """
    Classe responsável por manipular os registros de execução no banco de dados SQLite, que geram os relatórios analítico e sintético.

    Parâmetros:

    Retorna:
    """
    # Variaveis que podem ser acessadas externamente ao importar a classe em outros modulos
    _config = InitAllSettings.config
    _guid_execucao = InitAllSettings.guid_execucao
    id_registro_execucao = None
    id_registro_item = None
    _tabela_dados_execucao = _config["NomeTabelaDadosExecucao"]
    _tabela_dados_itens = _config["NomeTabelaDadosItens"]

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
            db_path = InitAllSettings.sqlite_caminho_bd_analit_sint
            sql_conn = sqlite3.connect(db_path)

            # Habilitar o suporte a chaves estrangeiras
            sql_conn.execute("PRAGMA foreign_keys = ON;")

            return sql_conn
        
        except Exception as err:
            Log.write_log(f"Erro ao conectar ao SQLite: {str(err)}")
            return None

    @classmethod
    def inserir_tabela_dados_execucao(cls, itens_fila:int):
        """
        Insere os dados de execução na tabela de itens.
        
        Parâmetros:
        - itens_fila (int): Número de itens na fila de execução.
        
        Retorna:
        """
        try:
            #Preparando os valores para inserir no banco
            guid_execucao = cls._guid_execucao
            nome_processo = cls._config["NomeProcesso"]
            nome_maquina = socket.gethostname()
            resolucao_tela = f"{GetSystemMetrics(0)}x{GetSystemMetrics(1)}"
            id_runner = ExecutionControl.runner_id
            versao_runner = ExecutionControl.versao_runner
            inicio_exec = InitAllSettings.datahora_inicio_exec
            total_itens_fila = itens_fila
            ultima_atualizacao = datetime.now()

            values = [guid_execucao, nome_processo, nome_maquina, resolucao_tela, id_runner,
                              versao_runner, inicio_exec, total_itens_fila, ultima_atualizacao]
            
            columns = "GUID_EXECUCAO, NOME_PROCESSO, NOME_MAQUINA, RESOLUCAO_TELA, ID_RUNNER, VERSAO_RUNNER, \
                               INICIO_EXEC, QTD_ITENS_FILA, ULTIMA_ATUALIZACAO"

            #Construindo o comando insert com placeholders '?' para SQLite
            insert = f"INSERT INTO {cls._tabela_dados_execucao} ({columns}) VALUES (?,?,?,?,?,?,?,?,?)"

            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                csr_cursor.execute(insert, values)
                sql_conn.commit()

                # Captura o id do último registro inserido
                cls.id_registro_execucao = csr_cursor.lastrowid
                    

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao inserir linhas na tabela dados execução do banco SQLite: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            return None

        finally:    
            #Fechando conexao    
            if csr_cursor is not None:
                csr_cursor.close()

    @classmethod
    def update_tabela_dados_execucao_itens_new(cls, itens_fila:int):
        """
        Atualiza a tabela de dados da execucao para definir a quantidade de itens na fila.
        
        Parâmetros:
        - itens_fila (int): Número de itens na fila de execução.
        
        Retorna:
        """
        try:
            #Preparando os valores para realizar update no banco
            id_registro_execucao = str(cls.id_registro_execucao)
            ultima_atualizacao = datetime.now()

            values_update = [itens_fila, ultima_atualizacao, id_registro_execucao]

            #Construindo o comando Update com placeholders '?'
            update = f"UPDATE {cls._tabela_dados_execucao} SET QTD_ITENS_FILA = ?, ULTIMA_ATUALIZACAO = ? WHERE ID_EXECUCAO = ?"

            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor() 
                csr_cursor.execute(update, values_update)
                sql_conn.commit()

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao atualizar quantidade de itens na tabela dados execução: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)

        finally:    
            #Fechando conexao    
            if csr_cursor is not None:
                csr_cursor.close()

    @classmethod
    def update_tabela_dados_execucao(cls, msg_erro_inicializacao:str=None, tipo_erro_inicializacao:str=None, screenshot_excecao:str=None):
        """
        Atualiza as informações da tabela de dados da execucao para inserir informações como: qtd_itens_sucesso, qtd_itens_business, qtd_itens_app, fim_exec.
        
        Parâmetros:
        - msg_erro_inicializacao (str): Mensagem de erro da inicialização caso aconteça.
        - tipo_erro_inicializacao (str): Tipo do erro da inicialização caso aconteça.
        - screenshot_excecao (str): Screenshot do erro da inicialização caso aconteça.
        
        Retorna:
        """
        try:
            #Preparando os valores para inserir no banco
            id_registro_execucao = str(cls.id_registro_execucao)
            data_hora_fim_execucao = datetime.now() if str(InitAllSettings.datahora_fim_exec) == '' else InitAllSettings.datahora_fim_exec
            ultima_atualizacao = datetime.now()

            # Atualiza a contagem dos itens da execucao
            cls.refresh_counting_items()  

            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()

                csr_cursor.execute(f"SELECT inicio_exec FROM {cls._tabela_dados_execucao} WHERE id_execucao = ? LIMIT 1;", [id_registro_execucao])
                data_hora_inicio_execucao = csr_cursor.fetchone()[0]

                tempo_exec_formatado = cls.formatar_tempo_execucao(data_hora_inicio_execucao, data_hora_fim_execucao)

                #Script Update Dados Finais Execucao
                with open(InitAllSettings.caminho_script_update_dados_execucao) as script_file:
                    script_update_dados_execucao = script_file.read()

                script_update_dados_execucao = script_update_dados_execucao.replace("<tabela_dados_execucao>",f"{cls._tabela_dados_execucao}")

                #Executando o comando update
                values_update = [data_hora_fim_execucao,InitAllSettings.qtde_itens_sucesso,
                                        InitAllSettings.qtde_itens_business_exception,InitAllSettings.qtde_itens_app_exception,
                                        tempo_exec_formatado,msg_erro_inicializacao,
                                        tipo_erro_inicializacao,screenshot_excecao,
                                        id_registro_execucao]
                
                #Executando o comando update
                csr_cursor.execute(script_update_dados_execucao, values_update)
                sql_conn.commit()

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao atualizar linhas na tabela dados execução do banco SQLite: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
           
        finally:    
            #Fechando conexao    
            if csr_cursor is not None:
                csr_cursor.close()   
    
    @classmethod
    def refresh_counting_items(cls):
        """
        Atualiza contagem de itens no InitAllSettings.
        
        Parâmetros:
        
        Retorna:
        """
        try:
            id_registro_execucao = str(cls.id_registro_execucao)

            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor() 
                
                with open(InitAllSettings.caminho_script_select_captura_qtd_itens) as script_file:
                    script_update_select_captura_qtd_itens = script_file.read()

                script_update_select_captura_qtd_itens = script_update_select_captura_qtd_itens.replace("<tabela_dados_itens>",f"{cls._tabela_dados_itens}")
                
                values_select = [id_registro_execucao]

                #Executando o comando select
                csr_cursor.execute(script_update_select_captura_qtd_itens,values_select)
                name_columns = list(map(lambda x: str(x[0]).lower(), csr_cursor.description))
                linha_qtd_itens = csr_cursor.fetchone()

                #O fetchall é para prevenir do erro Unread result found.
                csr_cursor.fetchall()

                dados_qtd_itens:dict = dict(zip(name_columns, linha_qtd_itens))

                InitAllSettings.qtde_itens_sucesso = dados_qtd_itens['qtd_itens_sucesso']
                InitAllSettings.qtde_itens_business_exception = dados_qtd_itens['qtd_itens_business']
                InitAllSettings.qtde_itens_app_exception = dados_qtd_itens['qtd_itens_app']
                InitAllSettings.qtde_itens_processados = sum(linha_qtd_itens)

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao capturar quantidade de itens processados na tabela dados execução do banco SQLite: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)

        finally:    
            #Fechando conexao    
            if csr_cursor is not None:
                csr_cursor.close()

    @classmethod
    def inserir_tabela_dados_itens(cls, id_item_fila:str, nome_fila:str, referencia:str, detalhes_item_fila:str):
        """
        Insere os dados de itens na tabela de itens.
        
        Parâmetros:
        - id_item_fila (str): ID do item na fila.
        - nome_fila (str): Nome da fila.
        - referencia (str): Referência do item.
        - detalhes_item_fila (str): Detalhes adicionais do item.
        
        Retorna:
        """
        try:
            id_execucao = cls.id_registro_execucao
            data_hora_inicio = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            ultima_atualizacao = datetime.now()
            
            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                # Insira um novo registro
                values_insert = [id_execucao, id_item_fila, data_hora_inicio, nome_fila, referencia, detalhes_item_fila, ultima_atualizacao]
                columns = "ID_EXECUCAO, ID_ITEM_FILA, DATA_HORA_INICIO, NOME_FILA, REFERENCIA, DETALHES_ITEM_FILA, ULTIMA_ATUALIZACAO"
                insert = f"INSERT INTO {cls._tabela_dados_itens} ({columns}) VALUES (?,?,?,?,?,?,?)"
                
                csr_cursor.execute(insert, values_insert)
                sql_conn.commit()

                cls.id_registro_item = csr_cursor.lastrowid

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao inserir linhas na tabela dados itens do banco SQLite: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)

        finally:    
            #Fechando conexao    
            if csr_cursor is not None:
                csr_cursor.close()

    @classmethod
    def update_tabela_dados_itens(cls, tipo_status:str, tipo_excecao:str=None, descricao_erro:str=None, screenshot_excecao:str=None):
        """
        Atualiza os dados de itens na tabela de dados itens.
        
        Parâmetros:
        - tipo_status (str): Status do item (SUCESSO/FALHA).
        - tipo_excecao (str): Tipo de exceção, se houver (NEGOCIO/SISTEMA).
        - descricao_erro (str): Descrição do erro, se houver.
        - screenshot_excecao (str): Caminho do screenshot do erro, se houver.
        
        Retorna:
        """
        try:
            id_registro_item = cls.id_registro_item
            data_hora_fim = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            ultima_atualizacao = datetime.now()

            match tipo_status.upper():
                case 'SUCESSO':
                    #Preparando os valores para realizar update no banco
                    status = 'SUCESSO'
                    tipo_excecao_final = ''
                    descricao_excecao = ''

                case 'FALHA':
                    #Preparando os valores para realizar update no banco
                    status = 'FALHA'
                    tipo_excecao_final = tipo_excecao
                    descricao_excecao = descricao_erro

                case default:
                    None

            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                csr_cursor.execute(f"SELECT DATA_HORA_INICIO FROM {cls._tabela_dados_itens} WHERE ID_ITEM = ?", [id_registro_item])
                data_inicio_str = csr_cursor.fetchone()[0]
                
                tempo_exec_formatado = cls.formatar_tempo_execucao(data_inicio_str, data_hora_fim)
                
                #Construindo o comando Update
                values_update = [data_hora_fim, status, tipo_excecao_final, descricao_excecao, 
                                        screenshot_excecao, ultima_atualizacao, tempo_exec_formatado, id_registro_item]
                
                update = f"UPDATE {cls._tabela_dados_itens} SET DATA_HORA_FIM = ?, STATUS = ?, TIPO_EXCECAO = ?, DESCRICAO_EXCECAO = ?, SCREENSHOT_EXCECAO = ?, ULTIMA_ATUALIZACAO = ?, tempo_execucao = ? WHERE ID_ITEM = ?"
                
                #Executando o comando update
                csr_cursor.execute(update, values_update)
                sql_conn.commit()

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao atualizar linhas na tabela dados itens do banco SQLite: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)

        finally:    
            #Fechando conexao    
            if csr_cursor is not None:
                csr_cursor.close()

    @classmethod
    def update_add_details_queue_item(cls, detalhes_item_add:dict=None):
        """
        Atualiza os detalhes dos itens na tabela dados itens.
        
        Parâmetros:
        - detalhes_item_add (dict): Dicionário com detalhes adicionais a serem adicionados.
        
        Retorna:
        """
        try:
            id_registro_item = cls.id_registro_item
            ultima_atualizacao = datetime.now()
            dados_capturados_fila = None
            
            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                csr_cursor.execute(f"SELECT DETALHES_ITEM_FILA FROM {cls._tabela_dados_itens} WHERE ID_ITEM = ?", [id_registro_item])
                
                name_columns = list(map(lambda x: str(x[0]).lower(), csr_cursor.description))
                linha_capturada_fila = csr_cursor.fetchone()
                dados_capturados_fila = None

                if(linha_capturada_fila is not None):
                    dados_capturados_fila:dict = dict(zip(name_columns, linha_capturada_fila))
                    
                    detalhes_item_fila = json.loads(json.dumps(dados_capturados_fila['detalhes_item_fila'],ensure_ascii=False))

                    if(isinstance(detalhes_item_fila,str)):
                        detalhes_item_fila = json.loads(detalhes_item_fila)

                    dados_capturados_fila['detalhes_item_fila'] = detalhes_item_fila
                else:
                    dados_capturados_fila['detalhes_item_fila'] = {}

                dados_capturados_fila['detalhes_item_fila'] = {**dados_capturados_fila['detalhes_item_fila'],**detalhes_item_add}
                
                values_update = [json.dumps(dados_capturados_fila['detalhes_item_fila'],ensure_ascii=False),ultima_atualizacao, id_registro_item]
                        
                #Construindo o comando Update
                update = f"UPDATE {cls._tabela_dados_itens} SET DETALHES_ITEM_FILA = ?, ULTIMA_ATUALIZACAO = ? WHERE ID_ITEM = ?"

                #Executando o comando update
                csr_cursor.execute(update, values_update)
                sql_conn.commit()

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao adicionar detalhes de item do registro da execução no banco SQLite: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)

        #Fechando conexao    
            if csr_cursor is not None:
                csr_cursor.close()

    @classmethod
    def update_change_value_details_queue_item(cls, key:str=None, any_value:Any=None):
        """
        Atualiza um valor específico nos detalhes dos itens na tabela de itens da fila.
        
        Parâmetros:
        - key (str): Chave a ser atualizada.
        - any_value (Any): Novo valor para a chave.
        
        Retorna:
        """
        try:
            id_registro_item = cls.id_registro_item
            ultima_atualizacao = datetime.now()

            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                csr_cursor.execute(f"SELECT DETALHES_ITEM_FILA FROM {cls._tabela_dados_itens} WHERE ID_ITEM = ?", [id_registro_item])
                        
                name_columns = list(map(lambda x: str(x[0]).lower(), csr_cursor.description))
                linha_capturada_fila = csr_cursor.fetchone()
                dados_capturados_fila = None

                if(linha_capturada_fila is not None):
                    dados_capturados_fila:dict = dict(zip(name_columns, linha_capturada_fila))
                    detalhes_item_fila = json.loads(json.dumps(dados_capturados_fila['detalhes_item_fila'],ensure_ascii=False))

                    if(isinstance(detalhes_item_fila,str)):
                        detalhes_item_fila = json.loads(detalhes_item_fila)

                    dados_capturados_fila['detalhes_item_fila'] = detalhes_item_fila
                else:
                    dados_capturados_fila['detalhes_item_fila'] = {}

                dados_capturados_fila['detalhes_item_fila'][key] = any_value
                    
                values_update = [json.dumps(dados_capturados_fila['detalhes_item_fila'],ensure_ascii=False),ultima_atualizacao, id_registro_item]
                        
                #Construindo o comando Update
                update = f"UPDATE {cls._tabela_dados_itens} SET DETALHES_ITEM_FILA = ?, ULTIMA_ATUALIZACAO = ? WHERE ID_ITEM = ?"

                #Executando o comando update
                csr_cursor.execute(update, values_update)
                sql_conn.commit()

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao alterar detalhes de item do registro da execução no banco SQLite: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)

        finally:    
            #Fechando conexao    
            if csr_cursor is not None:
                csr_cursor.close()

    @classmethod
    def formatar_tempo_execucao(cls, data_inicio:str, data_fim:str) -> str:
        """
        Calcula o tempo de execução entre duas datas e formata no padrão HH:MM:SS.
        
        Parâmetros:
        - data_inicio (str): Data/Hora de início da execução.
        - data_fim (str): Data/Hora de fim da execução.
        
        Retorna:
        - tempo_exec (str): Tempo de execução formatado.
        """
        try:
            #Convertendo strings das datas para o tipo Datetime
            inicio_exec = datetime.strptime(data_inicio, "%d/%m/%Y %H:%M:%S")
            fim_exec = datetime.strptime(data_fim, "%d/%m/%Y %H:%M:%S")

            diff_date = fim_exec - inicio_exec
            diff_days = str(diff_date.days)
            diff_hours = "{:02d}".format(diff_date.seconds // 3600)
            diff_minutes = "{:02d}".format((diff_date.seconds % 3600) // 60)
            diff_seconds = "{:02d}".format(diff_date.seconds % 60)

            return (diff_days + ' ' + diff_hours + ':' + diff_minutes + ':' + diff_seconds)
            
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao formatar data para realizar cálculo de horas: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)


    @classmethod
    def get_dados_analitico(cls) -> dict:
        """
        Captura os dados com detalhes sobre a execução dos itens para a montagem do relatório analítico.
        
        Parâmetros:
           
        Retorna:
           - dict: dicionario com as colunas e os dados dos itens da execucao em lista
        """
        try:
            col_dados_analitico = {}

            #Script para realizar a captura dos dados para relatorio analitico
            with open(InitAllSettings.caminho_script_select_dados_analitico) as script_file:
                script_captura_dados_analitico = script_file.read()
            
            script_captura_dados_analitico = script_captura_dados_analitico.replace("<tabela_dados_execucao>",f"{cls._tabela_dados_execucao}")
            script_captura_dados_analitico = script_captura_dados_analitico.replace("<tabela_dados_itens>",f"{cls._tabela_dados_itens}")
            
            # Realizando a conexão do banco
            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                # Executando o comando select
                csr_cursor.execute(script_captura_dados_analitico,[cls._guid_execucao])
                name_columns = list(map(lambda x: x[0], csr_cursor.description))
                dados_analitico = csr_cursor.fetchall()
            
            col_dados_analitico = {"Colunas": name_columns, "Dados": dados_analitico}
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao capturar dados analíticos: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
        finally:
            #Fechando conexao    
            if csr_cursor is not None:
                csr_cursor.close()

            return col_dados_analitico
    
    @classmethod
    def get_dados_sintetico(cls) -> dict:
        """
        Captura os dados gerais da execução para a montagem do relatório sintético.
        
        Parâmetros:
           
        Retorna:
           - dict: dicionario com as colunas e os dados da execucao em lista
        """
        try:
            col_dados_sintetico = {}

            #Script para realizar a captura dos dados para relatorio sintetico
            with open(InitAllSettings.caminho_script_select_dados_sintetico) as script_file:
                script_captura_dados_sintetico = script_file.read()
            
            script_captura_dados_sintetico = script_captura_dados_sintetico.replace("<tabela_dados_execucao>",f"{cls._tabela_dados_execucao}")
            script_captura_dados_sintetico = script_captura_dados_sintetico.replace("<tabela_dados_itens>",f"{cls._tabela_dados_itens}")

            # Realizando a conexão do banco
            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                # Executando o comando select
                csr_cursor.execute(script_captura_dados_sintetico,[cls._guid_execucao])
                name_columns = list(map(lambda x: x[0], csr_cursor.description))
                dados_sintetico = csr_cursor.fetchall()
                    
            
            col_dados_sintetico = {"Colunas": name_columns, "Dados": dados_sintetico}
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao capturar dados sintéticos: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
        finally:
            #Fechando conexao    
            if csr_cursor is not None:
                csr_cursor.close()
                
            return col_dados_sintetico
        
    @classmethod
    def get_dados_raas_analitico(cls) -> dict:
        """
        Captura os dados com detalhes sobre a execução dos itens para a montagem do relatório raas analítico.
        
        Parâmetros:
           
        Retorna:
           - dict: dicionario com as colunas e os dados dos itens da execucao em lista
        """
        try:
            col_dados_analitico = {}

            query_select_analitico_raas = f'SELECT nome_fila, referencia, detalhes_item_fila, data_hora_inicio, \
                                                data_hora_fim, status, tipo_excecao, \
                                                descricao_excecao FROM {cls._tabela_dados_itens} WHERE id_execucao = ( SELECT id_execucao FROM {cls._tabela_dados_execucao} \
                                                WHERE guid_execucao = ? )'
            
            # Realizando a conexão do banco
            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                # Executando o comando select
                csr_cursor.execute(query_select_analitico_raas,[cls._guid_execucao])
                name_columns = list(map(lambda x: x[0], csr_cursor.description))
                dados_analitico = csr_cursor.fetchall()
            
            col_dados_analitico = {"Colunas": name_columns, "Dados": dados_analitico}
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao capturar dados analíticos para RAAS: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
        finally:
            #Fechando conexao    
            if csr_cursor is not None:
                csr_cursor.close()

            return col_dados_analitico
    
    @classmethod
    def get_dados_raas_sintetico(cls) -> dict:
        """
        Captura os dados com detalhes sobre a execução para a montagem do relatório raas sintetico.
        
        Parâmetros:
           
        Retorna:
           - dict: dicionario com as colunas e os dados da execucao em lista
        """
        try:
            col_dados_sintetico = {}
            
            query_select_sintetico_raas = f'SELECT * FROM {cls._tabela_dados_execucao} WHERE guid_execucao = ?'
            
            # Realizando a conexão do banco
            with cls._connect() as sql_conn:
                csr_cursor = sql_conn.cursor()
                # Executando o comando select
                csr_cursor.execute(query_select_sintetico_raas,[cls._guid_execucao])
                name_columns = list(map(lambda x: x[0], csr_cursor.description))
                dados_sintetico = csr_cursor.fetchone()
            
            col_dados_sintetico = {"Colunas": name_columns, "Dados": dados_sintetico}
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao capturar dados sintéticos para RAAS: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
        finally:
            #Fechando conexao    
            if csr_cursor is not None:
                csr_cursor.close()

            return col_dados_sintetico