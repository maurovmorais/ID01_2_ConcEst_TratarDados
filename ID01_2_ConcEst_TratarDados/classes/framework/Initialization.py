# Imports dos módulos internos do projeto
# Carrega o InitAllSettings - Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings

import ID01_2_ConcEst_TratarDados.classes.utils.GenericReusable as GenericReusable
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllApplications import InitAllApplications
from ID01_2_ConcEst_TratarDados.classes.framework.KillAllProcesses import KillAllProcesses
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.ExecutionControl import ExecutionControl
from ID01_2_ConcEst_TratarDados.classes.utils.ScreenRecorder import ScreenRecorder
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException
from ID01_2_ConcEst_TratarDados.classes.utils.RobotStream import RobotStream
from ID01_2_ConcEst_TratarDados.classes.queue.QueueManager import QueueManager
from ID01_2_ConcEst_TratarDados.classes.dados_execucao.DadosExecucao import DadosExecucao
from ID01_2_ConcEst_TratarDados.classes.email.send.SendEmail import SendEmail
from ID01_2_ConcEst_TratarDados.classes.email.send.SendEmailOutlook import SendEmailOutlook
from ID01_2_ConcEst_TratarDados.classes.utils.BackupSqlite import BackupSQLite


# Imports dos pacotes externos
import traceback
from datetime import datetime
from os import path


class Initialization:
    """
    Classe Reponsavel pela organização do código da execução dos passos categorizados como inicialização
    
    Parâmetros:
    
    Retorna:
       
    """

    @staticmethod
    def execute():
        """
        Funcao de execução da inicialização.
        
        Parâmetros:
           
        Retorna:

        """
        try:
            Log.write_log("Initialization Started")
            Log.write_log(GenericReusable.get_system_info())
            Log.write_log(GenericReusable.get_computer_usage())
            Log.write_log(GenericReusable.get_monitor_info())

            # CLASSE COMENTADA POR PRECISAR DE PARÂMETROS CUSTOMIZADOS
            # DESCOMENTAR EM DESENVOLVIMENTO, apenas o que for necessário
            #SendEmail.configure_email(email_server_smtp="", email_porta_smtp=0, usuario="", senha="")

            # SendEmailOutlook.configure_email()

            # Executar Robot Stream
            if(InitAllSettings.config["IniciarRobotStream"].upper() == "SIM"):
                RobotStream.start()

            # Verifica se a opção de backup do banco de dados SQLite está ativada no arquivo de configuração. 
            # Se estiver, cria um backup do banco de dados no caminho especificado, também vindo da configuração.

            if(InitAllSettings.config["BackupSqlite"].upper() == "SIM"):
                BackupSQLite.criar_backup_sqlite(diretorio_backup=InitAllSettings.config["CaminhoBackupSqlite"],
                                                 dias_para_atualizar=3)

            # Insere um registro no banco com dados com informacoes do inicio da execucao
            DadosExecucao.inserir_tabela_dados_execucao(QueueManager.items_queue)

            Log.write_log("Itens encontrados na fila antes de adicionar novos itens a fila: " + QueueManager.items_queue.__str__())
            
            # Enviando email de inicialização
            if(InitAllSettings.config["EmailInicial"].upper() == "SIM"):
                SendEmail.send_email_inicial(envio_para=InitAllSettings.config["EmailDestinatarios"])
                #SendEmailOutlook.send_email_inicial(envio_para=InitAllSettings.config["EmailDestinatarios"])


            # Se for indicado que é necessário Gravar a tela do processo, inicia aqui.
            if(InitAllSettings.config["GravarTela"].upper() == "SIM"):
                ScreenRecorder.iniciar_gravacao()

            # Finaliza os aplicativos antes de iniciar as novas aplicações
            KillAllProcesses.execute(['excel.exe','winword.exe'])

            # Realiza a inicialização das aplicações
            InitAllApplications.execute(first_run=True)

            # Realiza o update da quantidade de itens a serem processados
            DadosExecucao.update_tabela_dados_execucao_itens_new(QueueManager.items_queue)
            
            # Atualiza no InitAllSettings a quantidade de itens que estão a serem processados
            InitAllSettings.qtd_itens_a_processar_ini_exec = QueueManager.items_queue

            Log.write_log("Initialization Finished")

        except BusinessRuleException as err:

            # Captura todo detalhe do erro
            traceback_erro = traceback.format_exc()
            InitAllSettings.exception_initialization = err

            caminho_screenshot = ''
            InitAllSettings.caminho_screenshot_erro_init = caminho_screenshot

            # Se deve Capturar Tela ou nao
            if(InitAllSettings.config["CapturarScreenshot"].upper() == "SIM"):
                #Tirando print do erro antes de fechar
                try:
                    caminho_screenshot = path.join(InitAllSettings.config["CaminhoExceptionScreenshots"] , "InitBusExceptionScreenshot_" + datetime.now().strftime("%d%m%Y_%H%M%S%f") + ".png")
                    InitAllSettings.save_screenshot(path=caminho_screenshot)
                    InitAllSettings.caminho_screenshot_erro_init = caminho_screenshot
                except Exception as err_screenshot:
                    Log.write_log(f"Não foi possível capturar a screenshot de erro. {str(err_screenshot)}",
                                      log_level=LogLevel.FATAL,
                                      error_type=ErrorType.APP_ERROR)

            # Enviando email de erro na Inicialização
            if(InitAllSettings.config["EmailErroInicializacao"].upper() == "SIM"):
                SendEmail.send_email_erro(business=True, 
                                anexos=[caminho_screenshot], 
                                detalhes_erro='Falha na inicialização ' + str(err), 
                                envio_para=InitAllSettings.config["EmailDestinatarios"])

                # SendEmailOutlook.send_email_erro(business=True, 
                #                 anexos=[caminho_screenshot], 
                #                 detalhes_erro='Falha na inicialização ' + str(err), 
                #                 envio_para=InitAllSettings.config["EmailDestinatarios"])

            # Incluindo erro no relatório sintético, com horário de finalização
            DadosExecucao.update_tabela_dados_execucao(traceback_erro,'BUSINESS_EXCEPTION',caminho_screenshot)

            ExecutionControl.send_error(err, caminho_screenshot)

            raise err
        except Exception as err:
            
            # Captura todo detalhe do erro
            traceback_erro = traceback.format_exc()
            InitAllSettings.exception_initialization = err

            caminho_screenshot = ''
            InitAllSettings.caminho_screenshot_erro_init = caminho_screenshot
            
            # Se deve Capturar Tela ou nao
            if(InitAllSettings.config["CapturarScreenshot"].upper() == "SIM"):
                #Tirando print do erro antes de fechar
                try:    
                    caminho_screenshot = path.join(InitAllSettings.config["CaminhoExceptionScreenshots"] , "InitAppExceptionScreenshot_" + datetime.now().strftime("%d%m%Y_%H%M%S%f") + ".png")
                    InitAllSettings.save_screenshot(path=caminho_screenshot)
                    InitAllSettings.caminho_screenshot_erro_init = caminho_screenshot
                except Exception as err_screenshot:
                    Log.write_log(f"Não foi possível capturar a screenshot de erro. {str(err_screenshot)}",
                                      log_level=LogLevel.FATAL,
                                      error_type=ErrorType.APP_ERROR)

            # Enviando email de erro na Inicialização
            if(InitAllSettings.config["EmailErroInicializacao"].upper() == "SIM"):
                SendEmail.send_email_erro(business=False, 
                                anexos=[caminho_screenshot], 
                                detalhes_erro='Falha na inicialização ' + str(err), 
                                envio_para=InitAllSettings.config["EmailDestinatarios"])
                
                # SendEmailOutlook.send_email_erro(business=True, 
                #                 anexos=[caminho_screenshot], 
                #                 detalhes_erro='Falha na inicialização ' + str(err), 
                #                 envio_para=InitAllSettings.config["EmailDestinatarios"])

            # Incluindo erro no relatório sintético, com horário de finalização
            DadosExecucao.update_tabela_dados_execucao(traceback_erro,'APPLICATION_EXCEPTION',caminho_screenshot)
            
            ExecutionControl.send_error(err, caminho_screenshot)
            
            raise err
