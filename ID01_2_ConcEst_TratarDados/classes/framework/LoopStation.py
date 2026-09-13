# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllApplications import InitAllApplications
from ID01_2_ConcEst_TratarDados.classes.framework.GetTransaction import GetTransaction
from ID01_2_ConcEst_TratarDados.classes.framework.KillAllProcesses import KillAllProcesses
from ID01_2_ConcEst_TratarDados.classes.dados_execucao.DadosExecucao import DadosExecucao
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.ExecutionControl import ExecutionControl
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException, TerminateException
from ID01_2_ConcEst_TratarDados.classes.queue.QueueManager import QueueManager
from ID01_2_ConcEst_TratarDados.classes.framework.Process import Process
from ID01_2_ConcEst_TratarDados.classes.email.send.SendEmail import SendEmail
from ID01_2_ConcEst_TratarDados.classes.email.send.SendEmailOutlook import SendEmailOutlook
import ID01_2_ConcEst_TratarDados.classes.utils.GenericReusable as GenericReusable

# Imports dos pacotes externos
import traceback
import json
from datetime import datetime
from os import path

class LoopStation:
    """
    Classe reponsável pela organização do código da execução dos passos categorizados como repetição para processar os itens da fila.
    
    Parâmetros:
       
    
    Retorna:
       
    """
    
    
    @classmethod
    def execute(cls):
        """
        Metodo de execução do loop dos itens da fila.
        
        Parâmetros:

        Retorna:

        """
        
        Log.write_log("Loop Station Started")

        GetTransaction.execute() 
        
        while GetTransaction.queue_item is not None:
            for tentativa in range(InitAllSettings.config["MaxRetryNumber"]):
                try:
                    
                    
                    exp_type_exception = None
                        
                    datahora_inicio_item = datetime.now()
                    datahora_inicio_item_str = datahora_inicio_item.strftime("%d/%m/%Y %H:%M:%S")

                    Log.write_log(mensagem_log="Tentativa: " + (tentativa+1).__str__(), 
                                        referencia=GetTransaction.queue_item['referencia'] )
                    
                    # Insere na tabela de dados da execução as informações sobre o item que está sendo executado
                    DadosExecucao.inserir_tabela_dados_itens(GetTransaction.queue_item['id'],
                                                                    InitAllSettings.config['FilaProcessamento'],
                                                                    GetTransaction.queue_item['referencia'],
                                                                    json.dumps(GetTransaction.queue_item['info_adicionais'],ensure_ascii=False))
                    



                    # Executando o process
                    Process.execute()

                    datahora_fim_item = datetime.now()
                    datahora_fim_item_str = datahora_fim_item.strftime("%d/%m/%Y %H:%M:%S")
                    
                    # Marcando item como finalizado com sucesso
                    QueueManager.update_status_item()
                    

                    # Realiza update na tabela de dados do item e insere que ocorreu sucesso
                    DadosExecucao.update_tabela_dados_itens('SUCESSO')

                    
                    # Reinicia a quantidade de tentativas consecutivas
                    InitAllSettings.qtde_itens_consecutive_exceptions = 0

                except BusinessRuleException as err:
                    traceback_erro = traceback.format_exc()

                    InitAllSettings.qtde_itens_consecutive_exceptions = 0
                                        
                    datahora_fim_item = datetime.now()
                    datahora_fim_item_str = datahora_fim_item.strftime("%d/%m/%Y %H:%M:%S")
                    
                    exp_type_exception = err

                    caminho_screenshot = ''
                    # Se deve Capturar Tela ou nao
                    if(InitAllSettings.config["CapturarScreenshot"].upper() == "SIM"):
                        try:    
                            caminho_screenshot = path.join(InitAllSettings.config["CaminhoExceptionScreenshots"], 
                                                                "ProcessBusExceptionScreenshot_" + datahora_fim_item.strftime("%d%m%Y_%H%M%S%f") + ".png")

                            # Captura screenshot de erro Business Exception
                            InitAllSettings.save_screenshot(path=caminho_screenshot)
                        except Exception as err_screenshots:
                            Log.write_log(f"Não foi possível capturar a screenshot de erro. {str(err_screenshots)}",
                                              referencia=GetTransaction.queue_item['referencia'],
                                              log_level=LogLevel.FATAL,
                                              error_type=ErrorType.APP_ERROR)
                    
                    Log.write_log(mensagem_log="Erro de negócio: " + traceback_erro, 
                                        referencia=GetTransaction.queue_item['referencia'], 
                                        log_level=LogLevel.ERROR, 
                                        error_type=ErrorType.BUSINESS_ERROR)
                    
                    

                    # Chamando a classe KillAllProcesses em caso de erro de business
                    KillAllProcesses.execute()
                    
                    #Marcando item como erro de business
                    QueueManager.update_status_item(excecao=err, obs=err.__str__())

                    # Realiza o update dos dados do item no banco de dados da execucao 
                    DadosExecucao.update_tabela_dados_itens('FALHA','NEGOCIO',err.__str__(),caminho_screenshot)

                    #Enviando email com print do erro
                    if(InitAllSettings.config["EmailCadaErro"].upper() == "SIM"):
                        SendEmail.send_email_erro(business=True, anexos=[caminho_screenshot], 
                                                  detalhes_erro=str(err), 
                                                  envio_para=InitAllSettings.config["EmailDestinatarios"])                        
                        # SendEmailOutlook.send_email_erro(business=True, 
                        #                                  anexos=[caminho_screenshot], 
                        #                                  detalhes_erro=str(err), 
                        #                                  envio_para=InitAllSettings.config["EmailDestinatarios"])
                    
                    break
                except TerminateException as err:
                    # Caso seja direcionado para esse tipoe de excecao significa que o processo nao precisou seguir até o final para resultar em sucesso, preciso ser parado previamente
                    
                    # Marcando item como finalizado com sucesso
                    QueueManager.update_status_item()
                    

                    # Realiza update na tabela de dados do item e insere que ocorreu sucesso
                    DadosExecucao.update_tabela_dados_itens('SUCESSO')

                    
                    # Reinicia a quantidade de tentativas consecutivas
                    InitAllSettings.qtde_itens_consecutive_exceptions = 0
                    
                except Exception as err:
                    traceback_erro = traceback.format_exc()

                    exp_type_exception = err
                    datahora_fim_item = datetime.now()
                    datahora_fim_item_str = datahora_fim_item.strftime("%d/%m/%Y %H:%M:%S")

                    Log.write_log(GenericReusable.get_computer_usage())
                    
                    Log.write_log(mensagem_log="Erro, tentativa " + (tentativa+1).__str__() + ": " + traceback_erro, 
                                      referencia=GetTransaction.queue_item['referencia'], 
                                      log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)

                    caminho_screenshot = ''
                    # Se deve Capturar Tela ou nao
                    if(InitAllSettings.config["CapturarScreenshot"].upper() == "SIM"):
                        #Tirando print de erro
                        try:
                            caminho_screenshot = path.join(InitAllSettings.config["CaminhoExceptionScreenshots"] ,
                                                                "ProcesAppExceptionScreenshot_" + datahora_fim_item.strftime("%d%m%Y_%H%M%S%f") + ".png")
                            InitAllSettings.save_screenshot(path=caminho_screenshot)
                        except Exception as err_screenshots:
                            Log.write_log(f"Não foi possível capturar a screenshot de erro. {str(err_screenshots)}",
                                              referencia=GetTransaction.queue_item['referencia'],
                                              log_level=LogLevel.FATAL,
                                              error_type=ErrorType.APP_ERROR)
                    
                    # Realiza Update para inserir que houve falha de sistema
                    # Insere no item que houve falha de sistema juntamente com a screenshot do erro
                    DadosExecucao.update_tabela_dados_itens('FALHA','SISTEMA',traceback_erro,caminho_screenshot)
                        

                    if(tentativa+1 == InitAllSettings.config["MaxRetryNumber"]):
                        QueueManager.update_status_item(excecao=err, obs=traceback_erro)
                        #Enviando email com print do erro
                        if(InitAllSettings.config["EmailCadaErro"].upper() == "SIM"):          
                            SendEmail.send_email_erro(business=False, 
                                                      anexos=[caminho_screenshot], 
                                                      detalhes_erro=str(err), 
                                                      envio_para=InitAllSettings.config["EmailDestinatarios"])
                            # SendEmailOutlook.send_email_erro(business=False, 
                            #                                  anexos=[caminho_screenshot], 
                            #                                  detalhes_erro=str(err), 
                            #                                  envio_para=InitAllSettings.config["EmailDestinatarios"])
                            

                    #Chamando a classe KillAllProcesses em caso de erro de aplicação
                    KillAllProcesses.execute()
        
                    # Iniciando novamente o processamento
                    InitAllApplications.execute(first_run=False)

                    if not(tentativa+1 == InitAllSettings.config["MaxRetryNumber"]):
                        continue

                else:
                    #Insira aqui qualquer código necessário para voltar ao estado inicial em caso de sucesso, para executar um possivel próximo item

                    break
            

            #Verifica se o item processado saiu do loop for com Exceção
            if type(exp_type_exception) == Exception:
                #Incrementa +1 devido a exceção do Item do Processado
                InitAllSettings.qtde_itens_consecutive_exceptions += 1

                # Verifica se a quantidade maxima de erro de sistema de itens consecutivos foi atingida
                if(InitAllSettings.qtde_itens_consecutive_exceptions >= InitAllSettings.config['MaxConsecutiveSystemExceptions']):
                    Log.write_log(f"Quantidade de Exceptions Consecutivas atingida: {InitAllSettings.config['MaxConsecutiveSystemExceptions']}") 
                    break

            #Verifica se foi solicitado a interrupção da execução
            interrupted = ExecutionControl.is_interrupted()

            #Processamento continua até a classe informar que não existe itens novos para processar
            if(interrupted == True): 
                GetTransaction.queue_item = None
            else: 
                GetTransaction.execute()


        Log.write_log("Loop Station Finished")
