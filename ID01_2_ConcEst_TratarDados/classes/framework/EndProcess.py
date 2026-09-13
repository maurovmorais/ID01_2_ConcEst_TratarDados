# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.framework.CloseAllApplications import CloseAllApplications
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.CredentialManager import CredentialManager
from ID01_2_ConcEst_TratarDados.classes.utils.ExecutionControl import ExecutionControl
from ID01_2_ConcEst_TratarDados.classes.utils.ScreenRecorder import ScreenRecorder
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException, TerminateException
from ID01_2_ConcEst_TratarDados.classes.framework.KillAllProcesses import KillAllProcesses
from ID01_2_ConcEst_TratarDados.classes.dados_execucao.DadosExecucao import DadosExecucao
from ID01_2_ConcEst_TratarDados.classes.relatorios.Relatorios import Relatorios
from ID01_2_ConcEst_TratarDados.classes.email.send.SendEmail import SendEmail
from ID01_2_ConcEst_TratarDados.classes.email.send.SendEmailOutlook import SendEmailOutlook


# Imports dos pacotes externos
from datetime import datetime
from os import path
from os import remove

class EndProcess:
    """
    Classe reponsável pela organização do código da execução dos passos categorizados como finalização da execução.
    
    Parâmetros:
       
    
    Retorna:
       
    """
    
    
    @classmethod
    def execute(cls):
        """
        Metodo de execução dos passos da finalização da execucao
        
        Parâmetros:

        Retorna:

        """
        # 468 Fim do Processamento

        Log.write_log("End Process Started")

        InitAllSettings.datahora_fim_exec = datetime.now()
        InitAllSettings.datahora_fim_exec = InitAllSettings.datahora_fim_exec.strftime("%d/%m/%Y %H:%M:%S")

        #Fechando aplicativos no final do processamento
        try:
            CloseAllApplications.execute()
            KillAllProcesses.execute(['chromedriver.exe'])
        except Exception as err:
            Log.write_log(mensagem_log="Fechando aplicativos pelo KillAllProcesses", log_level=LogLevel.WARN)
            KillAllProcesses.execute(['chromedriver.exe'])


        # Finaliza o Gravador de Tela
        if(InitAllSettings.config["GravarTela"].upper() == "SIM"):
            ScreenRecorder.finalizar_gravacao()
        
        
        # Inserir update para tabela dados execucao Finalizar
        DadosExecucao.update_tabela_dados_execucao()

        caminho_relatorio_analitico = Relatorios.preencher_analitico()
        caminho_relatorio_sintetico = Relatorios.preencher_sintetico()


        #Enviando email final com os relatórios analítico e sintético
        if(InitAllSettings.config["EmailFinal"].upper() == "SIM"):   
            SendEmail.send_email_final(envio_para=InitAllSettings.config["EmailDestinatarios"], 
                                       anexos=[caminho_relatorio_analitico, caminho_relatorio_sintetico], 
                                       sucesso=True)
            # SendEmailOutlook.send_email_final(envio_para=InitAllSettings.config["EmailDestinatarios"], 
            #                                   anexos=[caminho_relatorio_analitico, caminho_relatorio_sintetico], 
            #                                   sucesso=True)

            # Deleta arquivo excel sintetico e analitico caso exista
            if path.exists(caminho_relatorio_analitico):
                remove(caminho_relatorio_analitico)
            
            if path.exists(caminho_relatorio_sintetico):
                remove(caminho_relatorio_sintetico)


        # Se for indicando que é necessário incluir linhas no SQL Server, ou verificado que não está rodando em debug (vs code ou test task)
        if(InitAllSettings.config["RelatorioRAAS"].upper() == "SIM" and not ExecutionControl.is_test_task):

            caminho_csv_analitico = Relatorios.preencher_csv_analitico_banco_raas()
            caminho_csv_sintetico = Relatorios.preencher_csv_sintetico_banco_raas()
            SendEmail.configure_email(email_server_smtp=InitAllSettings.config['ServerSmtpRAAS'],
                                      email_porta_smtp=InitAllSettings.config['PortaSmtpRAAS'],
                                      usuario=CredentialManager.get_credential(label=InitAllSettings.config['CRED_EMAIL_RAAS'],
                                                                            key='USER'), 
                                      senha=CredentialManager.get_credential(label=InitAllSettings.config['CRED_EMAIL_RAAS'],
                                                                            key='PASSWORD'))
            
            SendEmail.send_email(corpo_email='',
                                 envio_para=InitAllSettings.config['DestinatarioRAAS'],
                                 assunto=f'Relatorio RAAS {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}',
                                 anexos=[caminho_csv_analitico,caminho_csv_sintetico])
            
            # Deleta arquivo csv sintetico e analitico caso exista
            if path.exists(caminho_csv_analitico):
                remove(caminho_csv_analitico)
            
            if path.exists(caminho_csv_sintetico):
                remove(caminho_csv_sintetico)
        
        Log.write_log("Finalizando task.")

        DadosExecucao.refresh_counting_items()
        if InitAllSettings.exception_initialization is not None:
            ExecutionControl.send_error(InitAllSettings.exception_initialization)
            ExecutionControl.finish_task(sucesso=False, mensagem="Task finalizada com falha na inicialização, verifique os logs de execução.")
        elif InitAllSettings.exception_process is not None:
            ExecutionControl.send_error(InitAllSettings.exception_process)
            ExecutionControl.finish_task(sucesso=False, mensagem="Task finalizada com falha na fase de processamento, verifique os logs de execução.")
        else: 
            ExecutionControl.finish_task(sucesso=True, mensagem="Task finalizada com sucesso.")
        
        # 77 Fim
        Log.write_log("End Process Finished")