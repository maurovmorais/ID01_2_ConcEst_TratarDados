# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
# Imports dos pacotes externos
import win32com.client as win32
from pathlib import Path
from os import path

class SendEmailOutlook:
    """
    Classe responsável pelo envio de email pelo outlook, enviando usuario, senha e qual servidor vai ser usado
    
    Parâmetros:

    Retorna:

    """

    @classmethod
    def configure_email(cls):
        """
        Construtor do envio de email.

        Parâmetros:
        

        Retorna:

        """ 
        cls.nome_processo = InitAllSettings.config['NomeProcesso']
        
        
    @classmethod
    def send_email_inicial(cls, envio_para:str, cc:str=None, bcc:str=None):
        """
        Envia o email inicial do robô, apenas precisando informar quem deve receber, separando por ;

        Parâmetros:
        - envio_para (str): destinatários separados por ';'.
        - cc (str): destinatários em cópia separados por ';' (opcional, default=None).
        - bcc (str): destinatários em cópia oculta separados por ';' (opcional, default=None).

        Retorna:

        """
        outlook:win32.CDispatch = win32.Dispatch('outlook.application')
        mail:win32.CDispatch = outlook.CreateItem(0) #Criando um item do tipo Email (0)

        #Lendo template
        file_template = open(InitAllSettings.caminho_template_email_inicio, mode="r",encoding="utf-8")
        email_texto = file_template.read()
        file_template.close()

        mail.HTMLBody = email_texto.replace("*NOME_ROBO*", cls.nome_processo)
        mail.Subject = "Inicio execução: " + cls.nome_processo
        mail.To = envio_para

        if(cc is not None): mail.CC = cc
        if(bcc is not None): mail.BCC = bcc
        
        #Envia o email inicial
        try:
            Log.write_log("Enviando email inicial")
            mail.Send()
            Log.write_log("Email enviado com sucesso")
        except Exception as err:
            Log.write_log(mensagem_log="Erro enviando email inicial: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err

    @classmethod  
    def send_email_final(cls, envio_para:str, anexos:list=None, sucesso:bool=True, cc:str=None, bcc:str=None):
        """
        Envia o email final do robô.
        Recebendo o horário do início da execução, o horário final, para quem é necessário enviar (separando por ;) e os relatórios finais
        
        Parâmetros:
        - envio_para (str): destinatários separados por ';'
        - anexos (list): lista de anexos. (opcional, default=None)
        - sucesso (bool): indica se a execução foi bem-sucedida. (opcional, default=True)
        - cc (str): destinatários em cópia separados por ';'. (opcional, default=None)
        - bcc (str): destinatários em cópia oculta separados por ';'. (opcional, default=None)

        Retorna:

        """
        outlook:win32.CDispatch = win32.Dispatch('outlook.application')
        mail:win32.CDispatch = outlook.CreateItem(0) #Criando um item do tipo Email (0)
        
        #Lendo template
        file_template = open(InitAllSettings.caminho_template_email_final, mode="r",encoding="utf-8")
        email_texto = file_template.read()
        file_template.close()
        
        horario_inicio = InitAllSettings.datahora_inicio_exec
        horario_fim = InitAllSettings.datahora_fim_exec

        status_finalizacao = "com sucesso" if sucesso else "com erros"
        mail.HTMLBody = email_texto.replace("*NOME_ROBO*", cls.nome_processo).replace("*DATAHORA_INI*", horario_inicio).replace("*DATAHORA_FIM*", horario_fim).replace("*FINALIZACAO*", status_finalizacao)
        mail.Subject = "Finalização da execução: " + cls.nome_processo
        mail.To = envio_para

        if(cc is not None): mail.CC = cc
        if(bcc is not None): mail.BCC = bcc
        
        #Inserindo anexos
        if(anexos is not None):
            for anexo in anexos:
                mail.Attachments.Add(anexo)

        #Envia o email de finalização
        try:
            Log.write_log("Enviando email de finalização")
            mail.Send()
            Log.write_log("Email enviado com sucesso")
        except Exception as err:
            Log.write_log(mensagem_log="Erro enviando email de finalização: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err

    @classmethod
    def send_email_erro(cls, envio_para:str, anexos:list, detalhes_erro:str, business:bool=False, cc:str=None, bcc:str=None):
        """
        Envia um email em casos de erro.

        Parâmetros:
        - envio_para (str): destinatários separados por ';'.
        - anexos (list): lista de anexos.
        - detalhes_erro (str): detalhes do erro.
        - business (bool): indica se o erro é de regra de negócio. (opcional, default=False)
        - cc (str): destinatários em cópia separados por ';'. (opcional, default=None)
        - bcc (str): destinatários em cópia oculta separados por ';'. (opcional, default=None)

        Retorna:

        """
        outlook:win32.CDispatch = win32.Dispatch('outlook.application')
        mail:win32.CDispatch = outlook.CreateItem(0) #Criando um item do tipo Email (0)

        #Lendo template
        file_template = open(InitAllSettings.caminho_template_email_erro_encontrado, mode="r",encoding="utf-8")
        email_texto = file_template.read()
        file_template.close()

        mail.HTMLBody = email_texto.replace("*NOME_ROBO*", cls.nome_processo).replace("*ERRO_DETALHES*", detalhes_erro)
        mail.HTMLBody = mail.HTMLBody.replace("*ERRO_TIPO*", "ERRO DE REGRA DE NEGÓCIO") if business else mail.HTMLBody.replace("*ERRO_TIPO*", "ERRO INESPERADO")
        mail.Subject = "Erro durante a execução: " + cls.nome_processo
        mail.To = envio_para

        if(cc is not None): mail.CC = cc
        if(bcc is not None): mail.BCC = bcc
        
        #Inserindo anexos
        if(anexos is not None):
            for anexo in anexos:
                mail.Attachments.Add(anexo)

        #Envia o email inicial
        try:
            Log.write_log("Enviando email de erro")
            mail.Send()
            Log.write_log("Email enviado com sucesso")
        except Exception as err:
            Log.write_log(mensagem_log="Erro enviando email de erro: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err
       
    @classmethod
    def send_email(cls, corpo_email:str, envio_para:str, assunto:str, anexos:list=None, html:bool=False, cc:str=None, bcc:str=None):
        """
        Simplesmente envia um email normal. Pode ser usado em vários lugares no código, porém é necessário informar um corpo para o email.

        Parâmetros:
        - corpo_email (str): corpo do email.
        - envio_para (str): destinatários separados por ';'.
        - assunto (str): assunto do email.
        - anexos (list): lista de anexos. (opcional, default=None).
        - html (bool): indica se o corpo do email é HTML (default=False).
        - cc (str): destinatários em cópia separados por ';' (opcional, default=None).
        - bcc (str): destinatários em cópia oculta separados por ';' (opcional, default=None).

        Retorna:
        
        """
        outlook:win32.CDispatch = win32.Dispatch('outlook.application')
        mail:win32.CDispatch = outlook.CreateItem(0) #Criando um item do tipo Email (0)
        
        if(html):
            mail.HTMLBody = corpo_email
        else:
            mail.Body = corpo_email

        mail.Subject = assunto
        mail.To = envio_para
        if(cc is not None): mail.CC = cc
        if(bcc is not None): mail.BCC = bcc
        
        #Inserindo anexos
        if(anexos is not None):
            for anexo in anexos:
                mail.Attachments.Add(anexo)

        #Envia o email customizado
        try:
            Log.write_log("Enviando email customizado")
            mail.Send()
            Log.write_log("Email enviado com sucesso")
        except Exception as err:
            Log.write_log(mensagem_log="Erro enviando email customizado: " +str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err
      