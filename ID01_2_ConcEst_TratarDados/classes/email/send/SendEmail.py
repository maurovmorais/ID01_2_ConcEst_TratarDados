# Imports dos módulos internos do projeto
# Carrega o InitAllSettings - Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings 
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType

# Imports dos pacotes externos
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os


class SendEmail:
    """
    Classe responsável pelo envio de email, enviando usuario, senha e qual servidor vai ser usado.

    Utiliza a biblioteca padrão `smtplib` do Python (SMTP), substituindo a dependência do plugin de
    email do Botcity.

    Parâmetros:

    Retorna:

    """

    @classmethod
    def configure_email(cls, email_server_smtp:str, email_porta_smtp:int, usuario:str, senha:str):
        """
        Realiza as configuracoes de email que fica atrelada a classe
        
        Parâmetros:
            - email_server_smtp (str): endereço do servidor SMTP.
            - email_porta_smtp (int): porta do servidor SMTP.
            - usuario (str): nome de usuário para autenticação.
            - senha (str): senha para autenticação.
        Retorna:
           
        """
        cls.email_server_smtp = email_server_smtp
        cls.email_porta_smtp = email_porta_smtp
        cls.usuario = usuario
        cls.senha = senha
        cls.nome_processo = InitAllSettings.config['NomeProcesso']

    @classmethod
    def _enviar_smtp(cls, assunto:str, corpo_email:str, envio_para:list, cc:list=None,
                      bcc:list=None, anexos:list=None, html:bool=True):
        """
        Método interno que realiza o envio efetivo do email via SMTP.

        Parâmetros:
        - assunto (str): assunto do e-mail.
        - corpo_email (str): corpo do e-mail.
        - envio_para (list): lista de destinatários.
        - cc (list): lista de destinatários em cópia. (opcional, default=None)
        - bcc (list): lista de destinatários em cópia oculta. (opcional, default=None)
        - anexos (list): lista de caminhos de anexos. (opcional, default=None)
        - html (bool): indica se o corpo do e-mail é HTML. (default=True)

        Retorna:
        """
        msg = MIMEMultipart()
        msg['From'] = cls.usuario
        msg['To'] = "; ".join(envio_para)
        msg['Subject'] = assunto

        cc_lista = cc or []
        if cc_lista:
            msg['Cc'] = "; ".join(cc_lista)

        msg.attach(MIMEText(corpo_email, "html" if html else "plain", "utf-8"))

        for caminho_anexo in (anexos or []):
            if not caminho_anexo or not os.path.exists(caminho_anexo):
                continue
            with open(caminho_anexo, "rb") as arquivo:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(arquivo.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(caminho_anexo)}")
            msg.attach(part)

        todos_destinatarios = envio_para + cc_lista + (bcc or [])

        with smtplib.SMTP(cls.email_server_smtp, cls.email_porta_smtp) as server:
            server.starttls()
            server.login(cls.usuario, cls.senha)
            server.sendmail(cls.usuario, todos_destinatarios, msg.as_string())

    @classmethod
    def send_email_inicial(cls, envio_para:str, cc:str=None, bcc:str=None):
        """
        Envia o email inicial do robô, apenas precisando informar quem deve receber (separado por ;) e o nome do robô

        Parâmetros:
        - envio_para (str): destinatários separados por ';'.
        - cc (str): destinatários em cópia separados por ';'. (opcional, default=None)
        - bcc (str): destinatários em cópia oculta separados por ';'. (opcional, default=None)

        Retorna:

        """

        #Lendo template
        file_template = open(InitAllSettings.caminho_template_email_inicio, mode="r",encoding="utf-8")
        email_texto = file_template.read()
        file_template.close()

        email_texto = email_texto.replace("*NOME_ROBO*", cls.nome_processo)
        email_assunto = "Inicio execução: " + cls.nome_processo
        
        envio_para_lista = envio_para.split(";") if(envio_para is not None) else []
        cc_lista = cc.split(";") if(cc is not None) else []
        bcc_lista = bcc.split(";") if(bcc is not None) else []

        try:
            Log.write_log("Enviando email inicial")
            cls._enviar_smtp(assunto=email_assunto, corpo_email=email_texto,
                              envio_para=envio_para_lista, cc=cc_lista, bcc=bcc_lista,
                              html=True)
            Log.write_log("Email enviado com sucesso")
        except Exception as err:
            Log.write_log(mensagem_log="Erro enviando email inicial: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err
    
    @classmethod
    def send_email_final(cls, envio_para:str, cc:str=None, bcc:str=None, anexos:list=None, sucesso:bool=True):
        """
        Envio do e-mail de finalização do robô.
        Recebendo o horário do início da execução, o horário final, para quem é necessário enviar (separado por ;) e os relatórios finais
        
        Parâmetros:
        - envio_para (str): destinatários separados por ';'.
        - cc (str): destinatários em cópia separados por ';'. (opcional, default=None)
        - bcc (str): destinatários em cópia oculta separados por ';'. (opcional, default=None)
        - anexos (list): lista de anexos (opcional, default=None).
        - sucesso (bool): indica se a execução foi bem-sucedida (opcional, default=True).
        
        Retorna:

        """
        
        #Lendo template
        file_template = open(InitAllSettings.caminho_template_email_final, mode="r",encoding="utf-8")
        email_texto = file_template.read()
        file_template.close()

        horario_inicio = InitAllSettings.datahora_inicio_exec
        horario_fim = InitAllSettings.datahora_fim_exec

        status_finalizacao = "com sucesso" if sucesso else "com erros"
        email_texto = email_texto.replace("*NOME_ROBO*", cls.nome_processo).replace("*DATAHORA_INI*", horario_inicio).replace("*DATAHORA_FIM*", horario_fim).replace("*FINALIZACAO*", status_finalizacao)
        email_assunto = "Finalização da execução: " + cls.nome_processo

        envio_para_lista = envio_para.split(";") if(envio_para is not None) else []
        cc_lista = cc.split(";") if(cc is not None) else []
        bcc_lista = bcc.split(";") if(bcc is not None) else []

        try:
            Log.write_log("Enviando email de finalização")
            cls._enviar_smtp(assunto=email_assunto, corpo_email=email_texto,
                              envio_para=envio_para_lista, cc=cc_lista, bcc=bcc_lista,
                              anexos=anexos, html=True)
            Log.write_log("Email enviado com sucesso")
        except Exception as err:
            Log.write_log(mensagem_log="Erro enviando email de finalização: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err
    
    @classmethod
    def send_email_erro(cls, envio_para:str, anexos:list, detalhes_erro:str, business:bool=False, cc:str=None, bcc:str=None):
        """
        Envio do e-mail em casos de erro durante a execução do robô.

        Parâmetros:
        - envio_para (str): destinatários separados por ';'.
        - anexos (list): lista de anexos.
        - detalhes_erro (str): detalhes do erro.
        - business (bool): indica se o erro é de regra de negócio. (opcional, default=False)
        - cc (str): destinatários em cópia separados por ';'. (opcional, default=None)
        - bcc (str): destinatários em cópia oculta separados por ';'. (opcional, default=None)

        Retorna:

        """
        
        #Lendo template
        file_template = open(InitAllSettings.caminho_template_email_erro_encontrado, mode="r",encoding="utf-8")
        email_texto = file_template.read()
        file_template.close()

        email_texto = email_texto.replace("*NOME_ROBO*", cls.nome_processo).replace("*ERRO_DETALHES*", detalhes_erro)
        email_texto = email_texto.replace("*ERRO_TIPO*", "ERRO DE REGRA DE NEGÓCIO") if business else email_texto.replace("*ERRO_TIPO*", "ERRO INESPERADO")
        email_assunto = "Erro durante a execução: " + cls.nome_processo
        
        envio_para_lista = envio_para.split(";") if(envio_para is not None) else []
        cc_lista = cc.split(";") if(cc is not None) else []
        bcc_lista = bcc.split(";") if(bcc is not None) else []

        try:
            Log.write_log("Enviando email de erro")
            cls._enviar_smtp(assunto=email_assunto, corpo_email=email_texto,
                              envio_para=envio_para_lista, cc=cc_lista, bcc=bcc_lista,
                              anexos=anexos, html=True)
            Log.write_log("Email enviado com sucesso")
        except Exception as err:
            Log.write_log(mensagem_log="Erro enviando email de erro: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err

    @classmethod 
    def send_email(cls, corpo_email:str, envio_para:str, cc:str, bcc:str, assunto:str, anexos:list=None, html:bool=False):        
        """
        Simplesmente envia um email normal. Pode ser usado em vários lugares no código, porém é necessário informar um corpo para o email 

        Parâmetros:
        - corpo_email (str): corpo do e-mail.
        - envio_para (str): destinatários separados por ';'.
        - cc (str): destinatários em cópia separados por ';'.
        - bcc (str): destinatários em cópia oculta separados por ';'
        - assunto (str): assunto do e-mail.
        - anexos (list): lista de anexos. (opcional, default=None)
        - html (bool): indica se o corpo do e-mail é HTML. (default=False)

        Retorna:

        """
        envio_para_lista = envio_para.split(";") if(envio_para is not None) else []
        cc_lista = cc.split(";") if(cc is not None) else []
        bcc_lista = bcc.split(";") if(bcc is not None) else []

        try:
            Log.write_log("Enviando email customizado")
            cls._enviar_smtp(assunto=assunto, corpo_email=corpo_email,
                              envio_para=envio_para_lista, cc=cc_lista, bcc=bcc_lista,
                              anexos=anexos, html=html)
            Log.write_log("Email enviado com sucesso")
        except Exception as err:
            Log.write_log(mensagem_log="Erro enviando email customizado: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise
