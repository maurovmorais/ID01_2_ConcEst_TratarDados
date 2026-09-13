# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
# Imports dos pacotes externos
import win32com.client as win32
import os

class GetEmailOutlook:
    """
    Classe responsável pela captura de emails do Outlook, aplicando filtros opcionais por pasta, assunto e remetente.

    Parâmetros:

    Retorna:
    """
    
    @classmethod
    def _obter_pasta(cls, pasta_email:str) -> win32.CDispatch:
        """
        Obtém a pasta do Outlook pelo nome.

        Parâmetros:
        - pasta_email (str): Nome da pasta para filtragem.

        Retorna:
        - pasta_email_obj (win32.CDispatch): Objeto pasta do Outlook.
        """
        try:
            # Conecta ao cliente Outlook
            outlook: win32.CDispatch = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

            # Obtém a lista de pastas raiz
            root_folders = outlook.Folders.Item(1)

            # Itera sobre as pastas para encontrar a pasta especificada
            for pasta_email_obj in root_folders.Folders:
                if pasta_email_obj.Name == pasta_email:
                    return pasta_email_obj

            Log.write_log(f"Pasta '{pasta_email}' não encontrada no Outlook.", log_level=LogLevel.ERROR)
            raise Exception(f"Pasta '{pasta_email}' não encontrada no Outlook.")
        
        except Exception as err:
            Log.write_log(f"Erro ao obter pasta: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err
    
    @classmethod
    def _filtrar_emails(cls, emails_outlook, assunto:str, remetente:str, palavra_corpo:str, email_lido: bool) -> list:
        """
        Aplica filtros nos emails capturados.

        Parâmetros:
        - emails_outlook: Lista de emails.
        - assunto (str): String contida no campo 'assunto' (opcional).
        - remetente (str): Email do remetente para filtragem (opcional).
        - palavra_corpo (str): Palavra a ser buscada no corpo do email (opcional).
        - email_lido (bool): Define o status de leitura (True para lidos, False para não lidos, None para todos).

        Retorna:
        - emails_lista (list): Lista de emails filtrados.
        """
        emails_lista = []
        try:
            for email in emails_outlook: 
                
                # Verifica o status de leitura se foi especificado
                if email_lido is not None and email.UnRead != (not email_lido):
                    continue
                
                # Filtra por assunto se especificado
                if assunto and assunto.lower() not in email.Subject.lower():
                    continue
                
                # Filtra por remetente se especificado
                if remetente and remetente.lower() not in email.Sender.GetExchangeUser().PrimarySmtpAddress.lower():
                    continue
                
                # Filtra por palavra no corpo se especificado
                if palavra_corpo and palavra_corpo.lower() not in email.Body.lower():
                    continue

                # Converte data do email para o formato '%d/%m/%Y %H:%M:%S'
                data_recebida = email.ReceivedTime
                data_formatada = data_recebida.strftime("%d/%m/%Y %H:%M:%S")

                # Detalhes do Email
                detalhe_email = {
                    'assunto': email.Subject,
                    'remetente': email.Sender.GetExchangeUser().PrimarySmtpAddress,
                    'data_recebida': data_formatada,
                    'corpo_email': email.Body
                }
                
                emails_lista.append(detalhe_email)

        except Exception as err:
            Log.write_log(f"Erro ao filtrar emails: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err
        
        return emails_lista

    @classmethod
    def _baixar_anexos(cls, email, extensao_anexo:str, caminho_pasta_anexos:str) -> None:
        """
        Baixa anexos do email.

        Parâmetros:
        - email: Objeto de email.
        - extensao_anexo (str): Extensão do anexo a ser baixado (opcional).
        - caminho_pasta_anexos (str): Caminho para salvar anexos.

        Retorna:
        """
        try:
            # Baixar todos os anexos ou filtrar pela extensão especificada
            for anexo in email.Attachments:
                # Baixar tudo ou filtrar pela extensão
                if not extensao_anexo or anexo.FileName.endswith(extensao_anexo):
                    anexo.SaveASFile(os.path.join(caminho_pasta_anexos, anexo.FileName))
        
        except Exception as err:
            Log.write_log(f"Erro ao baixar anexos: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err

    @classmethod
    def capturar_emails(cls, pasta_email:str, assunto:str = None, 
                       remetente:str = None, palavra_corpo:str = None, 
                       baixar_anexos: bool = False, extensao_anexo:str = None, 
                       caminho_pasta_anexos:str = None, email_lido: bool = None) -> list:
        
        """
        Captura emails do Outlook, aplicando filtros por pasta, assunto, remetente e corpo.

        Parâmetros:
        - pasta_email (str): Nome da pasta para filtragem.
        - assunto (str): String contida no campo 'assunto' (opcional).
        - remetente (str): Email do remetente para filtragem (opcional).
        - palavra_corpo (str): Palavra a ser buscada no corpo do email (opcional).
        - baixar_anexos (bool): Indica se os anexos devem ser baixados.
        - extensao_anexo (str): Extensão do anexo a ser baixado (opcional).
        - caminho_pasta_anexos (str): Caminho para salvar anexos.
        - email_lido (bool): Define o status de leitura (True para lidos, False para não lidos, None para todos).

        Retorna:
        - emails_lista (list): Lista de emails capturados.
        """
        try:
            Log.write_log("Iniciando captura de email Outlook")

            pasta = cls._obter_pasta(pasta_email)

            # Obtém os emails da pasta e ordena por data de recebimento (mais recente primeiro)
            emails_outlook = pasta.Items
            emails_outlook.Sort("[ReceivedTime]", True)

            # Aplica os filtros nos emails capturados
            emails_lista = cls._filtrar_emails(emails_outlook, assunto, remetente, palavra_corpo, email_lido)

            # Verifica se os anexos devem ser baixados
            if baixar_anexos:
                for email in emails_outlook:
                    cls._baixar_anexos(email, extensao_anexo, caminho_pasta_anexos)

            Log.write_log("Finalizando captura de email Outlook")

            return emails_lista

        except Exception as err:
            Log.write_log(f"Erro ao capturar emails: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err