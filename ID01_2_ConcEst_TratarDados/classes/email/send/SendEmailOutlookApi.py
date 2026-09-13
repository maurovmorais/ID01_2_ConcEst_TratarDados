# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.CredentialManager import CredentialManager
# Imports dos pacotes externos
import requests
import msal
import os
import base64

class SendEmailOutlookApi:
    """
    Classe responsável por enviar emails usando a API do Microsoft Graph.
    
    Parâmetros: 
        
    Retorna:
    """

    _config = InitAllSettings.config
    _timeout_request = int(_config.get('MaxTimeoutRequests', 30))

    @classmethod
    def configure_email(cls):
        """
        Inicializa as variáveis de configuração.

        Parâmetros:
        
        Retorna:
        """
        try:
            cls.tenant_id = cls._config['TenantId']          # Tenant ID registrada no Azure
            cls.client_id = cls._config['ClientId']          # Client ID registrada no Azure
            cls.client_secret = CredentialManager.get_credential(cls._config['CRED_KEY_CLIENT_SECRET'], cls._config['CRED_LABEL_CLIENT_SECRET'])   # Client Secret registrada no Azure
            cls.usuario = cls._config['EmailCredenciais']    # Nome de usuário para autenticação
            cls.senha = CredentialManager.get_credential(cls._config['CRED_KEY_SENHA'], cls._config['CRED_LABEL_EMAIL'])   # Senha do usuário para autenticação
        except Exception as err:
            Log.write_log(f"Erro na configuração: {str(err)}", log_level=LogLevel.ERROR)
            raise err

    @classmethod    
    def _conectar_api(cls):
        """
        Conecta na API do Microsoft Graph.

        Parâmetros:
        
        Retorna:
        """
        try:
            cls.app = msal.ConfidentialClientApplication(
                cls.client_id,
                authority=f"https://login.microsoftonline.com/{cls.tenant_id}",
                client_credential=cls.client_secret,
            )
        except Exception as err:
            Log.write_log(f"Erro ao conectar API: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err

    @classmethod
    def _obter_token(cls) -> str:
        """
        Captura o token de acesso.

        Parâmetros: 

        Retorna:
        - token_acesso (str): Token de acesso.
        """
        try:
            cls._conectar_api()
            result = cls.app.acquire_token_by_username_password(
                username=cls.usuario,
                password=cls.senha,
                scopes=['https://graph.microsoft.com/.default']
            )

            token_acesso = result.get("access_token")
            
            return token_acesso
        except Exception as err:
            Log.write_log(f"Erro ao obter token: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err

    @classmethod
    def send_email(cls, para:str, assunto:str, corpo_email:str, cc:str=None, anexos:list=None):
        """
        Envia um email usando a API do Microsoft Graph.

        Parâmetros:
        - para (str): Endereço de email do destinatário.
        - assunto (str): Assunto do email.
        - corpo_email (str): Corpo do email em formato HTML.
        - cc (str): Endereço de email para cópia (opcional).
        - anexos (list): Lista de caminhos para arquivos anexos (opcional).

        Retorna:
        """
        try:
            cls.token_acesso = cls._obter_token()

            if cls.token_acesso:
                headers = {
                    "Authorization": f"Bearer {cls.token_acesso}",
                    "Content-Type": "application/json"
                }

                # Estrutura do email
                email = {
                    "message": {
                        "subject": assunto,
                        "body": {
                            "contentType": "HTML",
                            "content": corpo_email
                        },
                        "toRecipients": [
                            {
                                "emailAddress": {
                                    "address": para
                                }
                            }
                        ],
                        "ccRecipients": [
                            {
                                "emailAddress": {
                                    "address": cc
                                }
                            }
                        ] if cc else [],
                        "attachments": []
                    }
                }

                # Adiciona anexos se existir
                if anexos:
                    for caminho_anexo in anexos:
                        with open(caminho_anexo, "rb") as arquivo:
                            conteudo_anexo = base64.b64encode(arquivo.read()).decode('utf-8')
                            anexo = {
                                "@odata.type": "#microsoft.graph.fileAttachment",
                                "name": os.path.basename(caminho_anexo),
                                "contentBytes": conteudo_anexo
                            }
                            email["message"]["attachments"].append(anexo)

                # Envia o email
                response = requests.post(f"https://graph.microsoft.com/v1.0/users/{cls.usuario}/sendMail", headers=headers, json=email, timeout=cls._timeout_request)

                if response.status_code == 202:
                    Log.write_log("Email enviado com sucesso.")
                else:
                    Log.write_log(f"Falha ao enviar email: {response.text}")
                    raise Exception(f"Falha ao enviar email: {response.text}")
            else:
                Log.write_log("Falha ao adquirir o token de acesso.", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
                raise Exception("Falha ao adquirir o token de acesso.")
        
        except Exception as err:
            Log.write_log(f"Erro ao enviar email: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao enviar email: {err}")