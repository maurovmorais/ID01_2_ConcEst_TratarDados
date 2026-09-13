# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.CredentialManager import CredentialManager
# Imports dos pacotes externos
from datetime import datetime
from bs4 import BeautifulSoup
import requests
import msal
import os

class GetEmailOutlookApi:
    """
    Classe responsável por capturar emails Outlook usando a API do Microsoft Graph, aplicando filtros opcionais por pasta, assunto e remetente.

    Parâmetros:

    Retorna:
    """
    
    _config = InitAllSettings.config
    _timeout_request = int(_config.get('MaxTimeoutRequests', 30))

    @classmethod
    def configure(cls):
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
            Log.write_log(f"Erro na configuração: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err
    
    @classmethod    
    def _conectar_api(cls):
        """
        Conecta na API do Microsoft Graph.

        Parâmetros:

        Retorna:
        """
        try:
            # Cria uma instância de aplicação confidencial MSAL para autenticação
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
            # Conecta na API e obtém o token de acesso usando as credenciais do usuário
            cls._conectar_api()
            result = cls.app.acquire_token_by_username_password(
                username=cls.usuario,
                password=cls.senha,
                scopes=['https://graph.microsoft.com/.default']
            )

            # Retorna o token de acesso
            token_acesso = result.get("access_token")
            return token_acesso
        except Exception as err:
            Log.write_log(f"Erro ao obter token: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err
    
    @classmethod
    def _obter_id_pasta(cls, nome_pasta:str) -> str:
        """
        Obtém o ID de uma pasta pelo nome.

        Parâmetros:
        - nome_pasta (str): Nome da pasta.

        Retorna:
        - id_pasta (str): ID da pasta.
        """
        try:
            # Define os headers com o token de autorização
            headers = {"Authorization": f"Bearer {cls.token_acesso}"}
            response = requests.get("https://graph.microsoft.com/v1.0/me/mailFolders", headers=headers, timeout=cls._timeout_request)

            if response.status_code == 200:
                # Processa a resposta e procura pela pasta especificada
                pastas = response.json().get('value', [])
                for pasta in pastas:
                    if pasta['displayName'].lower() == nome_pasta.lower():
                        id_pasta = pasta['id']
                        return id_pasta
                raise Exception(f"Pasta '{nome_pasta}' não encontrada.")
            else:
                raise Exception("Falha ao recuperar IDs das pastas. " + response.text)
        except Exception as err:
            Log.write_log(f"Erro ao obter ID da pasta: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err
    
    @classmethod
    def _baixar_anexos(cls, caminho_pasta_anexos:str, extensao_anexo:str) -> list:
        """
        Baixa anexos de um email.

        Parâmetros:
        - caminho_pasta_anexos (str): Caminho para salvar anexos.
        - extensao_anexo (str): Extensão do anexo a ser baixado. Se None, baixa todos.

        Retorna:
        - caminho_anexo (list): Lista de caminhos dos anexos baixados.
        """
        try:
            headers = {"Authorization": f"Bearer {cls.token_acesso}"}
            endpoint_anexos = f"https://graph.microsoft.com/v1.0/me/messages/{cls.id_mensagem}/attachments"
            response = requests.get(endpoint_anexos, headers=headers, timeout=cls._timeout_request)
            
            if response.status_code == 200:
                anexos = response.json().get('value', [])
                caminho_anexo = []

                for anexo in anexos:
                    nome_arquivo = anexo.get('name')

                    # Verifica se deve baixar o anexo com a extensão especificada
                    if extensao_anexo is None or nome_arquivo.endswith(extensao_anexo):
                        dados_anexo = requests.get(f"{endpoint_anexos}/{anexo['id']}/$value", headers=headers, timeout=cls._timeout_request)
                        
                        if dados_anexo.status_code == 200:
                            # Salva o anexo no caminho especificado
                            caminho_arquivo = os.path.join(caminho_pasta_anexos, nome_arquivo)
                            with open(caminho_arquivo, 'wb') as arquivo:
                                arquivo.write(dados_anexo.content)
                            caminho_anexo.append(caminho_arquivo)

                return caminho_anexo
            else:
                Log.write_log("Falha ao recuperar anexos.")
                return []
        except Exception as err:
            Log.write_log(f"Erro ao baixar anexos: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            return []
    
    @classmethod
    def _converter_html_para_texto(cls, html_corpo_email:str) -> str:
        """
        Converte uma string HTML em texto, removendo todas as tags HTML.

        Parâmetros:
        - html_corpo_email (str): String contendo o corpo do email em formato HTML.

        Retorna:
        - corpo_email (str): String contendo apenas o texto extraído do HTML, sem tags.
        """
        try:
            bs4_soup = BeautifulSoup(html_corpo_email, 'html.parser')
            corpo_email = bs4_soup.get_text()

            return corpo_email
        except Exception as err:
            Log.write_log(f"Erro ao converter HTML para texto: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err


    @classmethod
    def capturar_emails(cls, pasta_email:str, assunto:str=None, remetente:str=None, 
                       palavra_corpo:str=None, baixar_anexos:bool =False, extensao_anexo=None,
                       caminho_pasta_anexos:str=None, email_lido:bool =None) -> list:
        """
        Captura emails com base nos filtros fornecidos.

        Parâmetros:
        - pasta_email (str): Nome da pasta de emails.
        - assunto (str): Filtro por assunto (opcional).
        - remetente (str): Filtro por remetente (opcional).
        - palavra_corpo (str): Palavra a ser buscada no corpo do email (opcional).
        - baixar_anexos (bool): Indica se os anexos devem ser baixados.
        - extensao_anexo (str): Extensão do anexo a ser baixado (opcional).
        - caminho_pasta_anexos (str): Caminho para salvar anexos.
        - email_lido (bool): Define o status de leitura (True para lidos, False para não lidos, None para todos).

        Retorna:
        - emails_lista (list): Lista de dicionários contendo detalhes dos emails capturados.
        """
        try:
            Log.write_log("Iniciando captura de email Outlook")

            cls.configure()
            cls.token_acesso = cls._obter_token()
            
            if cls.token_acesso:
                # Obtém o ID da pasta especificada
                id_pasta = cls._obter_id_pasta(pasta_email)
                headers = {"Authorization": f"Bearer {cls.token_acesso}"}
                filtros = []
                
                # Adiciona filtros conforme especificado
                if assunto:
                    filtros.append(f"contains(subject, '{assunto.lower()}')")
                
                if remetente:
                    filtros.append(f"from/emailAddress/address eq '{remetente.lower()}'")
                
                if email_lido is not None:
                    filtros.append(f"isRead eq {str(email_lido).lower()}")
                
                consulta_filtro = ' and '.join(filtros)

                params = {
                    '$filter': consulta_filtro,
                    '$top': 50,
                }
                
                # Faz uma requisição GET para obter mensagens de email
                endpoint = f"https://graph.microsoft.com/v1.0/me/mailFolders/{id_pasta}/messages"
                response = requests.get(endpoint, headers=headers, params=params, timeout=cls._timeout_request)
                
                if response.status_code == 200:
                    mensagens = response.json().get('value', [])
                    emails_lista = []
                    
                    for mensagem in mensagens:
                        # Verifica se a palavra do corpo do email está presente
                        if palavra_corpo and palavra_corpo.lower() not in mensagem['body']['content'].lower():
                            continue
                        
                        # Converte o corpo do email de HTML para texto
                        corpo_email = cls._converter_html_para_texto(html_corpo_email=mensagem['body']['content'])
                        
                        # Converte a string para um objeto datetime
                        data_recebida = datetime.strptime(mensagem['receivedDateTime'], "%Y-%m-%dT%H:%M:%SZ")

                        # Converte o objeto datetime para o formato '%d/%m/%Y %H:%M:%S'
                        data_recebida_str = data_recebida.strftime("%d/%m/%Y %H:%M:%S")

                        # Dicionário com os detalhes do email
                        detalhe_email = {
                            'assunto': mensagem['subject'],
                            'remetente': mensagem['from']['emailAddress']['address'],
                            'data_recebida': data_recebida_str,
                            'corpo_email': corpo_email.replace("\n", "")
                        }
                        
                        # Verifica se os anexos devem ser baixados e adiciona ao dicionário
                        if baixar_anexos:
                            cls.id_mensagem = mensagem['id']
                            detalhe_email['anexos'] = cls._baixar_anexos(caminho_pasta_anexos, extensao_anexo)
                        
                        emails_lista.append(detalhe_email)
                    
                    Log.write_log("Finalizando captura de email Outlook")

                    return emails_lista
                else:
                    Log.write_log(f"Falha ao recuperar mensagens: " + response.text, log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
                    raise Exception(f"Falha ao recuperar mensagens: " + response.text)
            else: 
                Log.write_log("Falha ao adquirir o token de acesso.", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
                raise Exception("Falha ao adquirir o token de acesso.")
        except Exception as err:
            Log.write_log(f"Erro ao capturar emails: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err