# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.CredentialManager import CredentialManager
# Imports dos pacotes externos
import imaplib
import email
import os
from email.header import decode_header
from datetime import datetime

class GetEmail:
    """
    Classe responsável por capturar emails de uma conta IMAP, aplicando filtros opcionais por pasta, assunto e remetente.

    Parâmetros:
    
    Retorna:
    """

    _config = InitAllSettings.config

    @classmethod
    def configure(cls, servidor:str):
        """
        Inicializa as variáveis de configuração e conecta ao servidor IMAP.

        Parâmetros:
        - servidor (str): Endereço do servidor IMAP.

        Retorna:
        """
        cls.usuario = cls._config['EmailCredenciais']
        cls.senha = CredentialManager.get_credential(cls._config['CRED_KEY_SENHA'], cls._config['CRED_LABEL_EMAIL']) # Ajustar config com a Key e Label conforme necessário.
        cls.servidor = servidor
        cls.cls_conexao = None

        try:
            # Conecta ao servidor IMAP
            cls.cls_conexao = imaplib.IMAP4_SSL(cls.servidor)
            cls.cls_conexao.login(cls.usuario, cls.senha)
        
        except Exception as err:
            Log.write_log(mensagem_log="Erro ao conectar ao servidor IMAP: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err

    @classmethod
    def _desconectar(cls):
        """
        Fecha a conexão com o servidor IMAP.

        Parâmetros:

        Retorna:
        """
        try:
            if cls.cls_conexao:
                # Fecha e finaliza a sessão do IMAP
                cls.cls_conexao.close()
                cls.cls_conexao.logout()
        except Exception as err:
            Log.write_log(mensagem_log="Erro ao desconectar IMAP: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err
        
    @classmethod
    def _decodificar_assunto_email(cls, assunto_email:str) -> str:
        """
        Decodifica assuntos de email codificados.

        Parâmetros:
        - assunto_email (str): Assunto de email codificado.

        Retorna:
        - assunto_email_decodificado (str): Assunto de email decodificado.
        """
        try:
            # Decodifica o assunto do email
            assunto_email_decodificado, charset = decode_header(assunto_email)[0]
            
            if isinstance(assunto_email_decodificado, bytes):
                # Decodifica bytes para string usando o charset fornecido ou utf-8 por padrão
                return assunto_email_decodificado.decode(charset or 'utf-8')
            
            return assunto_email_decodificado
        except Exception as err:
            Log.write_log(mensagem_log="Erro ao decodificar Assunto Email: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err
        
    @classmethod
    def _obter_corpo_email(cls, msg_email:email.message.Message) -> str:
        """
        Retorna o corpo do email como texto simples.

        Parâmetros:
        - msg_email (Message): Objeto de mensagem do email.

        Retorna:
        - corpo_email (str): Texto do corpo do Email.
        """
        try:
            if msg_email.is_multipart():
                # Itera sobre as partes do email
                for parte in msg_email.walk():
                    if parte.get_content_type() == "text/plain":
                        corpo_email = parte.get_payload(decode=True).decode()
                        
                        # Retorna o texto do corpo do email
                        return corpo_email
            else:
                # Para emails não-multipart, retorna o payload direto
                corpo_email = msg_email.get_payload(decode=True).decode()
                
                return corpo_email
        except Exception as err:
            Log.write_log(mensagem_log="Erro ao obter Corpo Email: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err

    @classmethod
    def _baixar_anexos(cls, msg_email:email.message.Message, caminho_pasta_anexos:str, extensao_anexo:str) -> list:
        """
        Baixa anexos de um email com a extensão especificada.

        Parâmetros:
        - msg_email (Message): Objeto de mensagem do email.
        - caminho_pasta_anexos (str): Caminho para salvar anexos.
        - extensao_anexo (str): Extensão do anexo a ser baixado (opcional).

        Retorna:
        - anexos (list): Lista de caminhos dos anexos baixados.
        """
        try:
            anexos = []
            for parte in msg_email.walk():
                if parte.get_content_maintype() == 'multipart' or parte.get('Content-Disposition') is None:
                    continue
                
                # Obtém o nome do arquivo e verifica se deve ser baixado
                nome_arquivo = parte.get_filename()
                if nome_arquivo and (not extensao_anexo or nome_arquivo.endswith(extensao_anexo)):
                    caminho_arquivo = os.path.join(caminho_pasta_anexos, nome_arquivo)
                    
                    # Salva o anexo no caminho especificado
                    with open(caminho_arquivo, 'wb') as arquivo:
                        arquivo.write(parte.get_payload(decode=True))
                    anexos.append(caminho_arquivo)

            return anexos
        except Exception as err:
            Log.write_log(mensagem_log="Erro ao baixar anexos Email: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise err

    @classmethod
    def capturar_emails(cls, pasta_email:str, assunto:str=None, remetente:str=None, 
                        palavra_corpo:str=None, baixar_anexos:bool=False,
                        extensao_anexo:str=None, caminho_pasta_anexos:str=None, 
                        email_lido:bool=None) -> list:
        """
        Captura emails com base nos filtros fornecidos.

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
        - emails_lista (list): Lista de dicionários contendo detalhes dos emails capturados.
        """
        try:
            Log.write_log("Iniciando captura de email")
            
            # Seleciona a pasta de emails especificada
            cls.cls_conexao.select(pasta_email)
            criterios_pesquisa = []

            # Adiciona filtros de pesquisa conforme especificado
            if email_lido is not None:
                criterios_pesquisa.append('(SEEN)' if email_lido else '(UNSEEN)')
            
            if assunto:
                criterios_pesquisa.append(f'(HEADER Subject "{assunto}")')
            
            if remetente:
                criterios_pesquisa.append(f'(FROM "{remetente}")')

            # Se nenhum critério for especificado, busca todos os emails
            if not criterios_pesquisa:
                criterios_pesquisa = ['ALL']

            resultado, dados = cls.cls_conexao.search(None, *criterios_pesquisa)
            ids_emails = dados[0].split()

            emails_lista = []
            for id_email in ids_emails:
                resultado, dados_msg = cls.cls_conexao.fetch(id_email, '(RFC822)')
                email_msg = email.message_from_bytes(dados_msg[0][1])

                # Obtém o corpo do email
                corpo_email = cls._obter_corpo_email(email_msg)
                if palavra_corpo and palavra_corpo.lower() not in corpo_email.lower():
                    continue
                
                # Convertendo Data Extraída do email para o padrão '%d/%m/%Y %H:%M:%S'
                data_recebida = datetime.strptime(email_msg.get('Date'), '%a, %d %b %Y %H:%M:%S %z')
                data_recebida_str = data_recebida.strftime('%d/%m/%Y %H:%M:%S')
                
                detalhes_email = {
                    'assunto': cls._decodificar_assunto_email(email_msg['Subject']),
                    'remetente': email_msg.get('From'),
                    'data_recebida': data_recebida_str,
                    'corpo_email': corpo_email.replace("\r", "").replace("\n", "")
                }

                # Verifica se os anexos devem ser baixados e adiciona ao dicionário
                if baixar_anexos and caminho_pasta_anexos:
                    detalhes_email['anexos'] = cls._baixar_anexos(email_msg, caminho_pasta_anexos, extensao_anexo)

                emails_lista.append(detalhes_email)

            Log.write_log("Finalizando captura de email")

            return emails_lista

        except Exception as err:
            Log.write_log(mensagem_log="Erro ao capturar emails: " + str(err), log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao capturar emails: {err}")