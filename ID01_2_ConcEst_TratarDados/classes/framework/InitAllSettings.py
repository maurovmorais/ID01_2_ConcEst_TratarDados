# Imports dos módulos internos do projeto
# Carrega o InitAllSettings - Precisa ser o primeiro a ser carregado

# Imports dos pacotes externos
import os
import uuid
from enum import Enum
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pathlib import Path
from datetime import datetime
from webdriver_manager import chrome, firefox, microsoft
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium import webdriver
from tkinter import *
from tkinter import messagebox

# Para durante a execução conseguir capturar um arquivo no resources é necessário usar dessa maneira
ROOT_DIR = Path(__file__).parent.parent.parent


class Browser(Enum):
    """
    Classe enum, usada para indicar qual navegador será utilizado pelo Selenium.

    Parâmetros:

    Retorna:
    """

    CHROME = "chrome"
    EDGE = "edge"
    FIREFOX = "firefox"


class InitAllSettings:
    """
    Classe para carregar todas as configurações necessárias.

    Parâmetros:
    
    Retorna:
    """
    config:dict = None
    
    @staticmethod
    def load_config() -> dict:
        """
        Carrega o arquivo de configuração Config.xlsx e retorna um dicionário com as configurações.

        Parâmetros:

        Retorna:
        - dict: dicionário com as configurações lidas do arquivo Config.xlsx.
        """

        wbk_config:Workbook = load_workbook(filename=os.path.join(ROOT_DIR.__str__() , r"resources\config\Config.xlsx"))
        wsht_settings:Worksheet = wbk_config.get_sheet_by_name("Settings")
        wsht_constants:Worksheet = wbk_config.get_sheet_by_name("Constants")
        wsht_credentials:Worksheet = wbk_config.get_sheet_by_name("Credentials")
        wsht_assets:Worksheet = wbk_config.get_sheet_by_name("Assets")

        #Iniciando o dicionário
        config = dict()

        #Loop adicionando dados da aba settings para o dicionário
        total_rows = wsht_settings.max_row
        for row_number in range(1, total_rows):
            gnc_dict_key = wsht_settings["A" + (row_number+1).__str__()].value
            gnc_dict_obj = wsht_settings["B" + (row_number+1).__str__()].value
            #Apenas inclui no dicionário se as duas colunas não forem nulas
            if(gnc_dict_key is not None and gnc_dict_obj is not None): config[gnc_dict_key] = gnc_dict_obj 

        #Mesmo loop, agora com constants
        total_rows = wsht_constants.max_row
        for row_number in range(1, total_rows):
            gnc_dict_key = wsht_constants["A" + (row_number+1).__str__()].value
            gnc_dict_obj = wsht_constants["B" + (row_number+1).__str__()].value
            #Apenas inclui no dicionário se as duas colunas não forem nulas
            if(gnc_dict_key is not None and gnc_dict_obj is not None): config[gnc_dict_key] = gnc_dict_obj 

        #Mesmo loop, agora com credentials
        total_rows = wsht_credentials.max_row
        for row_number in range(1, total_rows):
            gnc_dict_key = wsht_credentials["A" + (row_number+1).__str__()].value
            gnc_dict_obj = wsht_credentials["B" + (row_number+1).__str__()].value
            #Apenas inclui no dicionário se as duas colunas não forem nulas
            if(gnc_dict_key is not None and gnc_dict_obj is not None): config[gnc_dict_key] = gnc_dict_obj 

        #Mesmo loop, agora com assets
        total_rows = wsht_assets.max_row
        for row_number in range(1, total_rows):
            gnc_dict_key = wsht_assets["A" + (row_number+1).__str__()].value
            gnc_dict_name = wsht_assets["B" + (row_number+1).__str__()].value
            gnc_dict_folder = wsht_assets["C" + (row_number+1).__str__()].value
            #Apenas inclui no dicionário se as duas colunas não forem nulas
            if(gnc_dict_key is not None and gnc_dict_folder is not None and gnc_dict_name is not None): 
                config[gnc_dict_key] = {
                    "folder": gnc_dict_folder,
                    "asset": gnc_dict_name
                }

        #Retorna o objeto dicionario para ser usado no código
            
        return config
    
    @classmethod
    def _build(cls):
        """
        Realiza a inicialização e verificação de variaveis de configuração.
        
        Parâmetros:
           
        Retorna:
        """
        cls.datahora_inicio_exec:datetime = datetime.now()
        cls.datahora_inicio_exec:str = cls.datahora_inicio_exec.strftime("%d/%m/%Y %H:%M:%S")
        cls.datahora_fim_exec:datetime = None
        cls.datahora_fim_exec:str = ''
        cls.guid_execucao = str(uuid.uuid4())
        cls.browser_escolhido:Browser = None
        cls.headless:bool = None

        # Caminhos hardcode utilizados em templates de email
        cls.caminho_template_email_inicio = os.path.join(ROOT_DIR , "resources\\templates\\Email_Inicio.txt")
        cls.caminho_template_email_final = os.path.join(ROOT_DIR , "resources\\templates\\Email_Final.txt")
        cls.caminho_template_email_erro_encontrado = os.path.join(ROOT_DIR , "resources\\templates\\Email_ErroEncontrado.txt")

        # Caminho hardcode utilizados em templates e scripts para relatorio final
        cls.sqlite_caminho_bd_analit_sint = cls.config["CaminhoBancoSqlite"] if(cls.config.__contains__("CaminhoBancoSqlite")) else os.path.join(ROOT_DIR, "resources\\sqlite\\banco_dados.db")
        cls.caminho_template_excel_analitico = os.path.join(ROOT_DIR, "resources\\templates\\Relatorio_Analitico.xlsx")
        cls.caminho_template_excel_sintetico = os.path.join(ROOT_DIR, "resources\\templates\\Relatorio_Sintetico.xlsx")
        cls.caminho_script_select_dados_analitico = os.path.join(ROOT_DIR,r'resources\scripts\analitico_sintetico\Script_Select_Analitico.sql')
        cls.caminho_script_select_dados_sintetico = os.path.join(ROOT_DIR,r'resources\scripts\analitico_sintetico\Script_Select_Sintetico.sql')
        cls.caminho_pasta_saida_rel_sint_anali = cls.config["CaminhoPastaRelatorios"] if(cls.config.__contains__("CaminhoPastaRelatorios")) else os.path.join(ROOT_DIR,r'classes\relatorios\output')
        cls.caminho_script_update_dados_execucao = os.path.join(ROOT_DIR,r'resources\scripts\analitico_sintetico\SCRIPT_UPDATE_DADOSEXECUCAO.SQL')
        cls.caminho_script_select_captura_qtd_itens = os.path.join(ROOT_DIR,r'resources\scripts\analitico_sintetico\SCRIPT_SELECT_CAPTURAQTDITENS.SQL')

        # Caminho RobotStream
        cls.caminho_executavel_robot_stream = os.path.join(ROOT_DIR,r'resources\robot_stream\StreamRobotScreen.exe')

        #Variáveis contadores já levadas em conta no framework
        cls.qtde_itens_processados:int = 0
        cls.qtde_itens_app_exception:int = 0
        cls.qtde_itens_consecutive_exceptions:int = 0
        cls.qtde_itens_business_exception:int = 0
        cls.qtde_itens_sucesso:int = 0
        
        #Variáveis que precisam ser levadas em conta pelo desenvolvedor (somar contadores e indicar uso ou não)
        cls.usa_captcha:bool = False
        cls.usa_ocr:bool = False
        cls.usa_api:bool = False
        cls.qtde_captcha:int = 0
        cls.qtde_ocr:int = 0
        cls.qtde_api:int = 0
        cls.qtd_itens_a_processar_ini_exec:int = None

        # Instância do WebDriver do Selenium utilizada pela automação (definida em initiate_web_manipulator)
        cls.web_driver = None

        # Variaveis de controle de excecao
        cls.exception_initialization = None
        cls.exception_process = None
        cls.caminho_screenshot_erro_init = None

        # Verifica se foi preenchido no arquivo config sobre a fila de processamento, caso não, preenche com o nome padrao   
        cls.config["FilaProcessamento"] = cls.config["FilaProcessamento"] if(cls.config.__contains__("FilaProcessamento")) else "tbl_Fila_Processamento"

    @classmethod
    def initiate_web_manipulator(cls, headless:bool, browser_escolhido:Browser, pasta_download:str=None, widht_resolution:int=None, height_resolution:int=None):
        """
        Método que inicia a instância do WebDriver do Selenium utilizada para manipular os seletores web.

        !! #HACK IMPORTANTE: NECESSARIO SER CHAMADO APÓS UM CLOSE/QUIT DO DRIVER OU KILLPROCESS DO NAVEGADOR !!
        
        Parâmetros:
            - headless (bool): indica se quer que execute o bot em modo background ou não
            - browser_escolhido (Browser): indica qual navegador que será utilizado
            - pasta_download (str): Caminho da pasta onde ficaram os downloads
            - widht_resolution (int): Largura da resolução do navegador
            - height_resolution (int): Altura da resolução do navegador
        Retorna:
        """
        cls.browser_escolhido = browser_escolhido
        cls.headless = headless

        if browser_escolhido == Browser.CHROME:
            options = webdriver.ChromeOptions()
            prefs = {
                # Impede que o Chrome exiba a janela de confirmação antes de iniciar um download.
                "download.prompt_for_download": False,
                
                # Garante que os arquivos sejam salvos no diretório certo sem popups ou falhas silenciosas.
                "download.directory_upgrade": True,
                
                # Garantindo que o Chrome continue protegendo contra arquivos maliciosos.
                "safebrowsing.enabled": True
            }

            if pasta_download:
                prefs["download.default_directory"] = pasta_download

            if headless:
                options.add_argument('--headless=new')

            options.add_experimental_option("prefs", prefs)

            # Desativa a flag AutomationControlled da engine Blink (usada pelo Chromium/Chrome). Torna o navegador menos detectável como "bot" — uma técnica de evasão de detecção por scripts de segurança.
            options.add_argument("--disable-blink-features=AutomationControlled")

            # Desativa o modo sandboxing de segurança do Chrome. Algumas execuções (principalmente em containers como Docker ou ambientes com permissões restritas) podem dar erro com o sandbox ativado.
            options.add_argument("--no-sandbox")

            # Impede o Chrome de usar o /dev/shm (Shared Memory) como espaço de memória temporária. 
            # Em contêineres Docker (ou sistemas com /dev/shm pequeno), o Chrome pode travar por falta de espaço compartilhado na RAM. 
            # Faz o Chrome usar o disco (/tmp) ao invés da RAM compartilhada, evitando crashes em ambientes limitados.
            options.add_argument("--disable-dev-shm-usage")

            # Manter só LOG FATAL
            options.add_argument("--log-level=3")  # 0=INFO, 1=WARNING, 2=LOG_ERROR, 3=LOG_FATAL

            chrome_service = ChromeService(chrome.ChromeDriverManager().install())
            cls.web_driver = webdriver.Chrome(service=chrome_service, options=options)

        elif browser_escolhido == Browser.EDGE:
            edge_service = EdgeService(microsoft.EdgeChromiumDriverManager().install())
            cls.web_driver = webdriver.Edge(service=edge_service)

        elif browser_escolhido == Browser.FIREFOX:
            firefox_service = FirefoxService(firefox.GeckoDriverManager().install())
            cls.web_driver = webdriver.Firefox(service=firefox_service)

        # Configura a resolução da tela do navegador, se especificada
        if widht_resolution and height_resolution and cls.web_driver:
            cls.web_driver.set_window_size(widht_resolution, height_resolution)

    @classmethod
    def save_screenshot(cls, path:str):
        """
        Captura uma screenshot da tela do desktop (não apenas do navegador), utilizada em cenários de erro
        para diagnóstico. Utiliza a biblioteca `pyautogui`.

        Parâmetros:
            - path (str): Caminho completo onde a screenshot será salva.

        Retorna:
        """
        import pyautogui
        pyautogui.screenshot(path)

    @classmethod
    def get_type_execution_using_vscode(cls):
        """
        Metodo que captura se a/o dev está executando pelo vscode em produção ou não, para ser contabilizado no RAAS.

        Parâmetros:
            
        Retorna:
        """

        # Verifica se está executando pelo vscode ou não
        if 'TERM_PROGRAM' in os.environ.keys() and os.environ['TERM_PROGRAM'] == 'vscode':
            cls.config['running_vscode'] = True

            # Identifica se é uma execução do tipo RAAS 
            if cls.config["RelatorioRAAS"].upper() == "SIM":

                # Coloca a janela da pergunta em primeiro nivel
                tkinter = Tk()
                tkinter.wm_attributes("-topmost", 1)
                tkinter.withdraw()

                # Exibe uma janela perguntando se Esta execução deverá ser contabilizada para cobrança do RAAS
                retorno_pergunta = messagebox.askquestion(title="Definição do tipo de execução", 
                                                                message="Esta execução deverá ser contabilizada para cobrança do RAAS?", 
                                                                parent=tkinter)

                # Finaliza as janelas tkinter
                tkinter.destroy()


                # Coloca a janela da pergunta em primeiro nivel
                tkinter = Tk()
                tkinter.wm_attributes("-topmost", 1)
                tkinter.withdraw()

                # Exibe uma janela confirmando a contabilização
                retorno_pergunta_confirmacao = messagebox.askquestion(title="Definição do tipo de execução", 
                                                                           message=f'Sua resposta foi {"Sim" if retorno_pergunta.upper() == "YES" else "Não"}. Você tem certeza?', 
                                                                           parent=tkinter)

                # Finaliza as janelas tkinter
                tkinter.destroy()


                # Caso a resposta seja que SIM para a contabilizacao e a confirmação da certeza seja sim, ira inserir o job id como 0
                # Caso a resposta seja que NÃO para a contabilizacao e a confirmação da certeza seja não, ira inserir o job id como 0
                # O id como 0 indicará para o robo que faz inserção dos dados no banco de cobrança do raas que é uma execução feita pelo vscode e nao tem job id
                if ((retorno_pergunta.upper() == "YES" and retorno_pergunta_confirmacao.upper() == "YES") or (retorno_pergunta.upper() == "NO" and retorno_pergunta_confirmacao.upper() == "NO")):
                    cls.config['job_id'] = 0
                elif retorno_pergunta.upper() == "NO":
                    cls.config['job_id'] = None
            else:
                cls.config['job_id'] = None
        else:
            cls.config['running_vscode'] = False
            cls.config['job_id'] = None
            print("A automação não está sendo executada pelo VSCode.")

    @classmethod
    def verify_required_fields_config(cls):
        """
        Metodo que verifica se os campos obrigatorios do Config.xlsx estão preenchidos.

        Parâmetros:
            
        Retorna:
        """

        # VERIFICA SE OS CAMPOS OBRIGATORIOS DO CONFIG ESTÃO PREENCHIDOS
        campos_importantes = ['NomeCliente','NomeProcesso','DescricaoProcesso','CaminhoExceptionScreenshots','CaminhoPastaRelatorios']
        if len(campos_importantes) > 0:
            for field in campos_importantes:
                if not field in cls.config.keys():
                    #Se for field de caminho de screenshot, mas foi informado que não precisa capturar nada, pula exception
                    if(field == "CaminhoExceptionScreenshots" and cls.config["CapturarScreenshot"] == "NÃO"):
                        continue

                    raise Exception(f"Por favor, preencha o campo {field} do arquivo Config.xlsx") 
                if cls.config[field] == "":
                    raise Exception(f"Por favor, preencha o campo {field} do arquivo Config.xlsx")


        
# Execute a função load_config para atribuir os valores do Config.xlsx ao atributo config
InitAllSettings.config = InitAllSettings.load_config()

InitAllSettings._build()

# VERIFICA SE OS CAMPOS OBRIGATORIOS DO CONFIG ESTÃO PREENCHIDOS
InitAllSettings.verify_required_fields_config()

# Identifica o tipo de execucao se é vscode ou nao
InitAllSettings.get_type_execution_using_vscode()
