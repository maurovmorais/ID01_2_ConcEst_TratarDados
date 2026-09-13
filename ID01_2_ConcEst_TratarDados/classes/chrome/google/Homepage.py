# Imports dos módulos internos do projeto
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings

# Imports dos pacotes externos
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep


class GoogleHomepage:
    """
    Classe criada para manipular os elementos da página inicial do site do google

    Parâmetros:

    Retorna:
    """

    @staticmethod
    def search_something(item_de_pesquisa: str):
        """
        Realiza a pesquisa na homepage do google

        Parâmetros:
        - item_de_pesquisa (str): O que será pesquisado no google.

        Retorna:
        """
        driver = InitAllSettings.web_driver

        elm_campo_pesquisa = driver.find_element(By.NAME, "q")
        elm_campo_pesquisa.clear()
        elm_campo_pesquisa.send_keys(item_de_pesquisa)
        sleep(1)
        elm_campo_pesquisa.send_keys(Keys.ENTER)

    @staticmethod
    def open_google_website(url_google: str):
        """
        Realiza a abertura do site do google

        Parâmetros:
        - url_google (str): Url do site a ser inicializado.

        Retorna:
        """
        driver = InitAllSettings.web_driver
        driver.get(f"https://{url_google}" if not url_google.startswith("http") else url_google)

    @staticmethod
    def close_google_website():
        """
        Realiza a finalização do site do google

        Parâmetros:

        Retorna:
        """
        driver = InitAllSettings.web_driver
        driver.close()
