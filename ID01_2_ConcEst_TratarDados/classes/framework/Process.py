# Imports dos módulos internos do projeto
# Carrega o InitAllSettingssSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException
from ID01_2_ConcEst_TratarDados.classes.framework.GetTransaction import GetTransaction
#FIXME Código Exemplo REMOVER
from ID01_2_ConcEst_TratarDados.classes.chrome.google.Homepage import GoogleHomepage

# Imports dos pacotes externos
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep




# Classe responsável pelo processamento principal, necessário preencher com o seu código no método execute
class Process:
    """
    Classe responsável pelo processamento principal.

    Parâmetros:
    
    Retorna:
    """
    _config = InitAllSettings.config
    
    
    #Parte principal do código, deve ser preenchida pelo desenvolvedor
    #Acesse o item a ser processado pelo queue_item
    @classmethod
    def execute(cls):
        """
        Método principal para execução do código.


        Parâmetros:


        Retorna:
        """
        cls.web_driver = InitAllSettings.web_driver
        Log.write_log('Process Started')
        #Informa valor no campo de pesquisa
        #GoogleHomepage.search_something(GetTransaction.queue_item['info_adicionais']['valor_pesquisa'])
        sleep(5)
 

        Log.write_log('Process Finished')
