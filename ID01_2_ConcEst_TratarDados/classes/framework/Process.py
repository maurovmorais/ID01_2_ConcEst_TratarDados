# Imports dos módulos internos do projeto
# Carrega o InitAllSettingssSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException
from ID01_2_ConcEst_TratarDados.classes.framework.GetTransaction import GetTransaction

# Imports dos pacotes externos
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep


class Process:
    """
    Classe responsável pelo processamento principal.

    Parâmetros:
    
    Retorna:
    """
    _config = InitAllSettings.config
    
    
    @classmethod
    def execute(cls):
        """
        Método principal para execução do código.

        """
        Log.write_log('Process Started')

        adquirente = GetTransaction.queue_item['info_adicionais']['Adquirentes']
   
        if adquirente == 'ConectCar':
            Log.write_log(f'Processando adquirente: {adquirente}')
     
        elif adquirente == 'Cielo':
            Log.write_log(f'Processando adquirente: {adquirente}')
      
        elif adquirente == 'Greenpass':
            Log.write_log(f'Processando adquirente: {adquirente}')
     
        elif adquirente == 'SemParar':
            Log.write_log(f'Processando adquirente: {adquirente}')
     
        elif adquirente == 'Bradesco':
            Log.write_log(f'Processando adquirente: {adquirente}')

        elif adquirente == 'Veloe':
            Log.write_log(f'Processando adquirente: {adquirente}')

        else:
            Log.write_log(f'Adquirente: {adquirente} não encontrado')


        Log.write_log('Process Finished')
