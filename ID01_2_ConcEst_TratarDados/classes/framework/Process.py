# Imports dos módulos internos do projeto
# Carrega o InitAllSettingssSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException
from ID01_2_ConcEst_TratarDados.classes.framework.GetTransaction import GetTransaction
from ID01_2_ConcEst_TratarDados.classes.queue.QueueManagerPerformer import QueueManagerPerformer
from ID01_2_ConcEst_TratarDados.classes.sqlite.criar_fila_performer import criar_fila_performer

# Imports dos pacotes externos
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
from pathlib import Path
import sqlite3


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

        #Criar a fila para o performer - Lançamento no site SoftCase
        Log.write_log('Criando a fila do Performer')
        caminho_banco = Path(InitAllSettings.config['CaminhoBancoSqlite'])
        criar_fila_performer(caminho_banco)

        Log.write_log('Process Finished')

