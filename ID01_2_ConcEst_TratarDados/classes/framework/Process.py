# Imports dos módulos internos do projeto
# Carrega o InitAllSettingssSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException
from ID01_2_ConcEst_TratarDados.classes.framework.GetTransaction import GetTransaction
from ID01_2_ConcEst_TratarDados.classes.excel.leitura_relat_adquirentes import (LeitorRelatoriosAdquirentes,)
from ID01_2_ConcEst_TratarDados.classes.sqlite.gravador_tbl_aux_dados import (GravadorTblAuxDados,)

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
        Log.write_log('Process Started')
   
        #Ler os arquivos dos Adquirentes
        leitor = LeitorRelatoriosAdquirentes()
        dados_consolidados = leitor.processar_todos()

        #Salvar na tabela tbl_aux_dados
        gravador = GravadorTblAuxDados()
        gravador.gravar_tbl_aux_dados(dados_consolidados)

        Log.write_log('Process Finished')
