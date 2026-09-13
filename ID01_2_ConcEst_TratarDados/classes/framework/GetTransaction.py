# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.ExecutionControl import ExecutionControl
from ID01_2_ConcEst_TratarDados.classes.queue.QueueManager import QueueManager
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException,TerminateException
from ID01_2_ConcEst_TratarDados.classes.dados_execucao.DadosExecucao import DadosExecucao

# Imports dos pacotes externos


class GetTransaction:
    """
    Classe Reponsavel pela organização do código da execução dos passos categorizados como itens de fila
    
    Parâmetros:
  
    Retorna:
    """
    queue_item:dict = None

    @classmethod
    def execute(cls):
        """
        Metodo de execução da captura de itens da fila.
        
        Parâmetros:

        Retorna:

        """
        # 491     Capturar Item Da Fila

        #Verifica se foi solicitado a interrupção da execução
        if(ExecutionControl.is_interrupted()): 
            cls.queue_item = None
            raise TerminateException("Interrupção da execução foi solicitada.")
        else:
            cls.queue_item = QueueManager.get_next_queue_item()
            

        #Processamento continua até a classe informar que não existe itens novos para processar
        if(cls.queue_item is not None): 
            DadosExecucao.refresh_counting_items()
            Log.write_log(f"{str(InitAllSettings.qtde_itens_processados)} itens processados de {InitAllSettings.qtd_itens_a_processar_ini_exec} totais para serem processados.")
            Log.write_log(mensagem_log=f"Executando item {str(InitAllSettings.qtde_itens_processados+1)} de {InitAllSettings.qtd_itens_a_processar_ini_exec}, referência: " + cls.queue_item['referencia'], referencia=cls.queue_item['referencia'] )

        else:
            DadosExecucao.refresh_counting_items()
            Log.write_log(f"{str(InitAllSettings.qtde_itens_processados)} itens processados de {InitAllSettings.qtd_itens_a_processar_ini_exec} totais para serem processados.")
            
            Log.write_log("Não existem itens a serem processados.")
            cls.queue_item = None
        
    
