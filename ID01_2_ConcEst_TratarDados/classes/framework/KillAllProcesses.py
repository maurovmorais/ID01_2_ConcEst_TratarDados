# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException


# Imports dos pacotes externos
import os, signal


class KillAllProcesses:
    """
    Classe para finalizar todos os processos necessários.
    Feita para ser invocada em casos de system exceptions no processamento, para resetar o processamento.
    Pode ser usado em outras partes do processo também, dependendo de como a automação for programada.

    Parâmetros:

    Retorna:
    """
    _config = InitAllSettings.config
    _web_driver = InitAllSettings.web_driver

    @classmethod
    def execute(cls,nomes_processo:list=[]):
        """
        Executa o método para finalizar os processos necessários, apenas com a estrutura em código.

        Parâmetros:
    
        Retorna:
        
        Raises:
            - BusinessRuleException: Se ocorrer um erro de regra de negócio durante o processamento.
            - Exception: Se ocorrer um erro não tratado durante o processamento.
        """
        
        #Edite o valor dessa variável a no arquivo Config.xlsx
        max_tentativas = cls._config["MaxRetryNumber"]

        for tentativa in range(max_tentativas):
            try:
                Log.write_log("Finalizando processos, tentativa " + (tentativa+1).__str__())
                
                for nome_processo in nomes_processo:
                    Log.write_log("Finalizando processo: " + nome_processo)

                    os.system(f"taskkill /f /im  {nome_processo}")

                Log.write_log("Aplicativos finalizados, continuando processamento...")
                break
            except BusinessRuleException as err:
                raise err
            except Exception as err:
                Log.write_log(mensagem_log="Erro, tentativa " + (tentativa+1).__str__() + ": " + str(err), 
                                  log_level=LogLevel.ERROR, 
                                  error_type=ErrorType.APP_ERROR)
                
                if(tentativa+1 == max_tentativas):
                    raise err



