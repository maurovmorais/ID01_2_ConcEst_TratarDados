# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import BusinessRuleException
#FIXME Código Exemplo REMOVER
from ID01_2_ConcEst_TratarDados.classes.chrome.google.Homepage import GoogleHomepage

# Imports dos pacotes externos

class CloseAllApplications:
    """
    Classe para fechar todos os aplicativos no final da automação.

    Parâmetros:
    
    Retorna:
    """
    _config = InitAllSettings.config
    
    @classmethod
    def execute(cls):
        """
        Executa o fechamento de todos os aplicativos necessários, apenas com a estrutura em código.

        Observação:
        - Edite o valor da variável `max_tentativas` no arquivo Config.xlsx.

        Parâmetros:
        
        Retorna:

        Raises:
        - BusinessRuleException: em caso de erro de regra de negócio.
        - Exception: em caso de erro geral.
        """
        #Edite o valor dessa variável a no arquivo Config.xlsx
        max_tentativas = cls._config["MaxRetryNumber"]

        for tentativa in range(max_tentativas):
            try:
                Log.write_log("Finalizando todos os processos, tentativa " + (tentativa+1).__str__())
                #Insira aqui seu código para fechar os aplicativos
                
                #FIXME Código Exemplo REMOVER

                GoogleHomepage.close_google_website()


            except BusinessRuleException as err:
                Log.write_log(mensagem_log="Erro de negócio: " + str(err), 
                                  log_level=LogLevel.ERROR, error_type=ErrorType.BUSINESS_ERROR)
                raise err
            except Exception as err:
                Log.write_log(mensagem_log="Erro, tentativa " + (tentativa+1).__str__() + ": " + str(err),
                                  log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
                
                if(tentativa+1 == max_tentativas): 
                    raise err
                else: 
                    # Inclua aqui o código responsável para reiniciar ao estado indicado para iniciar as aplicações novamente
                    
                    continue
            else:
                Log.write_log("Aplicativos finalizados, continuando processamento...")
                break
            