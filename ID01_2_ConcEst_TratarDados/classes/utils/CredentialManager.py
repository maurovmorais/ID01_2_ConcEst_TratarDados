# Imports dos módulos internos do projeto
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel

# Imports dos pacotes externos
import os


class CredentialManager:
    """
    Classe responsável por centralizar a obtenção de credenciais utilizadas pela automação.

    Ordem de busca:
        1. Variáveis de ambiente (recomendado para produção - boa prática de segurança).
        2. Aba "Credentials" do arquivo Config.xlsx (já carregada em InitAllSettings.config).

    Caso o seu orquestrador próprio possua um cofre de credenciais (ex.: Azure Key Vault, AWS Secrets Manager,
    ou uma API própria), integre a chamada aqui, mantendo a mesma assinatura de método.

    Parâmetros:

    Retorna:
    """

    @classmethod
    def get_credential(cls, key: str, label: str = None) -> str:
        """
        Retorna uma credencial para uso durante a execução.

        Parâmetros:
            - key (str): Chave da credencial (nome da variável de ambiente ou chave no Config.xlsx).
            - label (str): Mantido por compatibilidade com versões anteriores; não é utilizado na
              busca local, pois as credenciais do Config.xlsx já são carregadas por chave.

        Retorna:
            - str: valor da credencial encontrada, ou None caso não seja encontrada.
        """
        try:
            valor_env = os.environ.get(key)
            if valor_env is not None:
                return valor_env

            valor_config = InitAllSettings.config.get(key)
            if valor_config is not None:
                return valor_config

            Log.write_log(mensagem_log=f"Credencial '{key}' não encontrada em variáveis de ambiente nem no Config.xlsx.", log_level=LogLevel.WARN)
            return None
        except Exception as err:
            Log.write_log(mensagem_log=f"Não foi possível obter a credencial '{key}': {str(err)}", log_level=LogLevel.ERROR)
            return None
