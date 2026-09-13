"""
VERSÃO FRAMEWORK: 1.0.0

AVISO:

Certifique-se de instalar o bot com `pip install -e .` para obter todas as dependências
em seu ambiente Python.

Além disso, se você estiver usando PyCharm ou outro IDE, certifique-se de usar o MESMO interpretador Python
como seu IDE.
"""
# Imports dos módulos internos do projeto
# Carrega o InitAllSettings - Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log
from ID01_2_ConcEst_TratarDados.classes.utils.ExecutionControl import ExecutionControl
from ID01_2_ConcEst_TratarDados.classes.utils.Exceptions import *
from ID01_2_ConcEst_TratarDados.classes.framework.Initialization import Initialization
from ID01_2_ConcEst_TratarDados.classes.framework.LoopStation import LoopStation
from ID01_2_ConcEst_TratarDados.classes.framework.EndProcess import EndProcess
from ID01_2_ConcEst_TratarDados.classes.dados_execucao.DadosExecucao import DadosExecucao

# Imports dos pacotes externos
import traceback
import sys


class Bot:
    """
    Classe principal responsável por orquestrar a execução do robô (inicialização, loop de
    processamento dos itens da fila e finalização).

    Parâmetros:

    Retorna:
    """

    @classmethod
    def action(cls, execution=None):
        """
        Método principal para execução do bot.

        Parâmetros:
        - execution (objeto): objeto de execução (opcional, default=None). Reservado para integração
          futura com um orquestrador próprio.

        Retorna:
        """
        try:
            ExecutionControl.execution = execution
            Log.write_log("Iniciando execução do processo: " + InitAllSettings.config["NomeProcesso"])

            Initialization.execute()

            LoopStation.execute()

        except TerminateException as err:
            traceback_erro = traceback.format_exc()
            print(traceback_erro)
        except Exception as err:
            traceback_erro = traceback.format_exc()
            if InitAllSettings.exception_initialization is None:
                InitAllSettings.exception_process = err
            print(traceback_erro)

        try:
            EndProcess.execute()

        except Exception as err:
            # Fim do Processamento com Falha
            traceback_erro = traceback.format_exc()
            print(traceback_erro)
            ExecutionControl.send_error(err)
            DadosExecucao.refresh_counting_items()
            ExecutionControl.finish_task(sucesso=False, mensagem=f"Task finalizada com erros. Motivo: {traceback_erro}")

    @classmethod
    def main(cls):
        """
        Ponto de entrada padrão para execução local/manual do bot.

        Parâmetros:

        Retorna:
        """
        cls.action(None)


if __name__ == '__main__':
    if len(sys.argv) >= 5 and str(sys.argv[1]).lower() == "--execution".lower():
        Bot.action(None)
    else:
        Bot.main()
