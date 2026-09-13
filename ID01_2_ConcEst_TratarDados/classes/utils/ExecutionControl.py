# Imports dos módulos internos do projeto
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType

# Imports dos pacotes externos
import socket
import random
import datetime


class ExecutionControl:
    """
    Classe responsável por controlar dados e o ciclo de vida da execução local (runner, job, interrupção
    e finalização da task).

    Como o projeto não depende mais de um orquestrador específico (ex.: BotCity Maestro), esta classe mantém,
    localmente, as informações equivalentes que antes vinham do orquestrador. Caso o seu orquestrador próprio
    exponha uma API (ex.: para sinalizar interrupção ou reportar erros), integre essas chamadas aqui.

    Parâmetros:

    Retorna:
    """

    _config = InitAllSettings.config

    runner_id: str = socket.gethostname()
    versao_runner: str = "X.X.X"
    job_id = _config.get("job_id")
    # Quando não há job id informado (execução manual/local), considera-se execução de teste
    is_test_task: bool = _config.get("job_id") is None
    task = None
    execution = None

    @classmethod
    def is_interrupted(cls) -> bool:
        """
        Indica se foi solicitada a interrupção da execução.

        Como não há mais conexão com um orquestrador remoto, este método retorna sempre False.
        Integre aqui a verificação com o seu próprio orquestrador, caso deseje suportar interrupção remota
        (ex.: consultando uma tabela/endpoint de controle).

        Parâmetros:

        Retorna:
            bool: True caso a execução tenha sido interrompida, False caso contrário.
        """
        return False

    @classmethod
    def finish_task(cls, sucesso: bool, mensagem: str):
        """
        Registra a finalização da task/execução local.

        Integre aqui a chamada para o seu orquestrador, caso deseje reportar o status final da execução.

        Parâmetros:
            - sucesso (bool): Indica se a task deve ser finalizada como sucesso ou erro.
            - mensagem (str): Mensagem a ser adicionada no final da execução.

        Retorna:
        """
        level = LogLevel.INFO if sucesso else LogLevel.ERROR
        Log.write_log(mensagem_log=f"Execução finalizada. Sucesso: {sucesso}. {mensagem}", log_level=level)

    @classmethod
    def send_error(cls, exception: Exception, screenshot: str = None):
        """
        Registra um erro ocorrido durante a execução.

        Integre aqui a chamada para o seu orquestrador, caso deseje reportar erros remotamente.

        Parâmetros:
        - exception (Exception): Exceção que ocorreu na execução do projeto
        - screenshot (str): Caminho da Screenshot do Erro

        Retorna:
        """
        mensagem = f"Erro na execução: {str(exception)}"
        if screenshot:
            mensagem += f" (screenshot: {screenshot})"

        Log.write_log(mensagem_log=mensagem, log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)

    @classmethod
    def send_alert(cls, title_alert: str, message_alert: str, log_level: LogLevel = LogLevel.WARN):
        """
        Registra um alerta ocorrido durante a execução.

        Integre aqui a chamada para o seu orquestrador, caso deseje reportar alertas remotamente.

        Parâmetros:
        - title_alert (str): Título do alerta que será criado
        - message_alert (str): Mensagem do alerta que será criado
        - log_level (LogLevel): Nível do alerta (default = LogLevel.WARN)

        Retorna:
        """
        Log.write_log(mensagem_log=f"[ALERTA] {title_alert}: {message_alert}", log_level=log_level)
