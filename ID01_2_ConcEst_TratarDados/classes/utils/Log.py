# Imports dos módulos internos do projeto
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings

# Imports dos pacotes externos
import logging
from enum import Enum
from pathlib import Path
from datetime import datetime

ROOT_DIR = Path(__file__).parent.parent.parent


class LogLevel(Enum):
    """
    Classe enum, usada para colocar o nível de um novo log.

    Parâmetros:

    Retorna:
    """

    INFO = "INFO"
    WARN = "WARN"
    ERROR = "ERROR"
    FATAL = "FATAL"


class ErrorType(Enum):
    """
    Classe enum, usada para colocar o tipo do erro.

    Parâmetros:

    Retorna:
    """

    NONE = ""
    APP_ERROR = "APPLICATION"
    BUSINESS_ERROR = "BUSINESS"


_MAP_LOGLEVEL_TO_LOGGING = {
    LogLevel.INFO: logging.INFO,
    LogLevel.WARN: logging.WARNING,
    LogLevel.ERROR: logging.ERROR,
    LogLevel.FATAL: logging.CRITICAL,
}


def _build_logger() -> logging.Logger:
    """
    Configura e retorna o logger estruturado do projeto, gravando tanto no console quanto em arquivo.

    Parâmetros:

    Retorna:
    - logging.Logger: logger configurado.
    """
    logger = logging.getLogger("ID01_2_ConcEst_TratarDados")
    logger.setLevel(logging.DEBUG)

    if not logger.handlers:
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s", datefmt="%d/%m/%Y %H:%M:%S")

        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

        try:
            pasta_logs = ROOT_DIR / "resources" / "logs"
            pasta_logs.mkdir(parents=True, exist_ok=True)
            arquivo_log = pasta_logs / f"execucao_{datetime.now().strftime('%Y%m%d')}.log"

            file_handler = logging.FileHandler(arquivo_log, encoding="utf-8")
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        except Exception:
            # Se não for possível criar o arquivo de log, mantém apenas o log no console
            pass

    return logger


class Log:
    """
    Classe responsável por centralizar a geração de logs estruturados do projeto, utilizando o módulo `logging` do Python.

    Parâmetros:

    Retorna:
    """

    _logger: logging.Logger = _build_logger()

    @classmethod
    def write_log(cls, mensagem_log: str, referencia: str = "-", log_level: LogLevel = LogLevel.INFO, error_type: ErrorType = ErrorType.NONE):
        """
        Grava uma nova entrada de log estruturado (console + arquivo) com os dados fornecidos em argumentos.

        Parâmetros:
            - mensagem_log (str): A mensagem principal a ser gravada no log.
            - referencia (str): Referência do item da fila. `(DEFAULT = "-")`
            - log_level (LogLevel): Enum LogLevel usado para indicar a seriedade do log. `(DEFAULT = LogLevel.INFO)`
            - error_type (ErrorType): Enum ErrorType usado para indicar qual o tipo do erro. (DEFAULT = ErrorType.NONE)

        Retorna:
        """
        if referencia != "-":
            mensagem_formatada = f"{referencia} - {mensagem_log}"
        else:
            mensagem_formatada = mensagem_log

        if error_type != ErrorType.NONE:
            mensagem_formatada += f" - {error_type.value}"

        level = _MAP_LOGLEVEL_TO_LOGGING.get(log_level, logging.INFO)
        cls._logger.log(level, mensagem_formatada)
