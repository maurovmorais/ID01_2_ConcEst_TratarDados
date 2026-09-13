from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
from datetime import datetime
from types import FunctionType
import functools

def log_execucao_metodo(nome_classe):
    """
    Decorator para logar as classes e métodos.

    Parâmetros:
    nome_classe (str): Nome da classe onde o método está sendo decorado.

    Retorna:
    """

    def decorator(func):
        """
        Envolve a execução de uma função com logs de inicio, fim e tempo gasto.

        Parâmetros:

        Retorna:
        """

        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            """
            Executa a função decorada com logs de início, fim e tempo gasto.

            Parâmetros:

            Retorna:
            """
            # Verifica se os logs estão ativados no arquivo config.xlsx
            if InitAllSettings.config["AtivarLogs"].upper() == "NÃO":
                return func(*args, **kwargs)
            
            # Captura o horário de inicio da execução do método e formata para string
            data_inicio = datetime.now()
            data_inicio_str = data_inicio.strftime("%d/%m/%Y %H:%M:%S")
            
            # Obtém o nome do método que está sendo executado
            nome_metodo = func.__name__
            Log.write_log(f"Inicio - {nome_metodo}() - {nome_classe} - {data_inicio_str}")
            
            # Chama a função original e armazena o resultado
            nt_resultado = func(*args, **kwargs)
            
            # Captura o horário de término da execução do método e formata para string.
            data_fim = datetime.now()
            data_fim_str = data_fim.strftime("%d/%m/%Y %H:%M:%S")
            Log.write_log(f"Fim - {nome_metodo}() - {nome_classe} - {data_fim_str}")
            
            # Calcula o tempo gasto na execução do método
            tempo_gasto = data_fim - data_inicio
            tempo_gasto_str = str(tempo_gasto).split('.')[0]
            Log.write_log(f"Tempo Gasto na execução do Método {nome_metodo}: {tempo_gasto_str}")
            
            # Retorna o resultado da função original
            return nt_resultado
        
        # Retorna a função wrapper que inclui os logs
        return wrapper
    
    # Retorna o decorator que será aplicado na função
    return decorator

def log_metodo(cls):
    """
    Decorador para aplicar logs a todos os métodos de uma classe.

    Parâmetros:

    Retorna:
    """

    # Obtém o nome da classe
    nome_classe = cls.__name__
    
    # Itera sobre os atributos da classe
    for nome_atributo, valor_atributo in cls.__dict__.items():
        # Verifica se o atributo é uma função
        if isinstance(valor_atributo, FunctionType):
            cm_decorador = log_execucao_metodo(nome_classe)(valor_atributo)
            
            # Atualiza o atributo da classe com a função decorada
            setattr(cls, nome_atributo, cm_decorador)
        
        # Verifica se o atributo é um classmethod
        elif isinstance(valor_atributo, classmethod):
            func_funcao_original = valor_atributo.__func__
            cm_decorador = classmethod(log_execucao_metodo(nome_classe)(func_funcao_original))
            
            # Atualiza o atributo da classe com o classmethod decorado
            setattr(cls, nome_atributo, cm_decorador)
        
        # Verifica se o atributo é um staticmethod
        elif isinstance(valor_atributo, staticmethod):
            func_funcao_original = valor_atributo.__func__
            cm_decorador = staticmethod(log_execucao_metodo(nome_classe)(func_funcao_original))
            
            # Atualiza o atributo da classe com o staticmethod decorado
            setattr(cls, nome_atributo, cm_decorador)
    
    # Retorna a classe com os métodos decorados
    return cls