# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
# Imports dos pacotes externos
import base64
import psutil
import ctypes
import subprocess
from screeninfo import get_monitors
from os import path, remove, listdir
from os.path import getmtime, join, getsize
from datetime import datetime
from operator import itemgetter
from typing import Union, List
from win32 import win32cred


def image_to_base64(caminho_arquivo:str) -> str:
    """
    Converte uma imagem em base64.

    Parâmetros:
        - caminho_arquivo(str): caminho completo do arquivo que será convertido para base64.
   
    Retorna:
        - str: Retorna uma string contendo a string do arquivo em base64
   """
   
    with open(caminho_arquivo, "rb") as file:
        encoded_file = base64.b64encode(file.read()).__str__().lstrip('b').strip("'")
    

    return encoded_file

def get_computer_usage() -> str:
    """
    Retorna a quantidade de Memoria Ram, CPU e Disco Utilizada no momento do processamento
    
    Parâmetros:

    Retorna:
        - computer_usage (str): texto contendo a quantidade de Memoria Ram, CPU e Disco Utilizada no momento do processamento
    """   
    porcentagem_mem_utilizada = f"A porcentagem de Memória RAM utilizada é de {psutil.virtual_memory().percent}"
    porcentagem_cpu_utilizada = f"A porcentagem de CPU utilizada é de {psutil.cpu_percent()}"
    porcentagem_disco_utilizada = f"A porcentagem de Disco utilizada é de {psutil.disk_usage('/').percent}"

    computer_usage = '\n'.join(["Informações do Uso da Máquina", porcentagem_mem_utilizada, porcentagem_cpu_utilizada, porcentagem_disco_utilizada,])

    return computer_usage

def _get_dpi_scale(scr_monitor) -> float:
    """
    Obtém a escala de um monitor específico.

    Parâmetros:
        - scr_monitor: Objeto monitor do screeninfo que contém as coordenadas (x, y).

    Retorna:
        - escala_dpi (float): Retorna a escala de DPI do monitor.
    """

    # Define uma classe POINT que representa uma estrutura usada para coordenadas (x, y).
    class POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    
    # Cria uma instância da estrutura POINT com as coordenadas do monitor fornecido.
    point_monitor = POINT(scr_monitor.x, scr_monitor.y)
    
    # Obtém o handle do monitor que contém o ponto especificado (point_monitor).
    # O segundo argumento (2) indica que o monitor primário será retornado se o ponto estiver fora de qualquer monitor.
    handle_monitor = ctypes.windll.user32.MonitorFromPoint(point_monitor, 2)
    
    # Cria variáveis para armazenar os valores de DPI horizontal e vertical.
    uint_dpi_x = ctypes.c_uint()
    uint_dpi_y = ctypes.c_uint()
    
    # Obtém o DPI do monitor especificado.
    ctypes.windll.shcore.GetDpiForMonitor(
        handle_monitor, 0, ctypes.byref(uint_dpi_x), ctypes.byref(uint_dpi_y)
    )
    
    # Calcula a escala de DPI com base no DPI horizontal obtido. O valor 96.0 é considerado o DPI padrão
    escala_dpi = uint_dpi_x.value / 96.0
    
    return escala_dpi

def get_monitor_info() -> str:
    """
    Coleta informações de resolução e escala de cada monitor.

    Parâmetros:
    
    Retorna:
    - info_monitores_texto (str): informações de resolução e escala de cada monitor.
    """

    info_monitores = []
    
    for i, monitor in enumerate(get_monitors(), start=1):
        resolucao = f"{monitor.width}x{monitor.height}"
        escala = _get_dpi_scale(monitor) * 100
        info_monitores.append(f"A resolução do Monitor {i} é {resolucao}, escala de {escala:.0f}%")
    
    info_monitores_texto = '\n'.join(['Informações do Monitor', *info_monitores])
    
    return info_monitores_texto

def get_system_info() -> str:
    """
    Retorna informações detalhadas do sistema, incluindo RAM, disco e CPU.
    
    Parâmetros:

    Retorna:
        - info_sistema (str): Informações sobre RAM, disco e CPU.
    """

    # Informações de RAM
    total_ram_gb = psutil.virtual_memory().total / (1024 ** 3)

    # Informações de Disco
    total_disco_gb = psutil.disk_usage('/').total / (1024 ** 3)
    disco_usado_gb = psutil.disk_usage('/').used / (1024 ** 3)

    # Informações de CPU
    numero_cpus = psutil.cpu_count(logical=True)
    freq_max_cpu = psutil.cpu_freq().max / 1000 if psutil.cpu_freq() else None  # Convertendo MHz para GHz

    try:
        # Modelo CPU
        comando_power_shell = "Get-WmiObject Win32_Processor | Select-Object -ExpandProperty Name"
        sp_resultado = subprocess.run(['powershell', '-Command', comando_power_shell], capture_output=True, text=True, shell=True, timeout=3)
        descricao_cpu = sp_resultado.stdout.strip()
    except Exception as err:
        descricao_cpu = "CPU não reconhecida"

    # Informações do Sistema
    sistema = "Informações do Sistema"
    memoria_ram = f'Total de Memória RAM: {total_ram_gb:.2f} GB'
    disco = f'Disco usado: {disco_usado_gb:.2f} GB, total de: {total_disco_gb:.2f} GB'
    clocks_cpu = f'Número de Clocks CPU: {numero_cpus}'
    freq_cpu = f'Frequência Máxima da CPU: {freq_max_cpu:.2f} GHz' if freq_max_cpu else "Frequência da CPU não disponível"
    modelo_cpu = f'Modelo da CPU: {descricao_cpu}'

    info_sistema = '\n'.join([sistema, modelo_cpu, memoria_ram, 
                                    disco, clocks_cpu, freq_cpu])

    return info_sistema

def remover_arquivos_antigos(diretorio_arquivo:str, dias_excluir_arquivo:int, extensoes_arquivo: Union[str, List[str], tuple[str, ...]]):
    """
    Remove arquivos em um diretório que são mais antigos que um número especificado de dias.

    Parâmetros:
    - diretorio_arquivo (str): O caminho do diretório contendo os arquivos.
    - dias_excluir_arquivo (int): Número de dias de idade para que um arquivo seja removido.
    - extensoes_arquivo (Union[str, List[str], tuple[str, ...]]): Extensão, lista ou tupla de extensões a serem consideradas (ex: '.xlsx', ['.xlsx', '.csv']).

    Retorna:
    """
    try:
        # Data atual da execução
        data_atual = datetime.now()

        # Normaliza a(s) extensão(ões) para uma tupla.
        tuple_extensoes = extensoes_arquivo
        if isinstance(tuple_extensoes, str):
            tuple_extensoes = (tuple_extensoes,)

        # Garante que cada extensão na tupla comece com um ponto.
        tuple_extensoes_normalizadas = tuple(
            ext if ext.startswith('.') else f".{ext}" for ext in tuple_extensoes
        )

        # Percorre a pasta onde estão os arquivos.
        for arquivo in listdir(diretorio_arquivo):
            caminho_arquivo = join(diretorio_arquivo, arquivo)
            if path.isfile(caminho_arquivo) and caminho_arquivo.endswith(tuple_extensoes_normalizadas):
                modificacao_arq = datetime.fromtimestamp(getmtime(caminho_arquivo))
                diferenca_dias = (data_atual - modificacao_arq).days
                
                if diferenca_dias > dias_excluir_arquivo:
                    Log.write_log(f"Excluindo arquivo antigo: {caminho_arquivo} ({diferenca_dias} dias de idade, limite: {dias_excluir_arquivo}).")
                    remove(caminho_arquivo)
    except FileNotFoundError:
        Log.write_log(f"Diretório não encontrado: '{diretorio_arquivo}'.", log_level=LogLevel.ERROR)
    except OSError as err:
        Log.write_log(f"Erro ao acessar o diretório '{diretorio_arquivo}' ou seus arquivos: {err}", log_level=LogLevel.ERROR)
    
def controlar_limite_espaco(diretorio_arquivos:str, limite_gb_pasta:float, extensoes_arquivo: List[str]):
    """
    Controla o tamanho total dos arquivos em um diretório e exclui os mais antigos
    até que o tamanho total esteja abaixo do limite especificado.

    Parâmetros:
    - diretorio_arquivos (str): O caminho do diretório contendo os arquivos.
    - limite_gb_pasta (float): O limite de espaço da pasta em GB.
    - extensoes_arquivo List[str]: lista de extensões a serem consideradas (ex: '.xlsx', ['.xlsx', '.csv']).

    Retorna:
    """
    try:
        # Normaliza a(s) extensão(ões) para uma tupla.
        tuple_extensoes = extensoes_arquivo
        if isinstance(tuple_extensoes, str):
            tuple_extensoes = (tuple_extensoes,)

        # Garante que cada extensão na tupla comece com um ponto.
        tuple_extensoes_normalizadas = tuple(
            ext if ext.startswith('.') else f".{ext}" for ext in tuple_extensoes
        )

        # Lista todos os arquivos relevantes com seus tamanhos e datas de modificação
        arquivos = []
        for arquivo in listdir(diretorio_arquivos):
            caminho_arquivo = join(diretorio_arquivos, arquivo)
            if path.isfile(caminho_arquivo) and caminho_arquivo.endswith(tuple_extensoes_normalizadas):
                arquivos.append(
                    (caminho_arquivo, getmtime(caminho_arquivo), getsize(caminho_arquivo))
                )

        # Calcula o tamanho total em GB
        tamanho_total_pasta = sum(tamanho for _, _, tamanho in arquivos)
        tamanho_total_pasta_gb = tamanho_total_pasta / (1024 ** 3)

        # Ordena os arquivos pela data de modificação (mais antigo primeiro)
        arquivos.sort(key=itemgetter(1))

        # Se o tamanho exceder o limite, remove os arquivos mais antigos
        while tamanho_total_pasta_gb > limite_gb_pasta and len(arquivos) > 1:
            Log.write_log(f"Tamanho da pasta ({tamanho_total_pasta_gb:.4f} GB) excede o limite de {limite_gb_pasta} GB.")
            
            # Pega o arquivo mais antigo para remover
            arquivo_para_remover, _, tamanho_arquivo_removido = arquivos.pop(0)
            
            Log.write_log(f"Excluindo arquivo mais antigo para liberar espaço: {arquivo_para_remover}")
            remove(arquivo_para_remover)
            
            # Recalcula o tamanho total da pasta
            tamanho_total_pasta -= tamanho_arquivo_removido
            tamanho_total_pasta_gb = tamanho_total_pasta / (1024 ** 3)

        if tamanho_total_pasta_gb <= limite_gb_pasta:
            espaco_restante = limite_gb_pasta - tamanho_total_pasta_gb
            Log.write_log(f"Tamanho da pasta está dentro do limite. Espaço restante: {espaco_restante:.4f} GB.")

    except FileNotFoundError:
        Log.write_log(f"Diretório não encontrado: '{diretorio_arquivos}'.", log_level=LogLevel.ERROR)
    except OSError as e:
        Log.write_log(f"Erro ao acessar o diretório '{diretorio_arquivos}' ou seus arquivos: {e}", log_level=LogLevel.ERROR)

def get_windows_credential(target_name: str) -> dict | None:
    """
    Recupera uma credencial salva no Windows Credential Manager.

    Args:
        target_name (str): Nome do alvo da credencial (TargetName).

    Returns:
        dict | None: Dicionário com as informações da credencial ou None se não encontrada.
    """
    try:
        #Enumerando todas as credenciais na máquina
        creds = win32cred.CredEnumerate()

        #Iterando todos os resultados
        for cred in creds:
            #Se encontrar uma credencial com o nome certo, continua com decode e retorna o resultado
            if(cred["TargetName"] == target_name):
                usuario: str = cred["UserName"]
                senha: str = cred["CredentialBlob"].decode("utf-16")
                target: str = cred["TargetName"]

                dic_resultado: dict = {
                    "UserName": usuario,
                    "Password": senha,
                    "TargetName": target
                }

                return dic_resultado
        
        return None

    except Exception as e:
        print(f"Erro ao recuperar credencial: {e}")
        return None