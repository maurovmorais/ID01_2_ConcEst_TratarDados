# Imports dos módulos internos do projeto
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log

# Imports dos pacotes externos
import time
import threading
from os import path, remove, listdir
from os.path import getmtime, join, getsize
from datetime import datetime
from operator import itemgetter

import numpy as np
import cv2
import pyautogui


class ScreenRecorder:
    """
    Classe responsável por gravar a tela do processo.

    Utiliza `pyautogui` para capturar frames da tela e `opencv-python` (cv2) para gravar o vídeo,
    substituindo a dependência do plugin de gravação do Botcity.

    Parâmetros:

    Retorna:
    """

    _config = InitAllSettings.config

    if(_config["GravarTela"].upper() == "SIM"):
        # Alterar no caminho o nome do vídeo caso queira que sobrescreva para reduzir o espaço ocupado
        caminho_completo = path.join(_config["CaminhoSalvarVideo"], _config["NomeProcesso"] + "_" + datetime.now().strftime("%d%m%Y%H%M%S") + ".avi")
        started:bool = False
        _fps:int = 10
        _thread_gravacao = None

        # Caminho onde serão salvos os videos
        caminho_completo = _config["CaminhoSalvarVideo"]

    @classmethod
    def _loop_gravacao(cls, caminho_video:str):
        """
        Executado em uma thread separada: captura frames da tela periodicamente e grava no arquivo de vídeo.

        Parâmetros:
        - caminho_video (str): caminho completo do arquivo de vídeo a ser gravado.

        Retorna:
        """
        resolucao = pyautogui.size()
        fourcc = cv2.VideoWriter_fourcc(*"XVID")
        writer = cv2.VideoWriter(caminho_video, fourcc, cls._fps, resolucao)

        try:
            while cls.started:
                img_screenshot = pyautogui.screenshot()
                arr_frame = np.array(img_screenshot)
                arr_frame = cv2.cvtColor(arr_frame, cv2.COLOR_RGB2BGR)
                writer.write(arr_frame)
                time.sleep(1 / cls._fps)
        finally:
            writer.release()

    @classmethod
    def iniciar_gravacao(cls):
        """
        Inicia a gravação de tela do processo.

        Parâmetros:
    
        Retorna:
        """

        if cls._config["RemoverVideosAntigos"] == "SIM":    
            # Remover videos antigos com base em quantos dias deve ficar na pasta.
            dias_excluir_video = int(cls._config["DiasLimiteVideo"])
            cls._remover_videos_antigos(diretorio_videos=cls.caminho_completo, dias_excluir_video=dias_excluir_video)
        
        if cls._config["RemoverVideosLimiteGb"] == "SIM":
            # Remover videos com base no limite de GBs estabelecido.
            limite_gb_video = float(cls._config["TamanhoLimitePasta"])
            cls._verificar_limite_espaco(diretorio_videos=cls.caminho_completo, limite_gb_pasta=limite_gb_video)
            
        if not cls.started:
            Log.write_log("Iniciando Gravador de Tela...")
            nome_arquivo = path.join(cls.caminho_completo, cls._config["NomeProcesso"] + "_" + datetime.now().strftime("%d%m%Y%H%M%S") + ".avi")

            cls.started = True
            cls._thread_gravacao = threading.Thread(target=cls._loop_gravacao, args=(nome_arquivo,), daemon=True)
            cls._thread_gravacao.start()
            time.sleep(5)

    @classmethod
    def finalizar_gravacao(cls):
        """
        Finaliza a gravação de tela do processo.
        
        Parâmetros:
    
        Retorna:        
        """
        Log.write_log("Finalizando Gravador de Tela...")
        cls.started = False

        if cls._thread_gravacao is not None:
            cls._thread_gravacao.join(timeout=10)

        time.sleep(5)

    @classmethod
    def _remover_videos_antigos(cls, diretorio_videos:str, dias_excluir_video:int):
        """
        Remove videos no diretório que são mais antigos do que um número especificado de dias.

        Parâmetros:
        diretorio_videos (str): O caminho do diretório contendo os videos.
        dias_excluir_video (int): Número de dias limite para a existencia dos videos.

        Retorna:
        """

        # Data atual da execução
        data_atual = datetime.now()

        # Percorre a pasta onde estão os videos
        for arquivo in listdir(diretorio_videos):
            caminho_video = join(diretorio_videos, arquivo)
            if path.isfile(caminho_video) and (caminho_video.endswith('.avi') or caminho_video.endswith('.mp4')):
                modificacao_video = datetime.fromtimestamp(getmtime(caminho_video))
                
                # Calcula a diferença em dias
                diferenca_dias = (data_atual - modificacao_video).days
                
                # Compara a diferença de dias com o limite
                if diferenca_dias > dias_excluir_video:
                    Log.write_log(f"Excluindo vídeo: {caminho_video}, Data de Modificação: {modificacao_video}, Dias de diferença: {diferenca_dias}, Limite de dias: {dias_excluir_video}")
                    remove(caminho_video)
                else:
                    dias_restantes = dias_excluir_video - diferenca_dias
                    Log.write_log(f"Faltam {dias_restantes} dias para o vídeo {caminho_video} ser deletado devido ao limite de dias na pasta.")

    @classmethod
    def _verificar_limite_espaco(cls, diretorio_videos:str, limite_gb_pasta:float):
        """
        Verifica o tamanho total dos arquivos no diretório e exclui videos se o limite for excedido,
        mantendo pelo menos o video mais recente.

        Parâmetros:
        diretorio_videos (str): O caminho do diretório contendo os videos.
        limite_gb_pasta (float): O limite de espaço da pasta em GB.

        Retorna:
        """

        # Inicializa variável contadora do tamanho total da pasta
        tamanho_total_pasta = 0
        videos = []

        # Percorre a pasta onde estão os videos, realizando a somatória do tamanho dos arquivos.
        for arquivo in listdir(diretorio_videos):
            diretorio_video = join(diretorio_videos, arquivo)
            if path.isfile(diretorio_video) and (diretorio_video.endswith('.avi') or diretorio_video.endswith('.mp4')):
                tamanho_video = getsize(diretorio_video)
                tamanho_total_pasta += tamanho_video
                videos.append((diretorio_video, getmtime(diretorio_video)))

        # Converte o tamanho total de bytes para gigabytes
        tamanho_total_pasta_gb = tamanho_total_pasta / (1024 ** 3)

        # Se a somatória dos videos for maior que o tamanho limite estabelecido, percorre a pasta e deleta os arquivos.
        if tamanho_total_pasta_gb > limite_gb_pasta:
            Log.write_log(f"Tamanho total da pasta: {tamanho_total_pasta_gb:.2f} GB. Limite estabelecido: {limite_gb_pasta} GB")

            # Ordena os videos pela data de modificação (mais antigo primeiro)
            videos.sort(key=itemgetter(1))

            # Exclui todos os videos, exceto o mais recente
            for diretorio_video, _ in videos[:-1]:
                Log.write_log(f"Excluindo video devido ao limite de espaço: {diretorio_video}")
                remove(diretorio_video)

        else:
            espaco_restante = limite_gb_pasta - tamanho_total_pasta_gb
            Log.write_log(f"Faltam {espaco_restante:.2f} GB para atingir o limite da pasta. Nenhum vídeo foi excluído.")
