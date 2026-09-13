# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
# Imports dos pacotes externos
import os
import psutil
import traceback
from time import sleep


class RobotStream:
    """
    Classe responsável pela manipulação dos metódos do robot stream
    Robot stream realiza a possibilidade de visualização da tela que está executando o robo, via outro computador na mesma rede
    
    Parâmetros:

    Retorna:

    """
    nome_executavel = 'StreamRobotScreen.exe'

    @classmethod
    def start(cls):
        """
        Inicia o robot stream
        
        Parâmetros:

        Retorna:
        """
        try:
            Log.write_log("Iniciando o Robot Stream")
            sleep(2)
            #Fecha o RobotStream
            for proc in psutil.process_iter():
                if proc.name() == cls.nome_executavel:
                    proc.kill()

            caminho_executavel = InitAllSettings.caminho_executavel_robot_stream
                        
            # Abrindo o executável
            os.startfile(caminho_executavel)
            sleep(5)
            
            # window_id = win32gui.FindWindow(None, "Robot Stream")  
            # text_window = win32gui.GetWindowText(window_id)
            # print(text_window)

        except Exception as exception:
            traceback = traceback.format_exc()
            print(traceback)

    @classmethod
    def stop(cls):
        """
        Finaliza o processo do robot stream
        
        Parâmetros:

        Retorna:

        """
        try:
            Log.write_log("Fechando o Robot Stream")        
            sleep(2)
            #Fecha o RobotStream
            for proc in psutil.process_iter():
                if proc.name() == cls.nome_executavel:
                    proc.kill()
        except Exception as exception:
            traceback = traceback.format_exc()
            print(traceback)           
        
