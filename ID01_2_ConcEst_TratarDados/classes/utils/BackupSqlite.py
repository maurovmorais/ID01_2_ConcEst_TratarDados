# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
# Imports dos pacotes externos
import os
import shutil
from datetime import datetime, timedelta
from pathlib import Path
import glob

ROOT_DIR = Path(__file__).parent.parent.parent

class BackupSQLite:
    """
    Classe responsável pela criação de backup do banco de dados SQLite.

    Parâmetros:

    Retorna:
    """
    config = InitAllSettings.config
    
    @classmethod
    def criar_backup_sqlite(cls, diretorio_backup: str, dias_para_atualizar: int) -> None:
        """
        Gerencia um único arquivo de backup para um banco de dados SQLite.

        Esta função verifica se um arquivo de backup já existe. Se não existir,
        cria um novo imediatamente. Se existir, verifica a data da última
        modificação. Um novo backup (sobrescrevendo o antigo) só será criado
        se a idade do backup existente for maior ou igual ao número de dias
        especificado em 'dias_para_atualizar'.

        Parâmetros:
        - diretorio_backup(str): O diretório onde o único arquivo de backup será salvo.
        - dias_para_atualizar(int): O número de dias de idade que o backup deve ter para ser atualizado.
            
        Retorna:
        """
        try:
            caminho_banco_sqlite = cls.config["CaminhoBancoSqlite"]

            # --- 1. Validações e Preparação de Variáveis ---
            if not os.path.exists(caminho_banco_sqlite):
                Exception(f"Erro Crítico: O arquivo de banco de dados original em '{str(cls.config["CaminhoBancoSqlite"])}' não foi encontrado.")

            # Garante que o diretório de destino para o backup exista.
            try:
                os.makedirs(diretorio_backup, exist_ok=True)
            except OSError as erro:
                raise Exception(f"Erro Crítico: Não foi possível criar o diretório de backup em '{diretorio_backup}'. Erro: {erro}")

            # Busca arquivos de backup existentes com o padrão banco_backup*.db
            backups = glob.glob(os.path.join(diretorio_backup, "banco_backup*.db"))
            realizar_backup = False
            backup_antigo = None

            # --- 2. Lógica para Decidir se o Backup Deve Ser Atualizado ---
            if not backups:
                Log.write_log("Arquivo de backup não encontrado. Um novo backup será criado.")
                realizar_backup = True
            else:
                # Calcula a idade do arquivo de backup existente.
                # Pega o backup mais recente pelo timestamp de modificação
                backup_antigo = max(backups, key=os.path.getmtime)
                flt_timestamp_modificacao = os.path.getmtime(backup_antigo)
                data_modificacao = datetime.fromtimestamp(flt_timestamp_modificacao).date()
                hoje = datetime.now().date()
                idade_em_dias = (hoje - data_modificacao).days

                # Compara a idade com o parâmetro recebido.
                if idade_em_dias >= dias_para_atualizar:
                    Log.write_log(f"O backup existente tem {idade_em_dias} dia(s) (limite: {dias_para_atualizar}). Atualizando backup.")
                    realizar_backup = True
                else:
                    Log.write_log(f"Backup com {idade_em_dias} dia(s) de idade está dentro do limite de {dias_para_atualizar} dia(s). Nenhuma ação necessária.")

            # --- 3. Execução do Backup ---
            if realizar_backup:
                try:
                    data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
                    nome_arquivo_backup = f"banco_backup_{data_atual}.db"
                    caminho_completo_backup = os.path.join(diretorio_backup, nome_arquivo_backup)

                    # shutil.copy2 copia o arquivo e também seus metadados, como a data de modificação.
                    shutil.copy2(caminho_banco_sqlite, caminho_completo_backup)
                    Log.write_log(f"Backup criado/atualizado com sucesso em: {caminho_completo_backup}")

                    if backup_antigo and os.path.exists(backup_antigo):
                        try:
                            os.remove(backup_antigo)
                            Log.write_log(f"Backup antigo removido: {backup_antigo}")
                        except Exception as e:
                            Log.write_log(f"Erro ao remover backup antigo: {e}")
                            
                except Exception as erro:
                    Log.write_log(f"Erro ao tentar copiar o arquivo de backup: {erro}")
                    return

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao criar backup do arquivo SQLite: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao criar backup do arquivo SQLite: {str(err)}")
        
        