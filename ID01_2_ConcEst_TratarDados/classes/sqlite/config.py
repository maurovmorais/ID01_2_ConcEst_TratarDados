"""Configurações e caminhos utilizados pela automação ID01_2_ConcEst_TratarDados."""

from pathlib import Path
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings

# TODO: ajustar para o caminho real do arquivo DEPARA na máquina de produção.
CAMINHO_ARQUIVO_DEPARA = Path(
    #r"C:\Armazenamento\ID01_2_ConcEst_TratarDados\DePara\DEPARA_Estacionamento.xlsx"
    InitAllSettings.config['depara']
)

CAMINHO_BANCO_DADOS = Path(
    #r"C:\Armazenamento\ID01_2_ConcEst_TratarDados\Banco Dados\banco_dados.db"
    InitAllSettings.config['CaminhoBancoSqlite']
)

# Mapeamento: nome da planilha no DEPARA -> nome da tabela no SQLite.
NOME_TABELAS: dict[str, str] = {
    "Empresas": "tbl_empresas",
    "FormaDePagamento": "tbl_FormaDePagamento",
    "CondPagamt": "tbl_CondPagamt",
    "Adquirentes": "tbl_Adquirentes",
    "TaxaAdquirentes": "tbl_TaxaAdquirentes",
    "TaxaCielo": "tbl_TaxaCielo",
    "Dinheiro": "tbl_Dinheiro",
}
