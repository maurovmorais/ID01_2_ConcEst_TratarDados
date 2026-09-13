"""Leitura das planilhas do arquivo DEPARA_Estacionamento.xlsx.

Cada planilha do DEPARA possui sua própria função de leitura, para
manter responsabilidade única por aba: uma mudança de layout em uma
planilha não afeta a leitura das demais, e cada função pode ser
testada isoladamente.
"""

from __future__ import annotations

import logging
from pathlib import Path

from ID01_2_ConcEst_TratarDados.classes.excel.excel import LeitorExcel
from ID01_2_ConcEst_TratarDados.classes.sqlite.config import CAMINHO_ARQUIVO_DEPARA

logger = logging.getLogger(__name__)

LinhasPlanilha = list[dict[str, object]]


def ler_empresas(leitor: LeitorExcel) -> LinhasPlanilha:
    """Lê a aba 'Empresas' do DEPARA."""
    return leitor.ler_planilha("Empresas")


def ler_forma_de_pagamento(leitor: LeitorExcel) -> LinhasPlanilha:
    """Lê a aba 'FormaDePagamento' do DEPARA."""
    return leitor.ler_planilha("FormaDePagamento")


def ler_cond_pagamt(leitor: LeitorExcel) -> LinhasPlanilha:
    """Lê a aba 'CondPagamt' do DEPARA."""
    return leitor.ler_planilha("CondPagamt")


def ler_adquirentes(leitor: LeitorExcel) -> LinhasPlanilha:
    """Lê a aba 'Adquirentes' do DEPARA."""
    return leitor.ler_planilha("Adquirentes")


def ler_taxa_adquirentes(leitor: LeitorExcel) -> LinhasPlanilha:
    """Lê a aba 'TaxaAdquirentes' do DEPARA."""
    return leitor.ler_planilha("TaxaAdquirentes")


def ler_taxa_cielo(leitor: LeitorExcel) -> LinhasPlanilha:
    """Lê a aba 'TaxaCielo' do DEPARA.

    Observação: a aba possui uma linha totalmente vazia ao final do
    arquivo original; 'LeitorExcel.ler_planilha' já descarta linhas
    totalmente vazias automaticamente.
    """
    return leitor.ler_planilha("TaxaCielo")


def ler_dinheiro(leitor: LeitorExcel) -> LinhasPlanilha:
    """Lê a aba 'Dinheiro' do DEPARA."""
    return leitor.ler_planilha("Dinheiro")


# Funções de leitura por planilha, na ordem em que aparecem no DEPARA.
FUNCOES_LEITURA = {
    "Empresas": ler_empresas,
    "FormaDePagamento": ler_forma_de_pagamento,
    "CondPagamt": ler_cond_pagamt,
    "Adquirentes": ler_adquirentes,
    "TaxaAdquirentes": ler_taxa_adquirentes,
    "TaxaCielo": ler_taxa_cielo,
    "Dinheiro": ler_dinheiro,
}


def ler_todas_planilhas(
    caminho_arquivo: str | Path = CAMINHO_ARQUIVO_DEPARA,
) -> dict[str, LinhasPlanilha]:
    """Lê todas as planilhas do DEPARA.

    Args:
        caminho_arquivo: caminho do arquivo DEPARA_Estacionamento.xlsx.
            Por padrão usa 'config.CAMINHO_ARQUIVO_DEPARA'.

    Returns:
        Dicionário {nome_da_planilha: lista_de_linhas}.
    """
    dados: dict[str, LinhasPlanilha] = {}

    with LeitorExcel(caminho_arquivo) as leitor:
        planilhas_disponiveis = set(leitor.listar_planilhas())
        for nome_planilha, funcao_leitura in FUNCOES_LEITURA.items():
            if nome_planilha not in planilhas_disponiveis:
                logger.warning(
                    "Planilha '%s' não encontrada no arquivo; pulando.",
                    nome_planilha,
                )
                continue
            dados[nome_planilha] = funcao_leitura(leitor)

    return dados


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    resultado = ler_todas_planilhas()
    for planilha, linhas in resultado.items():
        print(f"{planilha}: {len(linhas)} linha(s)")
