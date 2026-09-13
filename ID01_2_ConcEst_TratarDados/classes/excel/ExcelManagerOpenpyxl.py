# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings as InitAllSetting
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType

# Imports dos pacotes externos
import xlwings
import time
import numpy as np
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, NamedStyle, Border, Side
from xlwings import Sheet, Range,Book,App
from datetime import datetime

class ExcelManagerOpenpyxl:
    """
    Classe para gerenciar operações comuns em planilhas Excel usando openpyxl.

    Parâmetros:

    Retorna:
    """

    @classmethod
    def inserir_cor(cls, caminho_planilha:str, cor_celula:str, range:str, any_aba:any=0) -> None:
        """
        Insere uma cor de preenchimento em um intervalo de células na planilha.

        Parâmetros:
        - caminho_planilha (str): Caminho da planilha Excel onde o Workbook será carregado.
        - cor_celula (str): Código Hexadecimal de cor para o preenchimento das células.
        - range (str): Intervalo de células no formato 'A1:B2' ou no formato 'A1' sendo A1 o inicio da inserção dos dados sem especificar o final de onde sera inserido o dado. 
        - any_aba (any): Recebe o nome ou o indice da aba que será trabalhada

        Retorna:
        """
        try:
            # Carrega o workbook a partir da planilha
            wb_planilha = load_workbook(caminho_planilha)

            if isinstance(any_aba, int):
                ws_planilha = wb_planilha.worksheets[any_aba]
            else:
                ws_planilha = wb_planilha[any_aba]

            if range.__contains__(':'):
                # Intervalo de células em células inicial e final
                celula_inicial, celula_final = range.split(':')
            else:
                # Intervalo de células em células inicial e final
                celula_inicial = range
                celula_final = range

            # Itera sobre cada célula no intervalo
            for row in ws_planilha[celula_inicial:celula_final]:
                for cell in row:
                    # Aplica o preenchimento de cor na célula
                    cell.fill = PatternFill(start_color=cor_celula, end_color=cor_celula, fill_type="solid")

            # Salva as alterações no arquivo
            wb_planilha.save(caminho_planilha)

            Log.write_log(f"Inserindo a cor {cor_celula} no intervalo {range}.")

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao inserir cor: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao inserir cor: {str(err)}")

    @classmethod
    def inserir_formatacao_texto(cls, caminho_planilha:str='', any_aba:any=0, range:str='', 
                                 formato:str = '', style_font:str = None,flt_font_size:float=12.0,
                                 font_color:str='000000',horizontal_align:str='center',quebra_linha=False) -> None:
        """
        Insere formatação de texto em um intervalo de células na planilha.

        Parâmetros:
        - caminho_planilha (str): Caminho da planilha Excel onde o Workbook será carregado.
        - range (str): Intervalo de células no formato 'A1:B2' ou no formato 'A1' sendo A1 o inicio da inserção dos dados sem especificar o final de onde sera inserido o dado. 
        - formato (str): Formatação do texto (ex.: 'bold italic').
        - style_font (str, opcional): Nome da fonte para aplicar nas células.
        - any_aba (any): Recebe o nome ou o indice da aba que será trabalhada.
        - flt_font_size (float): tamanho da fonte que será aplicada.
        - font_color (str):
        - horizontal_align (str): ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']
        - quebra_linha (bool): Caso seja true, será ativado na celula a opção quebra de linha automatica.

        Retorna:
        """
        try:
            # Carrega o workbook a partir da planilha
            wb_planilha = load_workbook(caminho_planilha)
            if isinstance(any_aba, int):
                ws_planilha = wb_planilha.worksheets[any_aba]
            else:
                ws_planilha = wb_planilha[any_aba]

            if range.__contains__(':'):
                # Intervalo de células em células inicial e final
                celula_inicial, celula_final = range.split(':')
            else:
                # Intervalo de células em células inicial e final
                celula_inicial = range
                celula_final = range

            # Itera sobre cada célula no intervalo
            for row in ws_planilha[celula_inicial:celula_final]:
                for cell in row:
                    # Cria um dicionário de formatação
                    formatacao = {}
                    alinhamento = {}

                    if 'bold' in formato:
                        formatacao['bold'] = True
                    if 'italic' in formato:
                        formatacao['italic'] = True
                    if 'underline' in formato:
                        formatacao['underline'] = 'single'
                    if 'strikethrough' in formato:
                        formatacao['strike'] = True
                    if style_font:
                        formatacao['name'] = style_font
                    if flt_font_size:
                        formatacao['size'] = flt_font_size
                    if font_color:
                        formatacao['color'] = font_color
                    if horizontal_align:
                        alinhamento['horizontal'] = horizontal_align
                    
                    alinhamento['wrap_text'] = quebra_linha
                        
                    # Aplica a formatação de fonte na célula
                    cell.font = Font(**formatacao)

                    # Aplica alinhamento nas celulas especificas
                    cell.alignment = Alignment(**alinhamento) 

            # Salva as alterações no arquivo
            wb_planilha.save(caminho_planilha)
            Log.write_log(f"Inserindo as formatações '{formato}' no intervalo {range}.")

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao inserir formatação: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao inserir formatação: {str(err)}")        

    @classmethod
    def alterar_largura_coluna(cls, caminho_planilha:str='',coluna:list[str]=[], flt_width_coluna:float=10.0, any_aba:any=0) -> None:
        """
        Método que altera a largura da coluna especificada nos argumentos.
        
        Parâmetros:
           - caminho_planilha(str): Caminho da planilha Excel onde o Workbook será carregado.
           - coluna(list[str]): lista de letras das colunas que sofrerão a alteração da largura.
           - flt_width_coluna(float): largura que será inserida na coluna.
           - any_aba(any): Recebe o nome ou o indice da aba que será trabalhada.

        Retorna:
        """
        try:
            wb_planilha = load_workbook(caminho_planilha)
            
            if isinstance(any_aba, int):
                ws_planilha = wb_planilha.worksheets[any_aba]
            else:
                ws_planilha = wb_planilha[any_aba]

            for coluna in coluna:
                ws_planilha.column_dimensions[coluna].width = flt_width_coluna

            wb_planilha.save(caminho_planilha)
            
            Log.write_log(f"Alterando a largura das '{len(coluna)}' colunas para {flt_width_coluna}.")
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao alterar largura da coluna: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao alterar largura da coluna: {str(err)}")
    
    @classmethod
    def alterar_altura_linha(cls, caminho_planilha:str='', linha:list[int]=[], flt_height_linha:float=10.0, any_aba:any=0) -> None:
        """
        Método que altera a altura da linha especificada nos argumentos.
        
        Parâmetros:
           - caminho_planilha(str): Caminho da planilha Excel onde o Workbook será carregado.
           - linha(list[int]): lista de inteiros que contem os numeros das linhas que sofrerão a alteração da altura.
           - flt_height_linha(float): altura que será inserida na linha.
           - any_aba(any): Recebe o nome ou o indice da aba que será trabalhada.
        
        Retorna:
        """
        try:
            wb_planilha = load_workbook(caminho_planilha)
            
            if isinstance(any_aba, int):
                ws_planilha = wb_planilha.worksheets[any_aba]
            else:
                ws_planilha = wb_planilha[any_aba]
            
            for linha in linha:
                ws_planilha.row_dimensions[linha].height = flt_height_linha

            wb_planilha.save(caminho_planilha)
            
            Log.write_log(f"Alterando a altura das '{len(linha)}' linhas para {flt_height_linha}.")
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao alterar altura das linhas: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao alterar altura das linhas: {str(err)}")
        
    @classmethod
    def mesclar_celulas(cls,caminho_planilha:str='', range:str='', any_aba:any=0) -> None:
        """
        Método que faz a mesclagem das celulas especificadas na range
        
        Parâmetros:
           - caminho_planilha(str): Caminho da planilha Excel onde o Workbook será carregado.
           - range(str): Intervalo de células no formato 'A1:B2'.
           - any_aba(any): Recebe o nome ou o indice da aba que será trabalhada.
        
        Retorna:
        """
        try:
            wb_planilha = load_workbook(caminho_planilha)
            
            if isinstance(any_aba, int):
                ws_planilha = wb_planilha.worksheets[any_aba]
            else:
                ws_planilha = wb_planilha[any_aba]

            ws_planilha.merge_cells(range)

            wb_planilha.save(caminho_planilha)
            
            Log.write_log(f"Mesclando celulas '{range}'")
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao mesclar celulas: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao mesclar celulas: {str(err)}")
        
    @classmethod
    def alterar_formato_dado(cls, caminho_planilha:str='',range:str='',formato_personalizado:str='',nome_formato_personalizado:str='custom',any_aba:any=0) -> None:
        """
        Metódo que altera o formato do intervalo de celulas no excel de maneira customizada.
        
        Parâmetros:
        - caminho_planilha(str): Caminho da planilha Excel onde o Workbook será carregado.
        - range (str): Intervalo de células no formato 'A1:B2' ou no formato 'A1' sendo A1 o inicio da inserção dos dados sem especificar o final de onde sera inserido o dado. 
        - formato_personalizado(str): formato que sera utilizado no intervalo especificado, por exemplo: [h]:mm:ss
        - nome_formato_personalizado(str): nome do formato, não tem padrão
        - any_aba(any): Recebe o nome ou o indice da aba que será trabalhada.
   
        Retorna:
        """
        try:
            wb_planilha = load_workbook(caminho_planilha)
            
            if isinstance(any_aba, int):
                ws_planilha = wb_planilha.worksheets[any_aba]
            else:
                ws_planilha = wb_planilha[any_aba]

            if range.__contains__(':'):
                # Intervalo de células em células inicial e final
                celula_inicial, celula_final = range.split(':')
            else:
                # Intervalo de células em células inicial e final
                celula_inicial = range
                celula_final = range

            for row in ws_planilha[celula_inicial:celula_final]:
                for cell in row:
                    try:
                        if nome_formato_personalizado not in wb_planilha.named_styles:
                            ns_new_style = NamedStyle(name=nome_formato_personalizado, number_format=formato_personalizado)
                            wb_planilha.add_named_style(ns_new_style)
                            
                        cell.style = nome_formato_personalizado
                    except Exception as err:
                        if 'exists already' in str(err):
                            cell.style = nome_formato_personalizado
                        else:
                            raise err

            # Salva as alterações no arquivo
            wb_planilha.save(caminho_planilha)
            
            Log.write_log(f"Inserindo formato personalizado '{formato_personalizado}' no intervalo {range}.")
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao inserir formato personalizado: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao inserir formato personalizado: {str(err)}")
        
    @classmethod
    def inserir_valores(cls, caminho_planilha:str, range:str='A1', valores:list[list[any]]=[[None]],any_aba:any=0) -> None:
        """
        Insere um valor em uma célula específica na planilha.

        Parâmetros:
        - caminho_planilha (str): Caminho da planilha Excel onde o Workbook será carregado.
        - range (str): Intervalo de células no formato 'A1:B2' ou no formato 'A1' sendo A1 o inicio da inserção dos dados sem especificar o final de onde sera inserido o dado. 
        - any_valor (list[any]): Valores a serem inseridos nas células.
        - any_aba (any): Recebe o nome ou o indice da aba que será trabalhada.

        Retorna:
        """
        try:
            # Carrega o workbook a partir da planilha
            wb_planilha = load_workbook(caminho_planilha)
            
            if isinstance(any_aba, int):
                ws_planilha = wb_planilha.worksheets[any_aba]
            else:
                ws_planilha = wb_planilha[any_aba]

            shape_valores = np.asarray(valores).shape
            Log.write_log(f'O shape da lista de valores é {shape_valores}')

            if range.__contains__(':'):
                celula_inicial, celula_final = range.split(':')
                shape_celulas = np.asarray(list(ws_planilha[celula_inicial:celula_final])).shape
            
                if shape_celulas != shape_valores:
                    raise Exception(f'O tamanho da lista de celulas é diferente do tamanho da lista de valores. Lista de Celulas {shape_celulas} | Lista de Valores {shape_valores}')
                
                # Insere o valor na célula
                for index_linha,row in enumerate(ws_planilha[celula_inicial:celula_final]):
                    for index_coluna,cell in enumerate(row):
                        cell.value = valores[index_linha][index_coluna]

            else:
                linha_inicial = ws_planilha[range].row
                coluna_inicial = ws_planilha[range].column

                for index_linha,linha in enumerate(range(linha_inicial,linha_inicial+shape_valores[0])):
                    for index_coluna,coluna in enumerate(range(coluna_inicial,coluna_inicial+shape_valores[1])):
                        ws_planilha.cell(row=linha,column=coluna,value=valores[index_linha][index_coluna])

            # Salva as alterações no arquivo
            wb_planilha.save(caminho_planilha)
            Log.write_log(f"Inserindo o valores no intervalo {range}.")

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao inserir valor: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao inserir valor: {str(err)}")
    
    @classmethod
    def config_linhas_grade(cls, caminho_planilha:str, ativar:bool=True, any_aba:any=0) -> None:
        """
        Desativa ou ativa as linhas de grade do excel

        Parâmetros:
        - caminho_planilha (str): Caminho da planilha Excel onde o Workbook será carregado.
        - ativar (bool): Caso seja True, ira ativar as linhas de grade, caso seja False vai desabilitar
        - any_aba (any): Recebe o nome ou o indice da aba que será trabalhada.

        Retorna:
        """
        try:
            # Carrega o workbook a partir da planilha
            wb_planilha = load_workbook(caminho_planilha)
                
            if isinstance(any_aba, int):
                ws_planilha = wb_planilha.worksheets[any_aba]
            else:
                ws_planilha = wb_planilha[any_aba]

            ws_planilha.sheet_view.showGridLines=ativar
            
            # Salva as alterações no arquivo
            wb_planilha.save(caminho_planilha)
            
            Log.write_log(f"Configuração das linhas de grade realizadas.")

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao configurar linhas de grade: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao configurar linhas de grade: {str(err)}")
    
    @classmethod
    def manipular_bordas(cls, caminho_planilha:str, range:str='A1', any_aba:any=0, border_stryle:str=None, lado_borda:list[str]=[],border_color:str='000000') -> None:
        """
        Desativa ou ativa as linhas de grade do excel

        Parâmetros:
        - caminho_planilha (str): Caminho da planilha Excel onde o Workbook será carregado.
        - range (str): Intervalo de células no formato 'A1:B2' ou no formato 'A1' sendo A1 o inicio da inserção dos dados sem especificar o final de onde sera inserido o dado. 
        - any_valor (list[any]): Valores a serem inseridos nas células.
        - any_aba (any): Recebe o nome ou o indice da aba que será trabalhada.
        - border_stryle (str): Estilo da borda que será utilizado, tipos aceitos:  ['dashDot','dashDotDot','dashed','dotted','double','hair', 'medium', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin']
        - lado_borda (list): Lado onde será aplicada a borda. Opções disponíveis: ['left', 'right', 'top', 'bottom']
        - border_color (str): Código Hexadecimal de cor para o preenchimento das células.

        Retorna:
        """
        try:
            # Carrega o workbook a partir da planilha
            wb_planilha = load_workbook(caminho_planilha)
            if isinstance(any_aba, int):
                ws_planilha = wb_planilha.worksheets[any_aba]
            else:
                ws_planilha = wb_planilha[any_aba]

            # Tipos de bordas disponiveis
            border_style= [None,'dashDot','dashDotDot', 'dashed','dotted',
                            'double','hair', 'medium', 'mediumDashDot', 'mediumDashDotDot',
                            'mediumDashed', 'slantDashDot', 'thick', 'thin']
            
            # Border sides disponiveis
            lados_borda_validos = ['left', 'right', 'top', 'bottom']
            

            if range.__contains__(':'):
                # Intervalo de células em células inicial e final
                celula_inicial, celula_final = range.split(':')
            else:
                # Intervalo de células em células inicial e final
                celula_inicial = range
                celula_final = range

            if border_stryle not in border_style:
                raise Exception('Estilo de borda não encontrado. Utilize algum descrito no comentario do metodo.')

            if len([item_atual for item_atual in lado_borda if item_atual not in lados_borda_validos]) > 0:
                raise Exception('Algum Border Side da lista não foi encontrado. Utilize apenas os descritos no comentario do metodo.')

            # Itera sobre cada célula no intervalo
            for row in ws_planilha[celula_inicial:celula_final]:
                for cell in row:
                    brd_current_border = cell.border if cell.border else Border()
                    sid_border_side_obj = Side(style=border_stryle,color=border_color)

                    # Atualiza a borda conforme o lado
                    if "top" in lado_borda:
                        cell.border = Border(top=sid_border_side_obj, bottom=brd_current_border.bottom, left=brd_current_border.left, right=brd_current_border.right)
                        brd_current_border = cell.border if cell.border else Border()

                    if "bottom" in lado_borda:
                        cell.border = Border(top=brd_current_border.top, bottom=sid_border_side_obj, left=brd_current_border.left, right=brd_current_border.right)
                        brd_current_border = cell.border if cell.border else Border()
                        
                    if "left" in lado_borda:
                        cell.border = Border(top=brd_current_border.top, bottom=brd_current_border.bottom, left=sid_border_side_obj, right=brd_current_border.right)
                        brd_current_border = cell.border if cell.border else Border()

                    if "right" in lado_borda:
                        cell.border = Border(top=brd_current_border.top, bottom=brd_current_border.bottom, left=brd_current_border.left, right=sid_border_side_obj)
                        brd_current_border = cell.border if cell.border else Border()
   
            # Salva as alterações no arquivo
            wb_planilha.save(caminho_planilha)
            
            Log.write_log(f"Bordas inseridas com sucesso.")

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao inserir bordas: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao inserir bordas: {str(err)}")