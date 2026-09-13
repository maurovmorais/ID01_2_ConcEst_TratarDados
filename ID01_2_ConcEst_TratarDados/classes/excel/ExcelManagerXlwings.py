# Imports dos módulos internos do projeto
# Carrega o InitAllSettingsSettings Precisa ser o primeiro a ser carregado
from ID01_2_ConcEst_TratarDados.classes.framework.InitAllSettings import InitAllSettings as InitAllSetting
from ID01_2_ConcEst_TratarDados.classes.utils.Log import Log, LogLevel, ErrorType
# Imports dos pacotes externos
import xlwings as xw

class ExcelManagerXlwings:
    """
    Classe para gerenciar operações comuns em planilhas Excel usando xlwings.

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
        - range (str): Intervalo de células no formato 'A1:B2'.
        - any_aba (any): Recebe o nome ou o índice da aba que será trabalhada.

        Retorna:
        """
        try:
            app_xw = xw.App(visible=False)
            wb_planilha = xw.Book(caminho_planilha)
            ws_planilha = wb_planilha.sheets[any_aba]

            # Converte o código hexadecimal para o formato RGB
            cor = tuple(int(cor_celula[i:i+2], 16) for i in (0, 2, 4))

            # Aplica a cor no intervalo de células
            ws_planilha.range(range).color = cor

            wb_planilha.save()
            wb_planilha.close()
            app_xw.quit()
            
            Log.write_log(f"Inserindo a cor {cor_celula} no intervalo {range}.")

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao inserir cor: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao inserir cor: {str(err)}")

    @classmethod
    def inserir_formatacao_texto(cls, caminho_planilha:str, any_aba:any=0, range:str = '',
                                formato:str = '', style_font:str = None, flt_font_size:float=12.0,
                                font_color:str = '000000', horizontal_align:str='center', quebra_linha=False) -> None:
        """
        Insere formatação de texto em um intervalo de células na planilha.

        Parâmetros:
        - caminho_planilha (str): Caminho da planilha Excel onde o Workbook será carregado.
        - any_aba (any): Recebe o nome ou o índice da aba que será trabalhada.
        - range (str): Intervalo de células no formato 'A1:B2'.
        - formato (str): Formatação do texto (ex.: 'bold italic').
        - style_font (str, opcional): Nome da fonte para aplicar nas células.
        - flt_font_size (float): Tamanho da fonte que será aplicada.
        - font_color (str): Código Hexadecimal da cor da fonte.
        - horizontal_align (str): Alinhamento horizontal do texto.
        - quebra_linha (bool): Caso seja true, será ativado na célula a opção quebra de linha automática.

        Retorna:
        """
        try:
            app_xw = xw.App(visible=False)
            wb_planilha = xw.Book(caminho_planilha)
            ws_planilha = wb_planilha.sheets[any_aba]

            xw_range = ws_planilha.range(range)
            
            # Define a formatação
            if 'bold' in formato:
                xw_range.api.Font.Bold = True
            if 'italic' in formato:
                xw_range.api.Font.Italic = True
            if 'underline' in formato:
                xw_range.api.Font.Underline = True
            if 'strikethrough' in formato:
                xw_range.api.Font.Strikethrough = True
            if style_font:
                xw_range.api.Font.Name = style_font
            if flt_font_size:
                xw_range.api.Font.Size = flt_font_size
            if font_color:
                # Convertendo o código hexadecimal em um valor inteiro RGB
                cor = int(font_color, 16)
                xw_range.api.Font.Color = cor

            alinhamento = {'general': -4108, 'left': -4131, 'center': -4108, 'right': -4152, 'fill': 5, 'justify': -4130,
                                'centerContinuous': 7, 'distributed': -4117}
                        
            if horizontal_align in alinhamento:
                xw_range.api.HorizontalAlignment = alinhamento[horizontal_align]
            
            if quebra_linha:
                xw_range.api.WrapText = True

            wb_planilha.save()
            wb_planilha.close()
            app_xw.quit()

            Log.write_log(f"Inserindo as formatações '{formato}' no intervalo {range}.")
            
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao inserir formatação: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao inserir formatação: {str(err)}")

    @classmethod
    def alterar_largura_coluna(cls, caminho_planilha:str, coluna:list=[], flt_largura_coluna:float=10.0, any_aba:any=0) -> None:
        """
        Altera a largura de colunas especificadas.

        Parâmetros:
        - caminho_planilha (str): Caminho da planilha Excel onde o Workbook será carregado.
        - coluna (list): Lista de letras das colunas que terão a largura alterada.
        - flt_largura_coluna (float): Nova largura das colunas.
        - any_aba (any): Recebe o nome ou o índice da aba que será trabalhada.

        Retorna:
        """
        try:
            app_xw = xw.App(visible=False)
            wb_planilha = xw.Book(caminho_planilha)
            ws_planilha = wb_planilha.sheets[any_aba]

            for coluna in coluna:
                ws_planilha.range(f'{coluna}:{coluna}').column_width = flt_largura_coluna

            wb_planilha.save()
            wb_planilha.close()
            app_xw.quit()

            Log.write_log(f"Alterando a largura das '{len(coluna)}' colunas para {flt_largura_coluna}.")
            
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao alterar largura da coluna: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao alterar largura da coluna: {str(err)}")

    @classmethod
    def alterar_altura_linha(cls, caminho_planilha:str, linha:list=[], flt_altura_linha:float=10.0, any_aba:any=0) -> None:
        """
        Altera a altura de linhas especificadas.

        Parâmetros:
        - caminho_planilha (str): Caminho da planilha Excel onde o Workbook será carregado.
        - linha (list): Lista de números das linhas que terão a altura alterada.
        - flt_altura_linha (float): Nova altura das linhas.
        - any_aba (any): Recebe o nome ou o índice da aba que será trabalhada.

        Retorna:
        """
        try:
            app_xw = xw.App(visible=False)
            wb_planilha = xw.Book(caminho_planilha)
            ws_planilha = wb_planilha.sheets[any_aba]

            for linha in linha:
                ws_planilha.range(f'{linha}:{linha}').row_height = flt_altura_linha

            wb_planilha.save()
            wb_planilha.close()
            app_xw.quit()

            Log.write_log(f"Alterando a altura das '{len(linha)}' linhas para {flt_altura_linha}.")

        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao alterar altura das linhas: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao alterar altura das linhas: {str(err)}")

    @classmethod
    def mesclar_celulas(cls, caminho_planilha:str, range:str, any_aba:any=0) -> None:
        """
        Mescla células em um intervalo especificado.

        Parâmetros:
        - caminho_planilha (str): Caminho da planilha Excel onde o Workbook será carregado.
        - range (str): Intervalo de células no formato 'A1:B2'.
        - any_aba (any): Recebe o nome ou o índice da aba que será trabalhada.

        Retorna:
        """
        try:
            app_xw = xw.App(visible=False)
            wb_planilha = xw.Book(caminho_planilha)
            ws_planilha = wb_planilha.sheets[any_aba]

            ws_planilha.range(range).merge()

            wb_planilha.save()
            wb_planilha.close()
            app_xw.quit()

            Log.write_log(f"Mesclando celulas '{range}'")
            
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao mesclar celulas: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao mesclar celulas: {str(err)}")

    @classmethod
    def alterar_formato_dado(cls, caminho_planilha:str, range:str, formato_personalizado:str,
                             nome_formato_personalizado:str = 'custom', any_aba:any=0) -> None:
        """
        Altera o formato de dados em um intervalo especificado.

        Parâmetros:
        - caminho_planilha (str): Caminho da planilha Excel onde o Workbook será carregado.
        - range (str): Intervalo de células no formato 'A1:B2'.
        - formato_personalizado (str): Formato personalizado a ser aplicado.
        - nome_formato_personalizado (str): Nome do formato personalizado.
        - any_aba (any): Recebe o nome ou o índice da aba que será trabalhada.

        Retorna:
        """
        try:
            app_xw = xw.App(visible=False)
            wb_planilha = xw.Book(caminho_planilha)
            ws_planilha = wb_planilha.sheets[any_aba]

            ws_planilha.range(range).number_format = formato_personalizado

            wb_planilha.save()
            wb_planilha.close()
            app_xw.quit()

            Log.write_log(f"Inserindo formato personalizado '{formato_personalizado}' no intervalo {range}.")
            
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao inserir formato personalizado: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao inserir formato personalizado: {str(err)}")

    @classmethod
    def inserir_valores(cls, caminho_planilha:str, range:str, valores:list[list[any]], any_aba:any=0) -> None:
        """
        Insere valores em um intervalo de células.

        Parâmetros:
        - caminho_planilha (str): Caminho da planilha Excel onde o Workbook será carregado.
        - range (str): Intervalo de células no formato 'A1:B2'.
        - valores (list[list[any]]): Valores a serem inseridos nas células.
        - any_aba (any): Recebe o nome ou o índice da aba que será trabalhada.

        Retorna:
        """
        try:
            app_xw = xw.App(visible=False)
            wb_planilha = xw.Book(caminho_planilha)
            ws_planilha = wb_planilha.sheets[any_aba]

            ws_planilha.range(range).value = valores

            wb_planilha.save()
            wb_planilha.close()
            app_xw.quit()

            Log.write_log(f"Inserindo o valor no intervalo {range}.")
            
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao inserir valor: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao inserir valor: {str(err)}")
        
    @classmethod
    def manipular_bordas(cls, caminho_planilha:str, range:str='A1', any_aba:any=0, estilo_borda:str=None, lado_borda:list[str]=[], border_color:str='000000') -> None:
        """
        Configura bordas em um intervalo de células.

        Parâmetros:
        - caminho_planilha (str): Caminho da planilha Excel onde o Workbook será carregado.
        - range (str): Intervalo de células no formato 'A1:B2'.
        - any_aba (any): Recebe o nome ou o índice da aba que será trabalhada.
        - estilo_borda (str): Estilo da borda que será utilizado, tipos aceitos:  ['dashDot','dashDotDot','dashed','dotted','double','hair', 'medium', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin']
        - lado_borda (list): Lado onde será aplicada a borda. Opções disponíveis: ['left', 'right', 'top', 'bottom']
        - border_color (str): Código Hexadecimal de cor para a borda.

        Retorna:
        """
        try:
            app_xw = xw.App(visible=False)
            wb_planilha = xw.Book(caminho_planilha)
            ws_planilha = wb_planilha.sheets[any_aba]

            # Converte o código hexadecimal para um valor inteiro RGB
            cor_borda = int(border_color, 16)

            # Define o estilo da borda, se fornecido
            estilos_borda = {
                'dashDot': 9, 'dashDotDot': 13, 'dashed': -4115, 'dotted': -4118,
                'double': -4119, 'hair': 1, 'medium': -4138, 'mediumDashDot': 10,
                'mediumDashDotDot': 11, 'mediumDashed': -4124, 'slantDashDot': 14,
                'thick': 4, 'thin': 2
            }
            estilo_borda_codigo = estilos_borda.get(estilo_borda, 1)  # Default to 'hair' if not found

            xw_range = ws_planilha.range(range)

            # Aplicar bordas
            for lado in lado_borda:
                if lado == 'left':
                    xw_borda = xw_range.api.Borders(1)
                elif lado == 'right':
                    xw_borda = xw_range.api.Borders(2)
                elif lado == 'top':
                    xw_borda = xw_range.api.Borders(3)
                elif lado == 'bottom':
                    xw_borda = xw_range.api.Borders(4)
                else:
                    raise Exception("Lado de borda desconhecido: " + lado)

                xw_borda.LineStyle = estilo_borda_codigo
                xw_borda.Color = cor_borda

            wb_planilha.save()
            wb_planilha.close()
            app_xw.quit()

            Log.write_log(f"Inserindo bordas no intervalo {range}.")
                
        except Exception as err:
            Log.write_log(mensagem_log=f"Erro ao inserir bordas: {str(err)}", log_level=LogLevel.ERROR, error_type=ErrorType.APP_ERROR)
            raise Exception(f"Erro ao inserir bordas: {str(err)}")