from openpyxl.styles import Font, PatternFill, Border, Side

def formatar_cabecalho(planilha):
    """
    Vai formatar o cabeçalho da planilha colocando em negrito e fundo cinza
    """
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDDDDD")  # cinza claro, altere conforme necessário
    for cell in planilha[1]:
        cell.font = header_font
        cell.fill = header_fill

def ajustar_largura_colunas(planilha):
    """
    Ajusta Largura e as colunas
    """
    for column in planilha.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try: 
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
            adjusted_width = (max_length + 2)
            planilha.column_dimensions[column[0].column_letter].width = adjusted_width

def adicionar_bordas(planilha):
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))
    for row in planilha:
        for cell in row:
            cell.border = thin_border

def remover_gridlines(planilha):
    planilha.sheet_view.showGridLines = False

def formatar_cpf(cpf):
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"

