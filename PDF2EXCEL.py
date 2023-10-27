import tkinter as tk
from tkinter import filedialog
from Extrair_dados import extrair_dados_pdf
import pandas as pd
from excel_format import formatar_cabecalho, formatar_cpf, ajustar_largura_colunas, adicionar_bordas, remover_gridlines

root = tk.Tk()
root.withdraw()

# Solicita ao usuário para selecionar o arquivo PDF
caminho_pdf= filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
print(f"Arquivo PDF selecionado: {caminho_pdf}")

# Solicita ao usuário para selecionar o local para salvar o arquivo Excel
caminho_excel = filedialog.asksaveasfilename(defaultextension=".xlsx")
print(f"Arquivo Excel será salvo em: {caminho_excel}")

# Abra o arquivo PDF
with open(caminho_pdf, 'rb') as f:
    df = extrair_dados_pdf(f)
    df['CPF'] = df['CPF'].apply(formatar_cpf)

# Formatar e salvar o excel
with pd.ExcelWriter(caminho_excel, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Planilha Completa', index=False)
    arquivo_excel = writer.sheets['Planilha Completa']
    formatar_cabecalho(arquivo_excel)
    ajustar_largura_colunas(arquivo_excel)
    adicionar_bordas(arquivo_excel)
    remover_gridlines(arquivo_excel)
