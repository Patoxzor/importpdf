import tkinter as tk
from tkinter import filedialog
from IMPORT_PDF_MARKAS import extrair_dados_pdf

# Cria uma janela raiz e a esconde
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
    # Agora você pode usar esses caminhos com suas funções existentes
    df = extrair_dados_pdf(f)
    df.to_excel(caminho_excel, index=False)
