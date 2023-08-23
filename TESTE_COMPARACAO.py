from PyPDF2 import PdfReader
import re
import pandas as pd
from tkinter import filedialog
from tkinter import Tk

def extract_field(text, pattern):
    match = re.search(pattern, text, re.MULTILINE)
    return match.group(1) if match else None

def extrair_cadastro(text):
    patterns = {
        'funcionario': r'FUNCIONÁRIO: (.*) PIS/PASEP',       
        'cpf': r'CPF/AG-C.C: (.*) /',
        'admissao': r'ADMISSÃO: (.*) Nascimento',
        'nascimento': r'Nascimento:\s*(\d{2}/\d{2}/\d{4})',
        'secretaria': r'SECRETARIA: (.*?)\nCARGO:',
        'centro_custo': r'CENTRO C.: (.*?) SECRETARIA',
        'cargo': r'CARGO: (.*) CPF/AG-C.C',
        'agencia': r'AG: (\d+)',
        'conta': r'CC: (\d+)'
        
    }

    results = {}
    for field, pattern in patterns.items():
        results[field] = extract_field(text, pattern)

    return [results.get(field, None) for field in patterns.keys()]

def extrair_dados_pdf(f):
    pdf = PdfReader(f)

    data = []
    for page in pdf.pages:
        text = page.extract_text() 

        sections = text.split('FUNCIONÁRIO:')
        for section in sections[1:]:  
            data.append(extrair_cadastro('FUNCIONÁRIO:' + section))  

    df = pd.DataFrame(data, columns=['Funcionário', 'CPF', 'Admissao', 'Nascimento', 'Secretaria','centro_custo', 'Cargo', 'Agencia', 'Conta'])
    return df

def comparar_pdfs(pdf1, pdf2, excel_file):
    df1 = extrair_dados_pdf(pdf1)
    df2 = extrair_dados_pdf(pdf2)

    funcionarios_df1 = set(zip(df1['Funcionário'], df1['CPF'], df1['Admissao'], df1['Nascimento']))
    funcionarios_df2 = set(zip(df2['Funcionário'], df2['CPF'], df2['Admissao'], df2['Nascimento']))

    apenas_df1 = funcionarios_df1 - funcionarios_df2
    apenas_df2 = funcionarios_df2 - funcionarios_df1

    df_apenas_df1 = df1[df1.set_index(['Funcionário', 'CPF', 'Admissao', 'Nascimento']).index.isin(apenas_df1)]
    df_apenas_df2 = df2[df2.set_index(['Funcionário', 'CPF', 'Admissao', 'Nascimento']).index.isin(apenas_df2)]

    with pd.ExcelWriter(excel_file) as writer:
        df_apenas_df1.to_excel(writer, sheet_name='Apenas no PDF1', index=False)
        df_apenas_df2.to_excel(writer, sheet_name='Apenas no PDF2', index=False)

root = Tk()
root.withdraw()

pdf1 = filedialog.askopenfilename(title = "Selecione o primeiro PDF")
pdf2 = filedialog.askopenfilename(title = "Selecione o segundo PDF")
excel_file = filedialog.asksaveasfilename(title = "Salvar arquivo Excel", defaultextension=".xlsx")

comparar_pdfs(pdf1, pdf2, excel_file)
