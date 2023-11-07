from PyPDF2 import PdfReader
import re
import pandas as pd
from tkinter import filedialog
from tkinter import Tk
from excel_format import formatar_cpf

def extract_field(text, pattern):
    match = re.search(pattern, text, re.MULTILINE)
    return match.group(1) if match else None

def extrair_cadastro(text):
    patterns = {
        'codigo': r'FUNCIONÁRIO: (.*) -',
        'nome': r'FUNCIONÁRIO: \d+\s*-+\s*(.*) PIS/PASEP',     
        'cpf': r'CPF/AG-C.C: (.*) /',
        'admissao': r'ADMISSÃO: (.*) Nascimento',
        'nascimento': r'Nascimento:\s*(\d{2}/\d{2}/\d{4})',
        'secretaria': r'SECRETARIA: (.*?)\nCARGO:',
        'centro_custo': r'CENTRO C.: (.*?) SECRETARIA',
        'cargo': r'CARGO: (.*) CPF/AG-C.C',
        'salario': r'100\s*SALARIO.*?(\d{1,3}(?:\.\d{3})*,\d{2}).*$',
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

    df = pd.DataFrame(data, columns=['Codigo','Nome', 'CPF', 'Admissao', 'Nascimento', 'Secretaria','centro_custo', 'Cargo','Salario', 'Agencia', 'Conta'])
    return df

def comparar_pdfs(pdf1, outros_pdfs, arquivo_excel):
    df1 = extrair_dados_pdf(pdf1)

    pdf_base_df1 = pd.DataFrame(columns=df1.columns)
    outros_pdfs_df2 = pd.DataFrame(columns=df1.columns)

    global_set= set()
    
    for pdf in outros_pdfs:
        df2 = extrair_dados_pdf(pdf)

        funcionarios_df2 = set(zip(df2['Codigo'], df2['CPF'], df2['Admissao'], df2['Nascimento']))
        global_set = global_set.union(funcionarios_df2)

    funcionarios_df1 = set(zip(df1['Codigo'], df1['CPF'], df1['Admissao'], df1['Nascimento']))
    apenas_df1 = funcionarios_df1 - global_set
    apenas_outros = global_set - funcionarios_df1

    pdf_base_df1 = df1[df1.set_index(['Codigo', 'CPF', 'Admissao', 'Nascimento']).index.isin(apenas_df1)]
    pdf_base_df1.loc[:, 'CPF'] = pdf_base_df1['CPF'].apply(formatar_cpf)

    for pdf in outros_pdfs:
        df2 = extrair_dados_pdf(pdf)
        df_apenas_outros = df2[df2.set_index(['Codigo', 'CPF', 'Admissao', 'Nascimento']).index.isin(apenas_outros)]
        outros_pdfs_df2 = pd.concat([outros_pdfs_df2, df_apenas_outros])
        outros_pdfs_df2.loc[:, 'CPF'] = outros_pdfs_df2['CPF'].apply(formatar_cpf)

    with pd.ExcelWriter(arquivo_excel) as writer:
        pdf_base_df1.drop_duplicates(subset=['CPF', 'Codigo']).to_excel(writer, sheet_name='Apenas no PDF1', index=False)
        outros_pdfs_df2.drop_duplicates(subset=['CPF', 'Codigo']).to_excel(writer, sheet_name='Não estão no PDF1', index=False)

root = Tk()
root.withdraw()

pdf1 = filedialog.askopenfilename(title = "Selecione o primeiro PDF")
outros_pdfs = filedialog.askopenfilenames(title = "Selecione os outros PDFS")
arquivo_excel = filedialog.asksaveasfilename(title = "Salvar arquivo Excel", defaultextension=".xlsx")

comparar_pdfs(pdf1, outros_pdfs, arquivo_excel)
