from PyPDF2 import PdfReader
import re
import pandas as pd

def extract_fields(text):
    funcionario = re.search(r'FUNCIONÁRIO: (.*) PIS/PASEP', text)
    centro_c = re.search(r'CENTRO C.: (.*?) SECRETARIA', text)
    secretaria = re.search(r'SECRETARIA: (.*?)\nCARGO:', text, re.DOTALL)
    cargo = re.search(r'CARGO: (.*) CPF/AG-C.C', text)
    cpf = re.search(r'CPF/AG-C.C: (.*) /', text)
    ag = re.search(r'AG: (\d+)', text)
    cc = re.search(r'CC: (\d+)', text)
    admissao = re.search(r'ADMISSÃO: (.*) Nascimento', text)
    nascimento = re.search(r'Nascimento:\s*(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
    salario = re.search(r'100\s*SALARIO.*?(\d{1,3}(?:\.\d{3})*,\d{2}).*$', text, re.MULTILINE)
    inss = re.search(r'300\s*INSS\s*(.*?)\s*0,00\s*(.*?)\s*$', text, re.MULTILINE)
    gratificacao = re.search(r'107\s*GRATIFICAÇÃO\s*(.*?)\s*(\d{1,3}(?:\.\d{3})*,\d{2})\s*0,00', text, re.MULTILINE)


    funcionario = funcionario.group(1) if funcionario else None
    centro_c = centro_c.group(1) if centro_c else None
    secretaria = secretaria.group(1) if secretaria else None
    cargo = cargo.group(1) if cargo else None
    cpf = cpf.group(1) if cpf else None
    ag = ag.group(1) if ag else None
    cc = cc.group(1) if cc else None
    admissao = admissao.group(1) if admissao else None
    nascimento = nascimento.group(1) if nascimento else None
    salario = salario.group(1) if salario else None
    percentage = inss.group(1) if inss else None
    value = inss.group(2) if inss else None
    gratificacao_qtd = gratificacao.group(1) if gratificacao else None
    gratificacao_valor = gratificacao.group(2) if gratificacao else None

    return [funcionario, centro_c, secretaria, cargo, cpf, ag, cc, admissao, nascimento, salario, percentage, value, gratificacao_qtd, gratificacao_valor]

# Abra o arquivo PDF
with open('D://ESOCIAL//04-2023 - DETALHADA.pdf', 'rb') as f:
    pdf = PdfReader(f)

    data = []
    for page in pdf.pages:
        text = page.extract_text()  # Use extract_text()

        sections = text.split('FUNCIONÁRIO:')
        for section in sections[1:]:  # Ignorar a primeira seção, pois estará vazia
            data.append(extract_fields('FUNCIONÁRIO:' + section))  # Adicionar 'FUNCIONÁRIO:' de volta ao início da seção

    # Criar um DataFrame com os resultados
    df = pd.DataFrame(data, columns=['Funcionário', 'Centro C.', 'Secretaria', 'Cargo', 'CPF', 'AG', 'CC', 'Admissão', 'Nascimento', 'Salario', 'QTD', 'INSS', 'QTD_GRAT', 'GRATIFICACAO'])

    # Escrever o DataFrame para um arquivo Excel
    df.to_excel('output.xlsx', index=False)
