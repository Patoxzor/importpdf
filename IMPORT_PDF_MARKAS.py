from PyPDF2 import PdfReader
import re
import pandas as pd

# Mapeamento de códigos para nomes de campos de "proventos"
proventos_codes = {
        '107': 'GRATIFICAÇÃO',
        '110': 'QUINQUENIO',
        '421': 'INCENTIVO FINANCEIRO'
        # Adicione mais códigos de proventos aqui
    }

    # Mapeamento de códigos para nomes de campos de "descontos"
descontos_codes = {
        '389': 'JUNDIA-PREV',
        # Adicione mais códigos de descontos aqui
    }

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

    # Dicionário para armazenar os resultados
    results = {}

    # Extrair campos de proventos
    for code, field_name in proventos_codes.items():
        pattern = rf'{code}\s*{field_name}\s*(.*?)\s*(\d{{1,3}}(?:\.\d{{3}})*,\d{{2}})\s*0,00'
        match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
        if match:
            results[field_name + '_QTD'] = match.group(1)
            results[field_name + '_VALOR'] = match.group(2)

    # Extrair campos de descontos
    for code, field_name in descontos_codes.items():
        pattern = rf'{code}\s*{field_name}\s*(.*?)\s*0,00\s*(.*?)\s*$'
        match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
        if match:
            results[field_name + '_QTD'] = match.group(1)
            results[field_name + '_VALOR'] = match.group(2)

    # Adicione os novos campos à lista de resultados
    return [funcionario, centro_c, secretaria, cargo, cpf, ag, cc, admissao, nascimento, salario] + [results.get(field + suffix, None) for field in proventos_codes.values() for suffix in ['_QTD', '_VALOR']] + [results.get(field + suffix, None) for field in descontos_codes.values() for suffix in ['_QTD', '_VALOR']]

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
    df = pd.DataFrame(data, columns=['Funcionário', 'Centro C.', 'Secretaria', 'Cargo', 'CPF', 'AG', 'CC', 'Admissão', 'Nascimento', 'Salario'] + 
                           [field + suffix for field in proventos_codes.values() for suffix in ['_QTD', '_VALOR']] +
                           [field + suffix for field in descontos_codes.values() for suffix in ['_QTD', '_VALOR']])


    # Escrever o DataFrame para um arquivo Excel
    df.to_excel('output.xlsx', index=False)
