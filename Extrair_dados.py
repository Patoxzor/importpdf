from PyPDF2 import PdfReader
import re
import pandas as pd

# Mapeamento de códigos para nomes de campos de "proventos"
proventos_codes = {
    '1': 'SALARIO BASE',
    '66': 'GRATIFICACAO',
    '421': 'INCENTIVO FINANCEIRO'
    # Adicione mais códigos de proventos aqui
}

# Mapeamento de códigos para nomes de campos de "descontos"
descontos_codes = {
    '389': 'JUNDIA-PREV',
    '300': 'INSS'
    # Adicione mais códigos de descontos aqui
}

def extract_field(text, pattern):
    match = re.search(pattern, text, re.MULTILINE)
    return match.group(1) if match else None

def extrair_cadastro(text):
    patterns = {
        'codigo': r'FUNCIONÁRIO: (.*) -',
        'nome': r'FUNCIONÁRIO: \d+\s*-+\s*(.*) PIS/PASEP',
        'centro_c': r'CENTRO C.: (.*?) SECRETARIA',
        'secretaria': r'SECRETARIA: (.*?)\nCARGO:',
        'cargo': r'CARGO: (.*) CPF/AG-C.C',
        'cpf': r'CPF/AG-C.C: (.*) /',
        'ag': r'AG: (\d+)',
        'cc': r'CC: (\d+)',
        'admissao': r'ADMISSÃO: (.*) Nascimento',
        'nascimento': r'Nascimento:\s*(\d{2}/\d{2}/\d{4})',
    }

    results = {}
    for field, pattern in patterns.items():
        results[field] = extract_field(text, pattern)

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
    return [results.get(field, None) for field in patterns.keys()] + [results.get(field + suffix, None) for field in proventos_codes.values() for suffix in ['_QTD', '_VALOR']] + [results.get(field + suffix, None) for field in descontos_codes.values() for suffix in ['_QTD', '_VALOR']]

def extrair_dados_pdf(f):
    pdf = PdfReader(f)

    data = []
    for page in pdf.pages:
        text = page.extract_text()  # Use extract_text()

        sections = text.split('FUNCIONÁRIO:')
        for section in sections[1:]:  # Ignorar a primeira seção, pois estará vazia
            data.append(extrair_cadastro('FUNCIONÁRIO:' + section))  # Adicionar 'FUNCIONÁRIO:' de volta ao início da seção

    # Criar um DataFrame com os resultados
    df = pd.DataFrame(data, columns=['Codigo','Nome', 'Centro C.', 'Secretaria', 'Cargo', 'CPF', 'AG', 'CC', 'Admissão', 'Nascimento'] + 
                           [field + suffix for field in proventos_codes.values() for suffix in ['_QTD', '_VALOR']] +
                           [field + suffix for field in descontos_codes.values() for suffix in ['_QTD', '_VALOR']])
    return df

