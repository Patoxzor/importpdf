from PyPDF2 import PdfReader
import re
import pandas as pd

# Funções auxiliares permanecem as mesmas
def extract_field(text, pattern):
    match = re.search(pattern, text, re.MULTILINE)
    return match.group(1) if match else None

def extrair_cadastro(text, codigos_proventos, codigos_desconto):
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
    for code, field_name in codigos_proventos.items():
        pattern = rf'{code}\s*{field_name}\s*(.*?)\s*(\d{{1,3}}(?:\.\d{{3}})*,\d{{2}})\s*0,00'
        match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
        if match:
            results[field_name + '_QTD'] = match.group(1)
            results[field_name + '_VALOR'] = match.group(2)

    # Extrair campos de descontos
    for code, field_name in codigos_desconto.items():
        pattern = rf'{code}\s*{field_name}\s*(.*?)\s*0,00\s*(.*?)\s*$'
        match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
        if match:
            results[field_name + '_QTD'] = match.group(1)
            results[field_name + '_VALOR'] = match.group(2)

    # Adicione os novos campos à lista de resultados
    return [results.get(field, None) for field in patterns.keys()] + [results.get(field + suffix, None) for field in codigos_proventos.values() for suffix in ['_QTD', '_VALOR']] + [results.get(field + suffix, None) for field in codigos_desconto.values() for suffix in ['_QTD', '_VALOR']]

def limpar_qtd_colunas(df):
    def limpar_celula(valor_celula):
        if isinstance(valor_celula, str):
            return re.sub(r'[^0-9,.]', '', valor_celula)
        return valor_celula

    for col in df.columns:
        if col.endswith("_QTD"):
            df[col] = df[col].apply(limpar_celula)
    return df

# O principal ajuste é passar os dicionários `codigos_proventos` e `codigos_desconto` como parâmetros para a função `extrair_dados_pdf`
def extrair_dados_pdf(f, codigos_proventos, codigos_desconto, mapeamento_codigos=None):
    pdf = PdfReader(f)

    data = []
    for page in pdf.pages:
        text = page.extract_text()
        sections = text.split('FUNCIONÁRIO:')
        for section in sections[1:]:
            data.append(extrair_cadastro('FUNCIONÁRIO:' + section, codigos_proventos, codigos_desconto))

    df = pd.DataFrame(data, columns=['Codigo', 'Nome', 'Centro C.', 'Secretaria', 'Cargo', 'CPF', 'AG', 'CC', 'Admissão', 'Nascimento'] +
                      [field + suffix for field in codigos_proventos.values() for suffix in ['_QTD', '_VALOR']] +
                      [field + suffix for field in codigos_desconto.values() for suffix in ['_QTD', '_VALOR']])
    if mapeamento_codigos:
        df['Codigo'] = df['Codigo'].map(mapeamento_codigos).fillna(df['Codigo'])


    df = limpar_qtd_colunas(df)
    return df
