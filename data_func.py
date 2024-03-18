from PyPDF2 import PdfReader
import re
import pandas as pd

def extract_field(text, pattern):
    match = re.search(pattern, text, re.MULTILINE)
    return match.group(1) if match else ""


def extrair_cadastro(text, codigos_proventos, codigos_desconto):
    patterns = {
        'Codigo': r'FUNCIONÁRIO: (.*) -',
        'Nome': r'FUNCIONÁRIO: \d+\s*-+\s*(.*) PIS/PASEP',
        'Centro_c': r'CENTRO C.: (.*?) SECRETARIA',
        'Secretaria': r'SECRETARIA: (.*?)\nCARGO:',
        'Cargo': r'CARGO: (.*) CPF/AG-C.C',
        'CPF': r'CPF/AG-C.C: (.*) /',
        'ag': r'AG: (\d+)',
        'cc': r'CC: (\d+)',
        'Vinculo': r'VÍNCULO E.: (.*?) ADMISSÃO',
        'Admissao': r'ADMISSÃO: (.*) Nascimento',
        'Nascimento': r'Nascimento:\s*(\d{2}/\d{2}/\d{4})',
    }

    results = {}
    for field, pattern in patterns.items():
        results[field] = extract_field(text, pattern)

    eventos_proventos_ocorrencias = {}
    for code, field_name in codigos_proventos.items():
        for match in re.finditer(rf'{code}\s*{field_name}\s*(.*?)\s*(\d{{1,3}}(?:\.\d{{3}})*,\d{{2}})\s*0,00', text, re.MULTILINE | re.IGNORECASE):
            ocorrencia = eventos_proventos_ocorrencias.get(field_name, 0) + 1
            eventos_proventos_ocorrencias[field_name] = ocorrencia

            coluna_qtd = f'{field_name}_QTD'
            coluna_valor = f'{field_name}_VALOR'
            if ocorrencia > 1:
                coluna_qtd = f'{coluna_qtd}_{ocorrencia}'
                coluna_valor = f'{coluna_valor}_{ocorrencia}'

            results[coluna_qtd] = match.group(1)
            results[coluna_valor] = match.group(2)

     # Extrair campos de descontos com controle de eventos repetidos
    eventos_ocorrencias = {}
    for code, field_name in codigos_desconto.items():
        for match in re.finditer(rf'{code}\s*{field_name}\s*(.*?)\s*0,00\s*(.*?)\s*$', text, re.MULTILINE | re.IGNORECASE):
            ocorrencia = eventos_ocorrencias.get(field_name, 0) + 1
            eventos_ocorrencias[field_name] = ocorrencia

            coluna_qtd = f'{field_name}_QTD'
            coluna_valor = f'{field_name}_VALOR'
            if ocorrencia > 1:
                coluna_qtd = f'{coluna_qtd}_{ocorrencia}'
                coluna_valor = f'{coluna_valor}_{ocorrencia}'

            results[coluna_qtd] = match.group(1)
            results[coluna_valor] = match.group(2)

    return results

def limpar_celula(valor_celula):
    # Verifica se o valor é uma string e, se for, limpa os caracteres indesejados
    if pd.isnull(valor_celula):
        return None
    if isinstance(valor_celula, str):
        return re.sub(r'[^\d.,]', '', valor_celula)
    return valor_celula

def limpar_qtd_colunas(df):
    # Utilizar uma expressão regular para encontrar colunas que terminam com _QTD seguido ou não de _ e números
    colunas_qtd = [col for col in df.columns if re.search(r'_QTD(_\d+)?$', col)]
    for col in colunas_qtd:
        df[col] = df[col].apply(limpar_celula)
    return df

def extrair_dados_pdf(f, codigos_proventos, codigos_desconto, mapeamento_codigos=None):
    pdf = PdfReader(f)
    data = []

    # Chaves do dicionário patterns na função extrair_cadastro
    colunas_padrao = ['Codigo', 'Nome', 'Centro_c', 'Secretaria', 'Cargo', 
                      'CPF', 'ag', 'cc', 'Vinculo', 'Admissao', 'Nascimento']

    nomes_eventos = []  # Para armazenar nomes de eventos

    for page in pdf.pages:
        text = page.extract_text()
        sections = text.split('FUNCIONÁRIO:')
        for section in sections[1:]:
            data.append(extrair_cadastro('FUNCIONÁRIO:' + section, codigos_proventos, codigos_desconto))

    # Adicionar nomes de eventos
    for code, field_name in {**codigos_proventos, **codigos_desconto}.items():
        nomes_eventos.append(field_name + '_QTD')
        nomes_eventos.append(field_name + '_VALOR')

    df = pd.DataFrame(data)

    # Função para ordenar as colunas
    def ordenar_colunas(coluna):
        if coluna in colunas_padrao:
            return (0, colunas_padrao.index(coluna))  # Primeiro grupo, ordenado pela posição na lista colunas_padrao
        elif coluna in nomes_eventos:
            return (1, nomes_eventos.index(coluna))  # Segundo grupo, ordenado pela posição na lista nomes_eventos
        return (2, 0)  # Terceiro grupo, outras colunas

    # Ordenar colunas com base na função de ordenação
    colunas_ordenadas = sorted(df.columns, key=ordenar_colunas)
    df = df[colunas_ordenadas]

    df = limpar_qtd_colunas(df)
    if mapeamento_codigos:
        df['Codigo'] = df['Codigo'].map(mapeamento_codigos).fillna(df['Codigo'])

    return df

