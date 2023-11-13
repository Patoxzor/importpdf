import re
import pdfplumber

def extrair_e_categorizar_dados(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        numero_de_paginas = len(pdf.pages)

        proventos = []
        descontos = []

        for i in [-2, -1]:  # Processar as duas últimas páginas
            pagina = pdf.pages[numero_de_paginas + i].extract_text()
            if pagina:
                padrao = r'^(\d+)\s+(.+?)\s+(\d+)\s+([\d\.,]+)\s+([\d\.,]+)$'
                resultados = re.findall(padrao, pagina, re.MULTILINE)

                for codigo, descricao, quantidade, valor1, valor2 in resultados:
                    if valor2 == '0,00':
                        proventos.append((codigo, descricao.strip(), quantidade, valor1))
                    else:
                        descontos.append((codigo, descricao.strip(), quantidade, valor2))

        return proventos, descontos
