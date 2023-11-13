import re
from PyPDF2 import PdfReader

def extrair_e_categorizar_dados(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        numero_de_paginas = len(reader.pages)

        # Lê as duas últimas páginas
        ultima_pagina = reader.pages[numero_de_paginas - 1].extract_text()
        penultima_pagina = reader.pages[numero_de_paginas - 2].extract_text()

        # Combina o texto das duas últimas páginas
        texto_combinado = penultima_pagina + ultima_pagina

        # Expressão regular atualizada para capturar também os valores monetários
        padrao = r'(\d{3})\s+(.+?)\s+\d+\s+([\d\.,]+)\s+([\d\.,]+)'
        resultados = re.findall(padrao, texto_combinado)

        proventos = []
        descontos = []

        for codigo, descricao, valor1, valor2 in resultados:
            if valor2 == '0,00':
                proventos.append((codigo, descricao.strip()))
            else:
                descontos.append((codigo, descricao.strip()))

        return proventos, descontos


