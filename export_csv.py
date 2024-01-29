import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import re

def selecionar_arquivos_excel():
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", ".xlsx;.xls")])
    return list(file_paths)

def ajustar_valores(df, coluna):
    # Primeiro, assegura-se que os dados são tratados como strings
    df[coluna] = df[coluna].astype(str)
    
    # Substitui pontos por vírgulas para estar de acordo com o formato brasileiro de números decimais
    df[coluna] = df[coluna].apply(lambda x: x.replace('.', ',') if ',' not in x else x)
    
    # Para os casos onde não há vírgulas (indicando milhares), substitui espaços por vírgulas
    df[coluna] = df[coluna].apply(lambda x: re.sub(r'(\d)(\d{3})$', r'\1,\2', x))
    
    return df

def gerar_csvs_por_grupo(file_path, coluna_codigo, sufixos):
    df = pd.read_excel(file_path)
    diretorio = os.path.dirname(file_path)
    regex_eventos = re.compile(r'(.+?)(_VALOR|_QTD)(_?\d+)?')
    eventos = set(regex_eventos.match(col).groups()[0] for col in df.columns if regex_eventos.match(col))

    for evento in eventos:
        colunas_evento = [col for col in df.columns if col.startswith(evento) and 
                          (col.endswith('_QTD') or col.endswith('_VALOR') or re.search(r'(_VALOR|_QTD)_\d+', col))]

        # Agrupar colunas por sufixo numérico
        grupos_sufixo = {}  # Definindo grupos_sufixo aqui
        for col in colunas_evento:
            match = regex_eventos.match(col)
            if match:
                sufixo = match.groups()[2] or ''  # '_1', '_2', etc., ou string vazia se não houver sufixo
                if sufixo not in grupos_sufixo:
                    grupos_sufixo[sufixo] = []
                grupos_sufixo[sufixo].append(col)

        # Remover sufixos numéricos dos nomes das colunas para a verificação de duplicatas
        colunas_evento_sem_sufixo = set(re.sub(r'_\d+$', '', col) for col in colunas_evento)

        # Processar cada grupo de sufixo
        todas_linhas = []
        for sufixo, cols_sufixo in grupos_sufixo.items():
            df_sufixo = df[[coluna_codigo] + cols_sufixo].dropna(how='all', subset=cols_sufixo)
            for col in cols_sufixo:
                df_sufixo = ajustar_valores(df_sufixo, col)
                novo_nome_col = re.sub(r'_\d+$', '', col)
                df_sufixo.rename(columns={col: novo_nome_col}, inplace=True)
            todas_linhas.append(df_sufixo)

        df_filtrado = pd.concat(todas_linhas, ignore_index=True)
        # Reordenar colunas
        colunas_ordenadas = [coluna_codigo] + [evento + sufixo for sufixo in sufixos]
        df_filtrado = df_filtrado[colunas_ordenadas]

        caminho_csv = os.path.join(diretorio, f"{evento}.csv")
        os.makedirs(os.path.dirname(caminho_csv), exist_ok=True)

        if os.path.exists(caminho_csv):
            df_existente = pd.read_csv(caminho_csv, sep=';')
            df_filtrado = pd.concat([df_existente, df_filtrado], ignore_index=True)

        # Usar nomes de colunas sem sufixos numéricos para drop_duplicates
        df_filtrado = df_filtrado.drop_duplicates(subset=[coluna_codigo] + list(colunas_evento_sem_sufixo))
        df_filtrado.to_csv(caminho_csv, index=False, sep=';', decimal=',', header=True, encoding='utf-8-sig', quotechar='|')
        print(f'Arquivo CSV atualizado: {caminho_csv}')

def main():
    print("Selecione os arquivos Excel.")
    arquivos_excel = selecionar_arquivos_excel()
    coluna_codigo = 'Codigo'
    sufixos = ('_VALOR', '_QTD')
    
    for arquivo in arquivos_excel:
        gerar_csvs_por_grupo(arquivo, coluna_codigo, sufixos)

if __name__ == "__main__":
    main()