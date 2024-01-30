import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import re

def selecionar_arquivos_excel():
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
    return list(file_paths)

def ajustar_valores(df, coluna_valor, coluna_qtd):
    df[coluna_valor] = df[coluna_valor].apply(lambda x: str(x).replace('.', '') if isinstance(x, str) else x)
    
    # Assegurando que todos os valores sejam tratados como strings
    df[coluna_qtd] = df[coluna_qtd].astype(str)
    
    # Substituindo ponto por vírgula e removendo caracteres não numéricos, exceto vírgula
    df[coluna_qtd] = df[coluna_qtd].apply(lambda x: re.sub(r'[^0-9,]', '', x.replace('.', ',')))

    return df

def gerar_csvs_por_grupo(file_path, coluna_codigo, sufixos):
    df = pd.read_excel(file_path)
    diretorio = os.path.dirname(file_path)
    eventos = set(col.split('_')[0] for col in df.columns if col.endswith(sufixos))

    for evento in eventos:
        coluna_valor = f"{evento}_VALOR"
        coluna_qtd = f"{evento}_QTD"

        if coluna_valor in df.columns and coluna_qtd in df.columns:
            df_filtrado = df[df[[coluna_valor, coluna_qtd]].notnull().any(axis=1)].copy()
            df_filtrado = ajustar_valores(df_filtrado, coluna_valor, coluna_qtd)

            colunas_para_exportar = [coluna_codigo, coluna_valor, coluna_qtd]
            df_export = df_filtrado[colunas_para_exportar].copy()

            caminho_csv = os.path.join(diretorio, f"{evento}.csv")
            os.makedirs(os.path.dirname(caminho_csv), exist_ok=True)

            if os.path.exists(caminho_csv):
                df_existente = pd.read_csv(caminho_csv, sep=';')
                df_export = pd.concat([df_existente, df_export])

            df_export.to_csv(caminho_csv, index=False, sep=';', decimal=',', header=True, encoding='utf-8-sig', quotechar='|')
            print(f'Arquivo CSV atualizado: {caminho_csv}')

def main():
    print("Selecione os arquivos Excel.")
    arquivos_excel = selecionar_arquivos_excel()
    coluna_codigo = 'Codigo'
    sufixos = ('_QTD', '_VALOR')
    
    for arquivo in arquivos_excel:
        gerar_csvs_por_grupo(arquivo, coluna_codigo, sufixos)

if __name__ == "__main__":
    main()
