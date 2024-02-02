import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import re
import tkinter.messagebox as mb

class ExcelToCsvConverter:
    def __init__(self, coluna_codigo, sufixos):
        self.coluna_codigo = coluna_codigo
        self.sufixos = sufixos

    def selecionar_arquivos_excel(self):
        root = tk.Toplevel()
        root.withdraw()
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", ".xlsx;.xls")])
        root.destroy()
        return list(file_paths)

    def ajustar_valores(self,df, coluna):
        # Primeiro, assegura-se que os dados são tratados como strings
        df[coluna] = df[coluna].astype(str)
        
        # Remove pontos que funcionam como separadores de milhares em valores decimais
        df[coluna] = df[coluna].apply(lambda x: x.replace('.', '') if ',' in x else x)
        
        # Substitui ponto por vírgula em valores decimais
        df[coluna] = df[coluna].apply(lambda x: x.replace('.', ',') if '.' in x and ',' not in x else x)

        # Para os casos onde não há vírgulas (indicando milhares), substitui espaços por vírgulas
        df[coluna] = df[coluna].apply(lambda x: re.sub(r'(\d)(\d{3})$', r'\1,\2', x))
        
        return df


    def gerar_csvs_por_grupo(self, file_path, coluna_codigo):
        df = pd.read_excel(file_path)
        diretorio = os.path.dirname(file_path)
        regex_eventos = re.compile(r'(.+?)(_VALOR|_QTD)(_?\d+)?$')
        eventos = set(regex_eventos.match(col).groups()[0] for col in df.columns if regex_eventos.match(col))

        for evento in eventos:
            colunas_evento = [col for col in df.columns if regex_eventos.match(col) and regex_eventos.match(col).groups()[0] == evento]

            # Agrupar colunas por sufixo numérico
            grupos_sufixo = {}
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
                df_sufixo = df[[self.coluna_codigo] + cols_sufixo].dropna(how='all', subset=cols_sufixo)
                for col in cols_sufixo:
                    df_sufixo = self.ajustar_valores(df_sufixo, col)
                    novo_nome_col = re.sub(r'_\d+$', '', col)
                    df_sufixo.rename(columns={col: novo_nome_col}, inplace=True)

                # Reordenar as colunas
                cols_valor = [col for col in df_sufixo.columns if col.endswith('_VALOR')]
                cols_qtd = [col for col in df_sufixo.columns if col.endswith('_QTD')]
                colunas_ordenadas = [self.coluna_codigo] + cols_valor + cols_qtd
                df_sufixo = df_sufixo[colunas_ordenadas]

                todas_linhas.append(df_sufixo)

            df_filtrado = pd.concat(todas_linhas, ignore_index=True)

            # Verifique se todas as colunas necessárias existem
            colunas_necessarias = [coluna_codigo] + list(colunas_evento_sem_sufixo)
            colunas_existentes = [col for col in colunas_necessarias if col in df_filtrado.columns]

            if len(colunas_existentes) != len(colunas_necessarias):
                print(f"Aviso: nem todas as colunas necessárias foram encontradas para o evento '{evento}'.")
                print("Colunas faltantes:", set(colunas_necessarias) - set(colunas_existentes))

            caminho_csv = os.path.join(diretorio, f"{evento}.csv")
            os.makedirs(os.path.dirname(caminho_csv), exist_ok=True)

            if os.path.exists(caminho_csv):
                df_existente = pd.read_csv(caminho_csv, sep=';')
                df_filtrado = pd.concat([df_existente, df_filtrado], ignore_index=True)

            df_filtrado.to_csv(caminho_csv, index=False, sep=';', decimal=',', header=True, encoding='utf-8-sig', quotechar='|')
            print(f'Arquivo CSV atualizado: {caminho_csv}')

    def processar_arquivos(self):
        print("Selecione os arquivos Excel.")
        arquivos_excel = self.selecionar_arquivos_excel()
        
        for arquivo in arquivos_excel:
            self.gerar_csvs_por_grupo(arquivo, self.coluna_codigo)
        
        mb.showinfo("Arquivos CSV", "Os arquivos csv foram gerados e estão na mesma pasta dos arquivos excel selecionados.")