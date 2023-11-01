import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import pandas as pd
from Extrair_dados import extrair_dados_pdf
from excel_format import formatar_cabecalho, formatar_cpf, ajustar_largura_colunas, adicionar_bordas, remover_gridlines
import json
import os

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrair dados de PDF ANDRE MARKAS - Bebeto Apps Inc.")

        self.frame_buttons = tk.Frame(root)
        self.frame_buttons.pack(pady=20)

        self.label_filename = tk.Label(self.root, text="Nenhum arquivo selecionado", font=("Arial", 10))
        self.label_filename.pack(pady=10)

        self.excel_filename = None

        # Botões
        self.button_open = tk.Button(self.frame_buttons, text="Selecionar PDF", command=self.open_pdf)
        self.button_open.pack(side=tk.LEFT, padx=10)

        self.button_extract = tk.Button(self.frame_buttons, text="Extrair Dados", command=self.extract_data, state=tk.DISABLED)
        self.button_extract.pack(side=tk.LEFT, padx=10)

        self.button_save = tk.Button(self.frame_buttons, text="Salvar em Excel", command=self.save_excel, state=tk.DISABLED)
        self.button_save.pack(side=tk.LEFT, padx=10)

        self.button_open_excel = tk.Button(self.frame_buttons, text="Abrir Excel", command=self.open_excel, state=tk.DISABLED)
        self.button_open_excel.pack(side=tk.LEFT, padx=10)

        # Canvas e Frame para o Treeview
        self.canvas = tk.Canvas(root)
        self.canvas.pack(expand=True, fill=tk.BOTH)
        self.tree_frame = ttk.Frame(self.canvas)
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.tree_frame, anchor='nw')

        # Treeview para exibir os dados
        self.tree = ttk.Treeview(self.tree_frame, columns=[str(i) for i in range(100)], show="headings")
        self.tree.pack(expand=True, fill=tk.BOTH)

        # Scrollbars
        self.scrollbar_vertical = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.scrollbar_horizontal = ttk.Scrollbar(root, orient="horizontal", command=self.canvas.xview)
        self.canvas.config(yscrollcommand=self.scrollbar_vertical.set, xscrollcommand=self.scrollbar_horizontal.set)

        self.scrollbar_vertical.pack(side=tk.RIGHT, fill=tk.Y)
        self.scrollbar_horizontal.pack(side=tk.BOTTOM, fill=tk.X)

        # Função para atualizar a área de rolagem do canvas
        self.tree_frame.bind("<Configure>", self.update_scrollregion) 

        self.button_view_codes = tk.Button(self.frame_buttons, text="Visualizar Códigos", command=self.view_codes)
        self.button_view_codes.pack(side=tk.LEFT, padx=10)

        # Listas de códigos
        self.codigos_proventos = {}
        self.codigos_desconto = {}

        self.codigos_data = self.codigos_data = App.load_from_file()
        self.codigos_proventos = self.codigos_data.get("proventos", {})
        self.codigos_desconto = self.codigos_data.get("descontos", {})


    def add_provento(self, listbox_proventos):
        code = simpledialog.askstring("Adicionar Provento", "Digite o código do provento:")
        description = simpledialog.askstring("Adicionar Provento", "Digite a descrição do provento:")
        if code and description:
            self.codigos_proventos[code] = description
            self.button_save.config(state=tk.DISABLED)
            self.button_extract.config(state=tk.NORMAL)
            self.update_listbox(listbox_proventos, 'provento')
            self.save_codigos()

    def add_desconto(self, listbox_descontos):
        code = simpledialog.askstring("Adicionar Desconto", "Digite o código do desconto:")
        description = simpledialog.askstring("Adicionar Desconto", "Digite a descrição do desconto:")
        if code and description:
            self.codigos_desconto[code] = description
            self.button_save.config(state=tk.DISABLED)
            self.button_extract.config(state=tk.NORMAL)
            self.update_listbox(listbox_descontos, 'desconto')
            self.save_codigos()

    def view_codes(self):
        # Criação de uma nova janela (popup)
        popup = tk.Toplevel(self.root)
        popup.title("Visualizar Códigos Adicionados")
        popup.geometry("500x500")

        # Frame para proventos
        frame_proventos = tk.Frame(popup)
        frame_proventos.pack(pady=10, padx=10, fill=tk.X)

        label_proventos = tk.Label(frame_proventos, text="Proventos", font=("Arial", 12))
        label_proventos.pack()

        listbox_proventos = tk.Listbox(frame_proventos)
        for code, desc in self.codigos_proventos.items():
            listbox_proventos.insert(tk.END, f"{code} - {desc}")
        listbox_proventos.pack(fill=tk.X)

        btn_add_provento = tk.Button(frame_proventos, text="Adicionar", command=lambda: self.add_provento(listbox_proventos))
        btn_add_provento.pack(side=tk.LEFT, padx=5)

        btn_edit_provento = tk.Button(frame_proventos, text="Editar", command=lambda: self.edit_code(listbox_proventos, 'provento'))
        btn_edit_provento.pack(side=tk.LEFT, padx=5)

        btn_remove_provento = tk.Button(frame_proventos, text="Remover", command=lambda: self.remove_code(listbox_proventos, 'provento'))
        btn_remove_provento.pack(side=tk.LEFT, padx=5)

        frame_descontos = tk.Frame(popup)
        frame_descontos.pack(pady=10, padx=10, fill=tk.X)

        label_descontos = tk.Label(frame_descontos, text="Descontos", font=("Arial", 12))
        label_descontos.pack()

        listbox_descontos = tk.Listbox(frame_descontos)
        for code, desc in self.codigos_desconto.items():
            listbox_descontos.insert(tk.END, f"{code} - {desc}")
        listbox_descontos.pack(fill=tk.X)

        btn_add_desconto = tk.Button(frame_descontos, text="Adicionar", command=lambda: self.add_desconto(listbox_descontos))
        btn_add_desconto.pack(side=tk.LEFT, padx=5)

        btn_edit_desconto = tk.Button(frame_descontos, text="Editar", command=lambda: self.edit_code(listbox_descontos, 'desconto'))
        btn_edit_desconto.pack(side=tk.LEFT, padx=5)

        btn_remove_desconto = tk.Button(frame_descontos, text="Remover", command=lambda: self.remove_code(listbox_descontos, 'desconto'))
        btn_remove_desconto.pack(side=tk.LEFT, padx=5)

        btn_close = tk.Button(popup, text="Fechar", command=popup.destroy)
        btn_close.pack(pady=20)

    def update_listbox(self, listbox, tipo):
        listbox.delete(0, tk.END)

        if tipo == 'provento':
            for code, desc in self.codigos_proventos.items():
                listbox.insert(tk.END, f"{code} - {desc}")
        elif tipo == 'desconto':
            for code, desc in self.codigos_desconto.items():
                listbox.insert(tk.END, f"{code} - {desc}")
  
    def remove_code(self, listbox, tipo):
        selected = listbox.curselection()

        if not selected:
            return

        index = selected[0]
        item_text = listbox.get(index) 
        code = item_text.split(" - ")[0]

        if tipo == 'provento':
            if code in self.codigos_proventos:
                del self.codigos_proventos[code]
        elif tipo == 'desconto':
            if code in self.codigos_desconto:
                del self.codigos_desconto[code]

        listbox.delete(index)
        self.save_codigos()

    def edit_code(self, listbox, tipo):
        selected = listbox.curselection()

        if not selected:
            return

        index = selected[0]
        item_text = listbox.get(index)
        code = item_text.split(" - ")[0]

        new_code = simpledialog.askstring("Editar", f"Digite o novo código para '{code}':")
        new_description = simpledialog.askstring("Editar", f"Digite a nova descrição para o código '{new_code}':")

        if not new_code or not new_description:
            return

        # Atualiza os dicionários.
        if tipo == 'provento':
            if code in self.codigos_proventos:
                del self.codigos_proventos[code] 
            self.codigos_proventos[new_code] = new_description
        elif tipo == 'desconto':
            if code in self.codigos_desconto:
                del self.codigos_desconto[code]
            self.codigos_desconto[new_code] = new_description

        listbox.delete(index)
        listbox.insert(index, f"{new_code} - {new_description}")
        self.save_codigos()

    def update_scrollregion(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def save_to_file(data, filename="codigos.json"):
        with open(filename, 'w') as file:
            json.dump(data, file)

    def load_from_file(filename="codigos.json"):
        try:
            with open(filename, 'r') as file:
                data = json.load(file)
            return data
        except FileNotFoundError:
            return {}
        
    def save_codigos(self):
        self.codigos_data["proventos"] = self.codigos_proventos
        self.codigos_data["descontos"] = self.codigos_desconto
        App.save_to_file(self.codigos_data)

    def open_pdf(self):
        self.filename = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
        if self.filename:
            self.button_extract.config(state=tk.NORMAL)
            base_name = os.path.basename(self.filename)
            self.label_filename.config(text=base_name)

    def extract_data(self):
        try:
            self.df = extrair_dados_pdf(self.filename, self.codigos_proventos, self.codigos_desconto)
            #Formatar CPF
            if 'CPF' in self.df.columns:
                self.df['CPF'] = self.df['CPF'].apply(formatar_cpf)
            self.populate_tree()
            self.button_save.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def populate_tree(self):
        # Limpar o Treeview antes de inserir novos dados
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Atualizar as colunas do Treeview com base nos dados
        self.tree["columns"] = list(self.df.columns)

        for idx, col in enumerate(self.df.columns):
            self.tree.heading(col, text=col)

            max_width = max(self.df[col].astype(str).apply(len).max(), len(col))
            self.tree.column(col, width=max_width * 10)

        for _, row in self.df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def save_excel(self):
        excel_filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos XLSX", "*.xlsx")])
        if excel_filename:
            with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
                self.df.to_excel(writer, index=False, sheet_name='Planilha')               
                planilha = writer.sheets['Planilha']
                
                formatar_cabecalho(planilha)
                ajustar_largura_colunas(planilha)
                adicionar_bordas(planilha)
                remover_gridlines(planilha)
                
            self.excel_filename = excel_filename
            self.button_open_excel.config(state=tk.NORMAL)
            messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")

    def open_excel(self):
        if self.excel_filename:
            os.system(f'start excel "{self.excel_filename}"')

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
