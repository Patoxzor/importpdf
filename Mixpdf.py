import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import pandas as pd
from data_func import extrair_dados_pdf
from excel_format import formatar_cabecalho, formatar_cpf, ajustar_largura_colunas, adicionar_bordas, remover_gridlines
from config_json import ConfigManager
from style_interface import configure_treeview_style, button_style, label_style
from update import check_for_updates
import os
import tkinter.colorchooser as colorchooser
from database_utils import get_sql_server_databases, verificar_todos_funcionarios


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrair dados de PDF ANDRE MARKAS - Bebeto Apps Inc. - 1.0.0")
        self.setup_ui()
        self.codigos_data = ConfigManager.load_from_file()
        current_directory = os.path.dirname(os.path.realpath(__file__))
        icon_path = os.path.join(current_directory, 'icone.ico')
        self.root.iconbitmap(icon_path)
        self.codigos_proventos = self.codigos_data.get("proventos", {})
        self.codigos_desconto = self.codigos_data.get("descontos", {})
        self.mapeamento_codigos = self.codigos_data.get("mapeamento", {})
        self.color_preferences = ConfigManager.load_color_preferences()
        self.root.configure(background=self.color_preferences.get('background'))
        self.frame_buttons.configure(background=self.color_preferences.get('background')),
        self.color_button_frame.configure(background=self.color_preferences.get('background'))
        self.label_filename.configure(bg=self.color_preferences.get('background'))
        check_for_updates()
    
    def setup_ui(self):
        self.color_button_frame = tk.Frame(self.root)
        self.color_button_frame.pack(anchor='ne', padx=10, pady=10)
        self.frame_buttons = tk.Frame(self.root)
        self.frame_buttons.pack(pady=20)
        self.create_database_combobox()
        self.create_verify_button()
        label_style_config = label_style()
        self.label_filename = tk.Label(self.root, text="Nenhum arquivo selecionado", **label_style_config)
        self.label_filename.pack(pady=10)
        self.excel_filename = None
        self.create_color_buttons()
        self.create_buttons()
        self.create_treeview()
        self.create_scrollbars()

    def create_buttons(self):
        style = button_style()
        self.button_open = tk.Button(self.frame_buttons, text="Selecionar PDF", command=self.open_pdf, **style)
        self.button_open.pack(side=tk.LEFT, padx=10)

        self.button_extract = tk.Button(self.frame_buttons, text="Extrair Dados", command=self.extract_data, state=tk.DISABLED, **style)
        self.button_extract.pack(side=tk.LEFT, padx=10)

        self.button_save = tk.Button(self.frame_buttons, text="Salvar em Excel", command=self.save_excel, state=tk.DISABLED, **style)
        self.button_save.pack(side=tk.LEFT, padx=10)

        self.button_open_excel = tk.Button(self.frame_buttons, text="Abrir Excel", command=self.open_excel, state=tk.DISABLED, **style)
        self.button_open_excel.pack(side=tk.LEFT, padx=10)

        self.button_view_codes = tk.Button(self.frame_buttons, text="Visualizar Eventos", command=self.view_codes, **style)
        self.button_view_codes.pack(side=tk.LEFT, padx=10)

        self.button_view_mapping = tk.Button(self.frame_buttons, text="Alteração de Códigos/Matricula", command=self.view_mapping, **style)
        self.button_view_mapping.pack(side=tk.LEFT, padx=10)

    def create_color_buttons(self):
        current_directory = os.path.dirname(os.path.realpath(__file__))
        
        palette_icon_path = os.path.join(current_directory, 'palette.png')

        palette_icon = tk.PhotoImage(file=palette_icon_path)

        self.button_change_bg_color = tk.Button(
            self.color_button_frame,
            image=palette_icon,
            command=lambda: self.choose_color('background'),
        )
        self.button_change_bg_color.pack(side=tk.LEFT, padx=2)
        self.button_change_bg_color.image = palette_icon
    
    def create_database_combobox(self):
        self.combobox_label = tk.Label(self.root, text="Selecione o banco de dados:", **label_style())
        self.combobox_label.pack(pady=5)

        self.database_combobox = ttk.Combobox(self.root, values=get_sql_server_databases(), state='readonly')
        self.database_combobox.pack(pady=5)

    def create_verify_button(self):
        self.button_verify = tk.Button(self.root, text="Verificar no SQL Server", command=self.verify_in_sql_server, state=tk.DISABLED, **button_style())
        self.button_verify.pack(pady=10)

    def verify_in_sql_server(self):
        selected_database = self.database_combobox.get()
        if selected_database:
            self.verificar_dados(selected_database)
        else:
            messagebox.showwarning("Aviso", "Por favor, selecione um banco de dados para verificar.")

    def verificar_dados(self, database):
        lista_funcionarios = self.df.to_dict('records')
        funcionarios_nao_encontrados = verificar_todos_funcionarios(database, lista_funcionarios)
        
        if funcionarios_nao_encontrados:
            self.mostrar_funcionarios_nao_encontrados(funcionarios_nao_encontrados)
        else:
            messagebox.showinfo("Verificação", "Todos os funcionários foram encontrados no banco de dados.")


    def mostrar_funcionarios_nao_encontrados(self, funcionarios_nao_encontrados):
        new_window = tk.Toplevel(self.root)
        new_window.title("Funcionários não encontrados")

        # Cria um Treeview na nova janela
        tree = ttk.Treeview(new_window, columns=list(funcionarios_nao_encontrados[0].keys()), show='headings')
        tree.pack(expand=True, fill='both', padx=10, pady=10)

        # Adiciona as colunas ao Treeview
        for col in funcionarios_nao_encontrados[0].keys():
            tree.heading(col, text=col)

        # Insere os dados no Treeview
        for funcionario in funcionarios_nao_encontrados:
            tree.insert('', 'end', values=list(funcionario.values()))

        # Use um DataFrame temporário para os dados não encontrados
        df_nao_encontrados = pd.DataFrame(funcionarios_nao_encontrados)

        button_save = tk.Button(new_window, text="Salvar em Excel", command=lambda: self.save_excel(df_nao_encontrados))
        button_save.pack(pady=10)


    def choose_color(self, target):
        color_code = colorchooser.askcolor(title="Escolha a cor", initialcolor=self.color_preferences.get(target))
        if color_code[1] is not None:
            self.color_preferences[target] = color_code[1]
            ConfigManager.save_color_preferences(self.color_preferences)
            
            if target == 'background':
                self.update_background_color(color_code[1])
            elif target == 'buttons':
                self.update_button_colors(color_code[1])


    def update_background_color(self, color):
        self.root.configure(background=color)
        self.frame_buttons.configure(background=color)
        self.color_button_frame.configure(background=color)
        self.label_filename.configure(bg=color)
        self.color_preferences['background'] = color
        ConfigManager.save_color_preferences(self.color_preferences)

    def create_treeview(self):
        self.canvas = tk.Canvas(self.root)
        self.canvas.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        self.tree_frame = ttk.Frame(self.canvas)
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.tree_frame, anchor='nw', 
                                                      tags="self.tree_frame")

        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.tree_frame.bind("<Configure>", self.on_frame_configure)

        # Treeview para exibir os dados
        self.tree = ttk.Treeview(self.tree_frame, columns=[str(i) for i in range(100)], show="headings")
        self.tree.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        configure_treeview_style()

    def create_scrollbars(self):
        self.scrollbar_vertical = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar_horizontal = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)

        self.scrollbar_vertical.pack(side=tk.RIGHT, fill=tk.Y)
        self.scrollbar_horizontal.pack(side=tk.RIGHT, fill=tk.X)
        self.tree.configure(yscrollcommand=self.scrollbar_vertical.set, xscrollcommand=self.scrollbar_horizontal.set)

        self.scrollbar_vertical.pack(side='right', fill='y')
        self.scrollbar_horizontal.pack(side='bottom', fill='x')

        self.tree.config(yscrollcommand=self.scrollbar_vertical.set)
        self.tree.config(xscrollcommand=self.scrollbar_horizontal.set)
        self.tree.pack(side="left", fill="both", expand=True)

    def on_canvas_configure(self, event):
        # Redimensionar a janela do frame dentro do canvas
        self.canvas.itemconfig("self.tree_frame", width=event.width, height=event.height)
        
        # Esta parte é importante para que o frame expanda horizontalmente
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def on_frame_configure(self, event):
        # Assegurar que a área de rolagem do canvas se ajuste ao tamanho do frame
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        

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
        popup = tk.Toplevel(self.root)
        popup.title("Visualizar Códigos Adicionados")
        popup.geometry("500x500")

        frame_proventos = tk.Frame(popup)
        frame_proventos.pack(pady=10, padx=10, fill=tk.X)

        label_proventos = tk.Label(frame_proventos, text="Proventos", font=("Arial", 12))
        label_proventos.pack()

        self.listbox_proventos = tk.Listbox(frame_proventos)
        self.update_listbox(self.listbox_proventos, 'provento')
        self.listbox_proventos.pack(fill=tk.X)

        btn_add_provento = tk.Button(frame_proventos, text="Adicionar", command=lambda: self.add_provento(self.listbox_proventos))
        btn_add_provento.pack(side=tk.LEFT, padx=5)

        btn_edit_provento = tk.Button(frame_proventos, text="Editar", command=lambda: self.edit_code(self.listbox_proventos, 'provento'))
        btn_edit_provento.pack(side=tk.LEFT, padx=5)

        btn_remove_provento = tk.Button(frame_proventos, text="Remover", command=lambda: self.remove_code(self.listbox_proventos, 'provento'))
        btn_remove_provento.pack(side=tk.LEFT, padx=5)

        frame_descontos = tk.Frame(popup)
        frame_descontos.pack(pady=10, padx=10, fill=tk.X)

        label_descontos = tk.Label(frame_descontos, text="Descontos", font=("Arial", 12))
        label_descontos.pack()

        self.listbox_descontos = tk.Listbox(frame_descontos)
        self.update_listbox(self.listbox_descontos, 'desconto')  # Atualiza a listbox com os dados atuais
        self.listbox_descontos.pack(fill=tk.X)

        btn_add_desconto = tk.Button(frame_descontos, text="Adicionar", command=lambda: self.add_desconto(self.listbox_descontos))
        btn_add_desconto.pack(side=tk.LEFT, padx=5)

        btn_edit_desconto = tk.Button(frame_descontos, text="Editar", command=lambda: self.edit_code(self.listbox_descontos, 'desconto'))
        btn_edit_desconto.pack(side=tk.LEFT, padx=5)

        btn_remove_desconto = tk.Button(frame_descontos, text="Remover", command=lambda: self.remove_code(self.listbox_descontos, 'desconto'))
        btn_remove_desconto.pack(side=tk.LEFT, padx=5)

        self.button_clear_all_proventos = tk.Button(frame_proventos, text="Remover Todos", command=self.clear_all_proventos)
        self.button_clear_all_proventos.pack(side=tk.LEFT, padx=5)

        self.button_clear_all_descontos = tk.Button(frame_descontos, text="Remover Todos", command=self.clear_all_descontos)
        self.button_clear_all_descontos.pack(side=tk.LEFT, padx=5)

        btn_close = tk.Button(popup, text="Fechar", command=popup.destroy)
        btn_close.pack(pady=20)
        popup.wait_window()


    def update_listbox(self, listbox, tipo):
        listbox.delete(0, tk.END)  # Limpa a listbox antes de adicionar novos itens

        # Escolhe o dicionário apropriado com base no tipo
        if tipo == 'provento':
            items = self.codigos_proventos.items()
        elif tipo == 'desconto':
            items = self.codigos_desconto.items()
        elif tipo == 'mapping':
            items = self.mapeamento_codigos.items()
        else:
            return

        for code, desc in items:
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
    
    def clear_all_proventos(self):
            self.codigos_proventos.clear()
            self.update_listbox(self.listbox_proventos, 'provento')
            self.save_codigos()

    def clear_all_descontos(self):
            self.codigos_desconto.clear()
            self.update_listbox(self.listbox_descontos, 'desconto')
            self.save_codigos()

    def save_to_file(self):
        ConfigManager.save_to_file(self.codigos_data)
        
    def save_codigos(self):
        self.codigos_data["proventos"] = self.codigos_proventos
        self.codigos_data["descontos"] = self.codigos_desconto
        self.codigos_data["mapeamento"] = self.mapeamento_codigos
        self.save_to_file()

    def view_mapping(self):
        popup = tk.Toplevel(self.root)
        popup.title("Visualizar Mapeamento de Códigos")
        popup.geometry("500x500")

        frame_mapping = tk.Frame(popup)
        frame_mapping.pack(pady=10, padx=10, fill=tk.X)

        label_mapping = tk.Label(frame_mapping, text="Mapeamento de Códigos", font=("Arial", 12))
        label_mapping.pack()

        self.listbox_mapping = tk.Listbox(frame_mapping)
        self.update_listbox(self.listbox_mapping, 'mapping')
        self.listbox_mapping.pack(fill=tk.X)

        btn_add_mapping = tk.Button(frame_mapping, text="Adicionar", command=self.add_mapping)
        btn_add_mapping.pack(side=tk.LEFT, padx=5)

        btn_edit_mapping = tk.Button(frame_mapping, text="Editar", command=self.edit_mapping)
        btn_edit_mapping.pack(side=tk.LEFT, padx=5)

        btn_remove_mapping = tk.Button(frame_mapping, text="Remover", command=self.remove_mapping)
        btn_remove_mapping.pack(side=tk.LEFT, padx=5)

        self.button_clear_all_mappings = tk.Button(frame_mapping, text="Remover Todos", command=self.clear_all_mappings)
        self.button_clear_all_mappings.pack(side=tk.LEFT, padx=5)

        btn_close = tk.Button(popup, text="Fechar", command=popup.destroy)
        btn_close.pack(pady=20)
        popup.wait_window()

    def add_mapping(self):
        new_code = simpledialog.askstring("Adicionar código/matricula", "Digite o código para ser substituido:")
        if new_code:
            # Aqui você pediria também para a descrição ou o mapeamento correspondente ao novo código
            new_mapping = simpledialog.askstring("Adicionar código/matricula", "Digite o código novo:")
            if new_mapping:
                self.mapeamento_codigos[new_code] = new_mapping
                self.update_listbox(self.listbox_mapping, 'mapping')
                self.save_codigos() 
    
    def edit_mapping(self):
        selected = self.listbox_mapping.curselection()
        if not selected:
            messagebox.showinfo("Editar código/matricula", "Selecione um código/matricula para editar.")
            return

        index = selected[0]
        old_code = self.listbox_mapping.get(index).split(" - ")[0]

        new_code = simpledialog.askstring("Editar código/matricula", f"Digite o novo código/matricula para '{old_code}':")
        if new_code:
            new_mapping = simpledialog.askstring("Editar código/matricula", f"Digite o novo código/matricula para o código '{new_code}':")
            if new_mapping:
                # Aqui você atualizaria o dicionário de mapeamentos
                del self.mapeamento_codigos[old_code]
                self.mapeamento_codigos[new_code] = new_mapping
                self.update_listbox(self.listbox_mapping, 'mapping')
                self.save_codigos() 

    def remove_mapping(self):
        selected = self.listbox_mapping.curselection()
        if not selected:
            messagebox.showinfo("Remover código/matricula", "Selecione um código/matricula para remover.")
            return

        index = selected[0]
        code = self.listbox_mapping.get(index).split(" - ")[0]

        if messagebox.askyesno("Remover código/matricula", f"Tem certeza que deseja remover o código/matricula para '{code}'?"):
            del self.mapeamento_codigos[code]
            self.listbox_mapping.delete(index)
            self.save_codigos() 

    def clear_all_mappings(self):
            self.mapeamento_codigos.clear()
            self.update_listbox(self.listbox_mapping, 'mapping')
            self.save_codigos()

    def open_pdf(self):
        self.filename = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
        if self.filename:
            self.button_extract.config(state=tk.NORMAL)
            self.button_save.config(state=tk.DISABLED)
            self.button_open_excel.config(state=tk.DISABLED)
            base_name = os.path.basename(self.filename)
            self.label_filename.config(text=base_name)

    def extract_data(self):
        try:
            self.df = extrair_dados_pdf(self.filename, self.codigos_proventos, self.codigos_desconto, self.mapeamento_codigos)
            #Formatar CPF
            if 'CPF' in self.df.columns:
                self.df['CPF'] = self.df['CPF'].apply(formatar_cpf)
            self.populate_tree()
            self.button_save.config(state=tk.NORMAL)
            self.button_verify.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def populate_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Atualizar as colunas do Treeview com base nos dados
        self.tree["columns"] = list(self.df.columns)

        # Definindo zebra stripping
        even_color = 'lightblue'
        odd_color = 'white'

        for idx, col in enumerate(self.df.columns):
            self.tree.heading(col, text=col)
            max_width = max(self.df[col].astype(str).apply(len).max(), len(col))
            self.tree.column(col, width=max_width * 10)

        for i, row in self.df.iterrows():
            self.tree.insert("", "end", values=list(row), 
                             tags=('evenrow' if i % 2 == 0 else 'oddrow'))

        self.tree.tag_configure('evenrow', background=even_color)
        self.tree.tag_configure('oddrow', background=odd_color)

    def save_excel(self, df=None):
        if df is None:
            df = self.df

        excel_filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos XLSX", "*.xlsx")])
        if excel_filename:
            with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Planilha')
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
