import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import pandas as pd
from decimal import Decimal, InvalidOperation
import datetime
from data_func import extrair_dados_pdf
from excel_format import formatar_cabecalho, formatar_cpf, ajustar_largura_colunas, adicionar_bordas, remover_gridlines
from config_json import ConfigManager
from style_interface import configure_treeview_style, button_style, label_style
from update import check_for_updates, download_and_install_update
import os
from log_erros import setup_logging, logger
import tkinter.colorchooser as colorchooser
from database_utils import get_sql_server_databases, verificar_todos_funcionarios, verificar_existencia, criar_conexao_sql_server, criar_registro, buscar_empresa_por_descricao, verificar_cargo_e_salario, adicionar_funcao, verificar_codigo_funcao, inserir_funcionarios_no_banco, buscar_codigos_por_cpfs
from eventos import extrair_e_categorizar_dados

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrair dados de PDF ANDRE MARKAS - Bebeto Apps Inc. - 1.0.1")
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
        self.combobox_cpfs = ttk.Combobox(self.root, values=[], state='readonly')
        self.alteracoes_pendentes = []
        self.label_proventos = None
        self.label_descontos = None
        self.database = None
        self.filename = None
        self.descricoes_atualizadas = False
        self.cpf_selecionado = None
        self.popup_windows = {}
        check_for_updates()
    
    def setup_ui(self):
        self.check_and_notify_updates()
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

    def check_and_notify_updates(self):
        current_version = ConfigManager.load_version()
        latest_version, download_url = check_for_updates()  

        if latest_version:
            self.notify_user_of_update(latest_version, download_url)

    def notify_user_of_update(self, latest_version, download_url):
        response = messagebox.askyesno("Atualização Disponível", 
                                    f"Uma nova versão {latest_version} está disponível. Deseja atualizar agora?",
                                    parent=self.root)
        if response:
            download_and_install_update(download_url, latest_version)

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

        self.button_verificar_codigos = tk.Button(self.frame_buttons, text="Verificar Codigos existentes", command=lambda:self.codigos_divergentes(self.tree), **style)
        self.button_verificar_codigos.pack(side=tk.LEFT, padx=10)

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
            self.database = criar_conexao_sql_server(selected_database)
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
        if 'mostrar_funcionarios_nao_encontrados' in self.popup_windows:
            self.popup_windows['mostrar_funcionarios_nao_encontrados'].destroy()
            del self.popup_windows['mostrar_funcionarios_nao_encontrados']
        new_window = tk.Toplevel(self.root)
        new_window.title("Funcionários não encontrados")
        self.popup_windows['mostrar_funcionarios_nao_encontrados'] = new_window
        
        container = tk.Frame(new_window)
        container.pack(fill='both', expand=True)

        vsb = ttk.Scrollbar(container, orient="vertical")
        vsb.pack(side='right', fill='y')

        hsb = ttk.Scrollbar(container, orient="horizontal")
        hsb.pack(side='bottom', fill='x')

        tree = ttk.Treeview(container, columns=list(funcionarios_nao_encontrados[0].keys()), show='headings',
                            yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(side='top', fill='both', expand=True)

        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)

        for col in funcionarios_nao_encontrados[0].keys():
            tree.heading(col, text=col)
            tree.column(col, anchor='w') 

        for funcionario in funcionarios_nao_encontrados:
            tree.insert('', 'end', values=list(funcionario.values()))

        for col in tree['columns']:
            tree.column(col, width=100, minwidth=50)

        df_nao_encontrados = pd.DataFrame(funcionarios_nao_encontrados)

        button_save = tk.Button(new_window, text="Salvar em Excel", command=lambda: self.save_excel(tree))
        button_save.pack(side=tk.LEFT, padx=10)

        button_check_centro_custos = tk.Button(new_window, text="Verificar Dados", command=lambda: self.verificar_centros_custos_e_secretarias_treeview(tree, 'centrocusto', 'localizacao'))
        button_check_centro_custos.pack(side=tk.LEFT, padx=10)

        button_update = tk.Button(new_window, text="Atualizar Descrições", command=lambda: self.atualizar_dados_e_vinculo(tree))
        button_update.pack(side=tk.LEFT, padx=10)

        tree.bind("<Double-1>", lambda event: self.on_item_double_click(event, tree))

        button_add_empresa = tk.Button(new_window, text="Adicionar Códigos de Empresa", command=lambda: self.adicionar_coluna_codigo_empresa(tree, self.database, 'empresas'))
        button_add_empresa.pack(side=tk.LEFT, padx=10)
        
        button_cadastrar_todos = tk.Button(new_window, text="Cadastrar Todos", command=lambda: self.cadastrar_todos_funcionarios(tree))
        button_cadastrar_todos.pack(pady=10)

        new_window.minsize(600, 400)

        def search_treeview(search_query, start_node=None, reverse=False):
            children = tree.get_children()
            start_index = 0
            if start_node:
                start_index = children.index(start_node) - 1 if reverse else children.index(start_node) + 1
            else:
                start_index = len(children) - 1 if reverse else 0

            search_range = range(start_index, -1, -1) if reverse else range(start_index, len(children))

            for i in search_range:
                if search_query.lower() in str(tree.item(children[i], 'values')).lower():
                    tree.see(children[i])
                    tree.selection_set(children[i])
                    return children[i]

            return None 

        def open_search_box():
            search_window = tk.Toplevel(new_window)
            search_window.title("Pesquisar")
            search_window.grab_set() 
            search_window.focus_set() 

            search_label = tk.Label(search_window, text="Procurar:")
            search_label.pack(side=tk.LEFT)
            search_entry = tk.Entry(search_window)
            search_entry.pack(side=tk.LEFT, padx=5)
            search_entry.focus_set() 

            def find_next():
                selected = tree.selection()
                search_treeview(search_entry.get(), selected[0] if selected else None)

            def find_previous():
                selected = tree.selection()
                search_treeview(search_entry.get(), selected[0] if selected else None, reverse=True)

            def on_search_entry_event(event):
                if event.keysym == 'Return' and event.state & (1 << 0):  # Verifica se o Shift está pressionado
                    find_previous()
                elif event.keysym == 'Return':
                    find_next()
                elif event.keysym == 'Escape':
                    search_window.destroy()

            search_entry.bind('<Return>', on_search_entry_event)
            search_entry.bind('<Escape>', on_search_entry_event)

            search_button_next = tk.Button(search_window, text="Próximo", command=find_next)
            search_button_next.pack(side=tk.LEFT)

            search_button_previous = tk.Button(search_window, text="Anterior", command=find_previous)
            search_button_previous.pack(side=tk.LEFT)

        new_window.bind('<Control-f>', lambda event: open_search_box())
    
    def codigos_divergentes(self, tree):
        self.resultados = {}
        if self.database is None:
            messagebox.showerror("Erro", "A conexão com o banco de dados não foi estabelecida.")
            return

        for item in tree.get_children():
            valores = tree.item(item, 'values')
            cpf = valores[5] 
            codigo = valores[0] 
            self.resultados[cpf] = codigo

        if self.resultados:
            self.exibir_resultados_codigos_divergentes(self.resultados)
        else:
            messagebox.showinfo("Informação", "Não foram encontrados códigos divergentes.")

    def exibir_resultados_codigos_divergentes(self, resultados):
        codigos_db = buscar_codigos_por_cpfs(self.database, list(self.resultados.keys()))
        print("Códigos DB:", codigos_db)
        print("Resultados Treeview:", resultados)

        cpfs_para_exibir = []
        for cpf, codigo_treeview in self.resultados.items():
            codigos_db_lista = codigos_db.get(cpf)
            if codigos_db_lista is None or int(codigo_treeview) not in codigos_db_lista:
                cpfs_para_exibir.append(cpf)
                print(f"CPF divergente: {cpf}, Código Treeview: {codigo_treeview}, Código DB: {codigos_db_lista}") 

        self.janela_resultados = tk.Toplevel(self.root)
        self.janela_resultados.title("Resultados dos Códigos Divergentes")
        self.janela_resultados.geometry("700x400")

        frame_selecao_cpf = tk.Frame(self.janela_resultados)
        frame_selecao_cpf.pack(fill='x', padx=10, pady=5)
        label_selecao_cpf = tk.Label(frame_selecao_cpf, text="Selecione um CPF: ")
        combobox_cpfs = ttk.Combobox(frame_selecao_cpf, values=cpfs_para_exibir, state='readonly')
        label_selecao_cpf.pack(side='left', padx=5)
        combobox_cpfs.pack(side='left', padx=5)
        combobox_cpfs['values'] = cpfs_para_exibir

        frame_atual = tk.Frame(self.janela_resultados)
        frame_atual.pack(fill='x', padx=10, pady=5)
        label_cpf_atual = tk.Label(frame_atual, text="CPF: ")
        label_codigo_atual = tk.Label(frame_atual, text="Código Atual: ")
        label_cpf_atual.pack(side='left', padx=5)
        label_codigo_atual.pack(side='left', padx=5)

        frame_codigos_db = tk.Frame(self.janela_resultados)
        frame_codigos_db.pack(side='left', fill='both', expand=True, padx=10, pady=5)
        label_codigos_db = tk.Label(frame_codigos_db, text="Códigos Disponíveis: ")
        label_codigos_db.pack(side='top', padx=5)
        self.listbox_codigos_db = tk.Listbox(frame_codigos_db)
        self.listbox_codigos_db.pack(fill='both', expand=True)

        frame_alteracoes_pendentes = tk.Frame(self.janela_resultados)
        frame_alteracoes_pendentes.pack(side='right', fill='both', expand=True, padx=10, pady=5)
        label_alteracoes_pendentes = tk.Label(frame_alteracoes_pendentes, text="Alterações Pendentes: ")
        label_alteracoes_pendentes.pack(side='top', padx=5)
        self.listbox_alteracoes_pendentes = tk.Listbox(frame_alteracoes_pendentes)
        self.listbox_alteracoes_pendentes.pack(fill='both', expand=True)

        botao_adicionar_lista = tk.Button(frame_codigos_db, text="Adicionar à Lista", command=self.adicionar_a_lista)
        botao_adicionar_lista.pack(side='bottom', padx=5, pady=5)

        botao_importar_para_mapping = tk.Button(frame_alteracoes_pendentes, text="Importar para Alterações de códigos", command=self.importar_para_view_mapping)
        botao_importar_para_mapping.pack(side='bottom', padx=5, pady=5)

        def on_cpf_selected(event):
            cpf_selecionado = combobox_cpfs.get()
            codigo_atual_treeview = resultados.get(cpf_selecionado, "Não encontrado na Treeview")
            codigos_db_lista = codigos_db.get(cpf_selecionado, [])

            label_cpf_atual["text"] = f"CPF: {cpf_selecionado}"
            label_codigo_atual["text"] = f"Código Atual: {codigo_atual_treeview}"
            self.listbox_codigos_db.delete(0, tk.END)
            for codigo in codigos_db_lista:
                self.listbox_codigos_db.insert('end', codigo)
            self.cpf_selecionado = cpf_selecionado
            self.codigo_atual_treeview = codigo_atual_treeview
            
        combobox_cpfs.bind("<<ComboboxSelected>>", on_cpf_selected)
    
    def adicionar_a_lista(self):
        indice_selecionado = self.listbox_codigos_db.curselection()
        if not indice_selecionado:
            messagebox.showwarning("Aviso", "Selecione um código disponível para adicionar à lista.")
            return

        codigo_selecionado = self.listbox_codigos_db.get(indice_selecionado)
        
        # Presumindo que o CPF já foi selecionado e está armazenado em self.cpf_selecionado
        cpf_selecionado = self.cpf_selecionado
        codigo_atual = self.codigo_atual_treeview

        self.adicionar_alteracao_pendente(cpf_selecionado, codigo_atual, codigo_selecionado)

    def adicionar_alteracao_pendente(self, cpf, codigo_atual, codigo_novo):
        # Verificando se já existe uma alteração pendente para o mesmo CPF
        for alteracao in self.alteracoes_pendentes:
            cpf_existente, _, _ = alteracao
            if cpf_existente == cpf:
                messagebox.showwarning("Aviso", f"Já existe uma alteração pendente para o CPF: {cpf}.")
                self.janela_resultados.focus_set()
                return

        # Verificando se o código novo já existe no view_mapping
        if codigo_novo in self.mapeamento_codigos:
            messagebox.showwarning("Aviso", f"O código {codigo_novo} já está mapeado.")
            self.janela_resultados.focus_set()
            return

        # Adicionando a alteração pendente se não houver conflito
        self.alteracoes_pendentes.append((cpf, codigo_atual, codigo_novo))
        self.atualizar_listbox_alteracoes_pendentes()


    def atualizar_listbox_alteracoes_pendentes(self):
        self.listbox_alteracoes_pendentes.delete(0, tk.END)
        for cpf, codigo_atual, codigo_novo in self.alteracoes_pendentes:
            self.listbox_alteracoes_pendentes.insert(tk.END, f"{codigo_atual} -> {codigo_novo}")

    def importar_para_view_mapping(self):
        dados_para_importar = [f"{codigo_atual} -> {codigo_novo}" for _, codigo_atual, codigo_novo in self.alteracoes_pendentes]
        for _, codigo_atual, codigo_novo in self.alteracoes_pendentes:
            self.mapeamento_codigos[str(codigo_atual)] = str(codigo_novo)
        
        self.save_codigos()
        self.view_mapping(dados_para_importar)
    
    def atualizar_vinculo(self, tree):
        indice_coluna_vinculo = 8  # Coluna do Vínculo
        for item in tree.get_children():
            valores = list(tree.item(item, 'values'))
            novo_valor = self.mapear_vinculo(valores[indice_coluna_vinculo])
            if novo_valor is not None:
                valores[indice_coluna_vinculo] = novo_valor
                tree.item(item, values=valores)

    def mapear_vinculo(self, descricao_vinculo):
        mapeamento_vinculos = {
            '30 - Regime Juídico Vinculado a Regime Próprio': '1',
            '45 - Comissão': '2',
            '65 - Contrato Temporário': '3',
            '35 - Servidor Público Não Efetivo': '4',
            '55 - Prestação de Serviços': '5',
            '90 - Estatutário': '6',
            '50 - Temporário': '8',
        }
        return mapeamento_vinculos.get(descricao_vinculo)

        
    def adicionar_coluna_codigo_empresa(self, tree, database, nome_tabela):
        nome_tabela = 'empresas'
        coluna_empresa = "Empresa"
        colunas_atuais = list(tree["columns"])

        # Adiciona a nova coluna se ainda não existir
        if coluna_empresa not in colunas_atuais:
            colunas_atualizadas = colunas_atuais + [coluna_empresa]
            tree.configure(columns=colunas_atualizadas)

            for coluna in colunas_atualizadas:
                tree.heading(coluna, text=coluna)
                tree.column(coluna, width=100) 

        for item in tree.get_children():
            valores = list(tree.item(item, 'values'))
            descricao_secretaria = valores[3] 

            descricao_para_busca = self.extrair_descricao_para_busca(descricao_secretaria)
            codigo_empresa = buscar_empresa_por_descricao(database, nome_tabela, descricao_para_busca)
            
            if len(valores) < len(colunas_atualizadas):
                valores.append(codigo_empresa)
            else:
                valores[-1] = codigo_empresa

            tree.item(item, values=valores)

    def extrair_descricao_para_busca(self, descricao):
        palavras_chave = ["educacao", "saude", "assistencia"]      
        descricao_lower = descricao.lower()

        # Verificando cada palavra-chave na descrição
        for palavra in palavras_chave:
            if palavra in descricao_lower:
                return palavra

        return "pref"
    
    def atualizar_dados_e_vinculo(self, tree):
        self.atualizar_descricoes_por_codigos(tree)
        self.atualizar_vinculo(tree)

    def on_item_double_click(self, event, tree):
        item = tree.identify('item', event.x, event.y)
        column = tree.identify_column(event.x)
        col_index = int(column[1:]) - 1 

        try:
            current_value = tree.item(item, 'values')[col_index]
        except IndexError:
            current_value = ""

        new_value = simpledialog.askstring("Editar Célula", "Editar:", initialvalue=current_value)
        if new_value is not None:
            values = list(tree.item(item, 'values'))
            
            if len(values) <= col_index:
                values.extend([None] * (col_index - len(values) + 1))

            values[col_index] = new_value
            tree.item(item, values=values)
 
    def verificar_centros_custos_e_secretarias_treeview(self, tree, nome_tabela_centro_custos, nome_tabela_secretaria):
        centros_custos_unicos = set()
        secretarias_unicas = set()
        combinacoes_nao_existentes = set()
        centros_custos_nao_existentes = []
        secretarias_nao_existentes = []

        for item in tree.get_children():
            valores = tree.item(item, 'values')
            descricao_centro_custos = valores[2]
            descricao_secretaria = valores[3]
            cargo = valores[4] 
            salario_valor = valores[-1]
            nome= valores[1]

            salario_valor_formatado = salario_valor.replace('R$', '').replace('.', '').replace(',', '.')

            try:
                salario_valor_decimal = Decimal(salario_valor_formatado).quantize(Decimal('0.00'))
            except InvalidOperation as e:
                error_message = f"[verificar_Dados] Valor de salário inválido para o nome: '{nome} - '{cargo}': {salario_valor}."
                logger.error(error_message)
                messagebox.showerror("Erro", error_message)
                continue

            centros_custos_unicos.add(descricao_centro_custos)
            secretarias_unicas.add(descricao_secretaria)

            resultado = verificar_cargo_e_salario(self.database, cargo[:50], salario_valor_decimal)
            if resultado:  
                combinacao = f"{cargo} - {salario_valor}"
                combinacoes_nao_existentes.add(combinacao)

        for centro in centros_custos_unicos:
            if not verificar_existencia(self.database, nome_tabela_centro_custos, centro):
                centros_custos_nao_existentes.append(centro)
            
        for secretaria in secretarias_unicas:
            if not verificar_existencia(self.database, nome_tabela_secretaria, secretaria):
                secretarias_nao_existentes.append(secretaria)

        self.mostrar_dados_nao_existentes(centros_custos_nao_existentes, secretarias_nao_existentes, combinacoes_nao_existentes)

    def mostrar_dados_nao_existentes(self, centros_custos_nao_existentes, secretarias_nao_existentes, combinacoes_nao_existentes):
        new_window = tk.Toplevel(self.root)
        new_window.title("Dados não existentes")
        new_window.grab_set()

        main_frame = tk.Frame(new_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        frame_centro_custos = tk.Frame(main_frame, bd=2, relief=tk.GROOVE)
        frame_secretarias = tk.Frame(main_frame, bd=2, relief=tk.GROOVE)
        frame_combinacoes = tk.Frame(main_frame, bd=2, relief=tk.GROOVE)

        frame_centro_custos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        frame_secretarias.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0)) 
        frame_combinacoes.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0)) 
        
        label_cc = tk.Label(frame_centro_custos, text="Centros de Custos não existentes")
        label_cc.pack()
        self.listbox_cc = tk.Listbox(frame_centro_custos)
        for centro in centros_custos_nao_existentes:
            self.listbox_cc.insert(tk.END, centro)
        self.listbox_cc.pack(fill=tk.BOTH, expand=True)

        label_sec = tk.Label(frame_secretarias, text="Secretarias não existentes")
        label_sec.pack()
        self.listbox_sec = tk.Listbox(frame_secretarias)
        for secretaria in secretarias_nao_existentes:
            self.listbox_sec.insert(tk.END, secretaria)
        self.listbox_sec.pack(fill=tk.BOTH, expand=True)

        label_combinacoes = tk.Label(frame_combinacoes, text="Combinações de Cargos e Salários não existentes")
        label_combinacoes.pack()
        self.listbox_combinacoes = tk.Listbox(frame_combinacoes)
        for combinacao in combinacoes_nao_existentes:
            self.listbox_combinacoes.insert(tk.END, combinacao)
        self.listbox_combinacoes.pack(fill=tk.BOTH, expand=True)

        frame_botoes = tk.Frame(new_window)
        frame_botoes.pack(fill=tk.X, pady=(5, 10)) 

        btn_adicionar_todos = tk.Button(frame_botoes, text="Adicionar Todos os Dados", command=lambda: self.adicionar_todos_os_dados(centros_custos_nao_existentes, secretarias_nao_existentes, combinacoes_nao_existentes))
        btn_adicionar_todos.pack(pady=5, padx=10)

    def adicionar_centros_custos(self, centros_custos_nao_existentes):
        nome_tabela = 'centrocusto'
        adicionados_com_sucesso = []

        for descricao in centros_custos_nao_existentes:
            try:
                criar_registro(self.database, nome_tabela, descricao)
                adicionados_com_sucesso.append(descricao)
            except Exception as e:
                print(f"Erro ao adicionar descrição '{descricao}'. Erro: {e}")

        return adicionados_com_sucesso

    def adicionar_secretarias(self, secretarias_nao_existentes):
        nome_tabela = 'localizacao'
        adicionadas_com_sucesso = []

        for descricao in secretarias_nao_existentes:
            try:
                criar_registro(self.database, nome_tabela, descricao)
                adicionadas_com_sucesso.append(descricao)
            except Exception as e:
                print(f"Erro ao adicionar secretaria '{descricao}'. Erro: {e}")

        return adicionadas_com_sucesso

    def adicionar_funcoes(self, combinacoes):
        adicionadas_com_sucesso = []

        for combinacao in combinacoes:
            descricao, faixa_salarial = combinacao.rsplit(' - ', 1)
            faixa_salarial = Decimal(faixa_salarial.replace('R$', '').replace('.', '').replace(',', '.')).quantize(Decimal('0.00'))
            
            descricao_truncada = descricao[:50]
            faixa_salarial_str = str(faixa_salarial)

            try:
                adicionar_funcao(self.database, descricao_truncada, faixa_salarial_str)
                adicionadas_com_sucesso.append(combinacao)
            except Exception as e:
                print(f"Erro ao adicionar função '{combinacao}'. Erro: {e}")

        return adicionadas_com_sucesso
    
    def adicionar_todos_os_dados(self, centros_custos_nao_existentes, secretarias_nao_existentes, combinacoes_nao_existentes):
        centros_custos_adicionados = self.adicionar_centros_custos(centros_custos_nao_existentes)
        secretarias_adicionadas = self.adicionar_secretarias(secretarias_nao_existentes)
        funcoes_adicionadas = self.adicionar_funcoes(combinacoes_nao_existentes)

        self.atualizar_listbox_apos_adicao(self.listbox_cc, centros_custos_adicionados)
        self.atualizar_listbox_apos_adicao(self.listbox_sec, secretarias_adicionadas)
        self.atualizar_listbox_apos_adicao(self.listbox_combinacoes, funcoes_adicionadas)

        messagebox.showinfo("Concluído", "Todos os dados não existentes foram adicionados.")
        
    def atualizar_listbox_apos_adicao(self, listbox, itens_adicionados):
        itens_listbox = list(listbox.get(0, tk.END))
        for item in itens_listbox:
            if item in itens_adicionados:
                index = listbox.get(0, tk.END).index(item)
                listbox.delete(index)

    def atualizar_descricoes_por_codigos(self, tree):
        for item in tree.get_children():
            valores = tree.item(item, 'values')
            descricao_centro_custos = valores[2]
            descricao_secretaria = valores[3]
            descricao_cargo = valores[4] 
            salario_valor = valores[-2] 

            try:
                # Removendo 'R$' e substituindo ',' por '.'
                salario_valor_limpo = salario_valor.replace('R$', '').replace('.', '').replace(',', '.')
                salario_valor_formatado = Decimal(salario_valor_limpo).quantize(Decimal('0.00'))
            except InvalidOperation:
                print(f"Erro de conversão para o valor do salário: {salario_valor}")
                continue

            codigo_centro_custos = verificar_existencia(self.database, 'centrocusto', descricao_centro_custos)
            codigo_secretaria = verificar_existencia(self.database, 'localizacao', descricao_secretaria)
            codigo_funcao = verificar_codigo_funcao(self.database, descricao_cargo, salario_valor_formatado)

            novos_valores = list(valores)
            if codigo_centro_custos:
                novos_valores[2] = codigo_centro_custos
            if codigo_secretaria:
                novos_valores[3] = codigo_secretaria
            if codigo_funcao:
                novos_valores[4] = codigo_funcao

            tree.item(item, values=novos_valores)

    def cadastrar_todos_funcionarios(self, tree):
        lista_dados_funcionarios = []

        for item in tree.get_children():
            dados_brutos = tree.item(item, 'values')
            nome = dados_brutos[1]
            dados_funcionario = []

            for i, valor in enumerate(dados_brutos):
                try:
                    if i == 12:  # Coluna do Salario
                        valor_convertido = float(valor.replace('.', '').replace(',', '.'))
                    elif i in [9, 10]:  # Colunas dataadm e nascimento
                        valor_convertido = datetime.datetime.strptime(valor, '%d/%m/%Y')
                    else:
                        valor_convertido = valor
                except ValueError as e:
                    error_message = f"Erro na linha '{nome}', coluna {i}: {valor}. Detalhes: {str(e)}"
                    logger.error(error_message)
                    messagebox.showerror("Erro", error_message)
                    return

                dados_funcionario.append(valor_convertido)
            lista_dados_funcionarios.append(dados_funcionario)

        sucesso, mapeamento_novos_codigos = inserir_funcionarios_no_banco(self.database, lista_dados_funcionarios, self.root)
        print("Mapeamento após inserção:", mapeamento_novos_codigos)

        if sucesso:

            for codigo_original, codigo_novo in mapeamento_novos_codigos.items():
                self.mapeamento_codigos[str(codigo_original)] = str(codigo_novo)

            print("Antes de salvar mapeamento:", self.mapeamento_codigos)
            self.save_codigos() 
            print("Depois de salvar mapeamento:", self.mapeamento_codigos)
            self.update_listbox(self.listbox_mapping, 'mapping')
    
    def atualizar_listbox_mapping(self):
        self.listbox_mapping.delete(0, tk.END)
        for codigo_original, codigo_novo in self.mapeamento_codigos.items():
            self.listbox_mapping.insert(tk.END, f"{codigo_original} -> {codigo_novo}")


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

    def import_events_from_pdf(self):
        if not self.filename:
            messagebox.showwarning("Aviso", "Por favor, selecione um arquivo PDF primeiro.")
            return

        try:
            proventos, descontos = extrair_e_categorizar_dados(self.filename)
            
            self.codigos_proventos.update({codigo: descricao for codigo, descricao, _, _ in proventos})
            self.codigos_desconto.update({codigo: descricao for codigo, descricao, _, _ in descontos})
            self.save_codigos()         
            self.update_listbox(self.listbox_proventos, 'provento')
            self.update_listbox(self.listbox_descontos, 'desconto')
            
            messagebox.showinfo("Sucesso", "Eventos importados com sucesso do PDF.")
        except Exception as e:
            messagebox.showerror("Erro", str(e))


    def view_codes(self):
        if 'view_codes' in self.popup_windows and self.popup_windows['view_codes'].winfo_exists():
            self.popup_windows['view_codes'].lift()
            return
        
        popup = tk.Toplevel(self.root)
        popup.title("Visualizar Códigos Adicionados")
        self.popup_windows['view_codes'] = popup 
        popup.geometry("500x500")
        popup.grab_set()

        frame_proventos = tk.Frame(popup)
        frame_proventos.pack(pady=10, padx=10, fill=tk.X)

        self.label_proventos = tk.Label(frame_proventos, text="Proventos", font=("Arial", 12))
        self.label_proventos.pack()

        self.listbox_proventos = tk.Listbox(frame_proventos)
        self.update_listbox(self.listbox_proventos, 'provento')
        self.listbox_proventos.pack(fill=tk.X)

        btn_add_provento = tk.Button(frame_proventos, text="Adicionar", command=lambda: self.add_provento(self.listbox_proventos))
        btn_add_provento.pack(side=tk.LEFT, padx=5)

        btn_edit_provento = tk.Button(frame_proventos, text="Editar", command=lambda: self.edit_code(self.listbox_proventos, 'provento'))
        btn_edit_provento.pack(side=tk.LEFT, padx=5)

        btn_remove_provento = tk.Button(frame_proventos, text="Remover", command=lambda: self.remove_code(self.listbox_proventos, 'provento'))
        btn_remove_provento.pack(side=tk.LEFT, padx=5)

        btn_import_events = tk.Button(frame_proventos, text="Importar do PDF", command=self.import_events_from_pdf)
        btn_import_events.pack(side=tk.LEFT, padx=5)

        frame_descontos = tk.Frame(popup)
        frame_descontos.pack(pady=10, padx=10, fill=tk.X)

        self.label_descontos = tk.Label(frame_descontos, text="Descontos", font=("Arial", 12))
        self.label_descontos.pack()

        self.listbox_descontos = tk.Listbox(frame_descontos)
        self.update_listbox(self.listbox_descontos, 'desconto') 
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
        if listbox.winfo_exists():  # Verifica se o widget ainda existe
            listbox.delete(0, tk.END)  

        if tipo == 'provento':
            items = self.codigos_proventos.items()
            total = len(items)
            if self.label_proventos:
                self.label_proventos.config(text=f"Proventos ({total})")
        elif tipo == 'desconto':
            items = self.codigos_desconto.items()
            total = len(items)
            if self.label_descontos:
                self.label_descontos.config(text=f"Descontos ({total})")
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
        print("Mapeamento salvo:", self.mapeamento_codigos)

    def view_mapping(self, dados_importados=None):
        if 'view_mapping' in self.popup_windows and self.popup_windows['view_mapping'].winfo_exists():
            self.popup_windows['view_mapping'].lift()
            return
        popup = tk.Toplevel(self.root)
        self.popup_windows['view_mapping'] = popup
        popup.title("Visualizar Mapeamento de Códigos")
        popup.geometry("500x500")
        popup.grab_set

        frame_mapping = tk.Frame(popup)
        frame_mapping.pack(pady=10, padx=10, fill=tk.X)

        label_mapping = tk.Label(frame_mapping, text="Mapeamento de Códigos", font=("Arial", 12))
        label_mapping.pack()

        scrollbar = ttk.Scrollbar(frame_mapping, orient="vertical")
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox_mapping = tk.Listbox(frame_mapping, yscrollcommand=scrollbar.set)
        self.update_listbox(self.listbox_mapping, 'mapping')
        self.listbox_mapping.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar.config(command=self.listbox_mapping.yview)

        if dados_importados:
            for item in dados_importados:
                self.listbox_mapping.insert(tk.END, item)

        for codigo_original, codigo_novo in self.mapeamento_codigos.items():
            # Adicione o mapeamento à interface, por exemplo, a uma listbox
            self.listbox_mapping.insert(tk.END, f"{codigo_original} -> {codigo_novo}")

        # Frame para os botões
        button_frame = tk.Frame(popup)
        button_frame.pack(fill=tk.X)

        btn_add_mapping = tk.Button(button_frame, text="Adicionar", command=self.add_mapping)
        btn_add_mapping.pack(side=tk.LEFT, padx=5)

        btn_edit_mapping = tk.Button(button_frame, text="Editar", command=self.edit_mapping)
        btn_edit_mapping.pack(side=tk.LEFT, padx=5)

        btn_remove_mapping = tk.Button(button_frame, text="Remover", command=self.remove_mapping)
        btn_remove_mapping.pack(side=tk.LEFT, padx=5)

        self.button_clear_all_mappings = tk.Button(button_frame, text="Remover Todos", command=self.clear_all_mappings)
        self.button_clear_all_mappings.pack(side=tk.LEFT, padx=5)

        btn_close = tk.Button(button_frame, text="Fechar", command=lambda: [self.save_codigos(), popup.destroy()])
        btn_close.pack(pady=20)
        popup.wait_window()
        
    def add_mapping(self):
        new_code = simpledialog.askstring("Adicionar código/matricula", "Digite o código para ser substituido:")
        if new_code:
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

    def save_excel(self, tree=None):
        if tree is not None:
            # Atualizando as colunas do DataFrame para corresponder às da TreeView
            colunas_treeview = tree["columns"]
            novos_dados = []

            for item in tree.get_children():
                valores = tree.item(item, 'values')
                novos_dados.append(valores)

            # Criando um novo DataFrame com os dados da TreeView
            df_atualizado = pd.DataFrame(novos_dados, columns=colunas_treeview)
        else:
            df_atualizado = self.df

        excel_filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos XLSX", "*.xlsx")])
        if excel_filename:
            with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
                df_atualizado.to_excel(writer, index=False, sheet_name='Planilha')
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
    setup_logging()
    root = tk.Tk()
    app = App(root)
    root.mainloop()