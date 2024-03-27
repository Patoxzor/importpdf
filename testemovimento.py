import tkinter as tk
from tkinter import ttk, messagebox

# Supondo que esta função esteja dentro da classe App
def criar_tela_movimentos(self):
    # Cria uma nova janela toplevel
    self.janela_movimentos = tk.Toplevel(self.root)
    self.janela_movimentos.title("Adicionar Movimentos")
    self.janela_movimentos.geometry("800x600")  # Ajuste conforme necessário

    # Frame para a Treeview e Scrollbars
    frame_treeview = tk.Frame(self.janela_movimentos)
    frame_treeview.pack(fill=tk.BOTH, expand=True)

    # Cria a Treeview
    colunas = ['Empresa', 'Funcionário', 'Período', 'Tipo de Cálculo']
    self.treeview_movimentos = ttk.Treeview(frame_treeview, columns=colunas, show='headings')
    for coluna in colunas:
        self.treeview_movimentos.heading(coluna, text=coluna)
    self.treeview_movimentos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Scrollbar Vertical
    scrollbar_vertical = ttk.Scrollbar(frame_treeview, orient="vertical", command=self.treeview_movimentos.yview)
    self.treeview_movimentos.configure(yscrollcommand=scrollbar_vertical.set)
    scrollbar_vertical.pack(side=tk.RIGHT, fill=tk.Y)

    # Scrollbar Horizontal
    scrollbar_horizontal = ttk.Scrollbar(frame_treeview, orient="horizontal", command=self.treeview_movimentos.xview)
    self.treeview_movimentos.configure(xscrollcommand=scrollbar_horizontal.set)
    scrollbar_horizontal.pack(side=tk.BOTTOM, fill=tk.X)

    # Frame para Botões
    frame_botoes = tk.Frame(self.janela_movimentos)
    frame_botoes.pack(fill=tk.X)

    # Botão para adicionar movimento
    botao_adicionar_movimento = tk.Button(frame_botoes, text="Adicionar Movimento", command=self.adicionar_movimento_interface)
    botao_adicionar_movimento.pack(side=tk.LEFT, padx=5, pady=5)

    # Preenche a Treeview com dados (exemplo)
    # Aqui, você precisará ajustar para usar os dados que deseja exibir
    for i in range(10):  # Exemplo de dados
        self.treeview_movimentos.insert('', 'end', values=(f'Empresa {i}', f'Funcionário {i}', '01/2023', 'Normal'))

def adicionar_movimento_interface(self):
    # Implementação da lógica para adicionar um movimento.
    # Aqui você pode abrir uma nova janela para entrada de dados ou utilizar uma caixa de diálogo (simpledialog) para coletar informações.
    messagebox.showinfo("Informação", "Aqui você implementa a lógica para adicionar um movimento.")
