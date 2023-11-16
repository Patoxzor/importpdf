from tkinter import messagebox
import pyodbc
import socket

def criar_conexao_sql_server(database_sqlserver = None):
    try:
        server = f'{socket.gethostname()}\\SQL2008'
        username = 'SA'
        password = ''
        driver = '{SQL Server Native Client 10.0}' # Driver for SQL Server 2008
        
        if database_sqlserver:
            conn_str = f'DRIVER={driver};SERVER={server};DATABASE={database_sqlserver};UID={username};PWD={password}'
        else:
            conn_str = f'DRIVER={driver};SERVER={server};UID={username};PWD={password}'
        conn = pyodbc.connect(conn_str)
        return conn
    except pyodbc.Error as e:
        messagebox.showerror("Aviso", "Não foi possível estabelecer uma conexão com o SQL Server.")

def get_sql_server_databases():
    with criar_conexao_sql_server() as conn:
        cursor = conn.cursor()
        query = "SELECT name FROM sys.databases"
        cursor.execute(query)
        # Lista de bancos de dados padrão do SQL Server para excluir
        default_databases = ['master', 'tempdb', 'model', 'msdb', 'ReportServer$SQL2008', 'ReportServer$SQL2008TempDB']
        databases = [row[0] for row in cursor.fetchall() if row[0] not in default_databases and row[0] != 'FOLHA']
        return ['FOLHA'] + databases
    
def verificar_funcionario(database, codigo, cpf):
    try:
        with criar_conexao_sql_server(database) as conn:
            cursor = conn.cursor()
            query = "SELECT CODIGO, NOME, CPF, ATIVO FROM funcionarios WHERE Codigo = ? AND CPF = ?"
            cursor.execute(query, codigo, cpf)
            result = cursor.fetchone()
            cursor.close()
            return result
    except pyodbc.Error as e:
        messagebox.showerror("Erro", f"Erro ao verificar o funcionário no banco de dados: {e}")
        return None

# Função para verificar todos os funcionários extraídos do PDF
def verificar_todos_funcionarios(database, lista_funcionarios):
    funcionarios_nao_encontrados = []
    for funcionario in lista_funcionarios:
        codigo = funcionario['Codigo']
        cpf = funcionario['CPF']
        if not verificar_funcionario(database, codigo, cpf):
            funcionarios_nao_encontrados.append(funcionario)
    return funcionarios_nao_encontrados

def verificar_existencia(database, nome_tabela, descricao):
    cursor = database.cursor()
    query = f"SELECT codigo, descricao FROM {nome_tabela} WHERE descricao = ?"
    cursor.execute(query, (descricao,))
    resultado = cursor.fetchone()
    return resultado[0] if resultado else None

def obter_proximo_codigo(database, nome_tabela):
    cursor = database.cursor()
    query = f"SELECT MAX(codigo) FROM {nome_tabela}"
    cursor.execute(query)
    resultado = cursor.fetchone()
    return (int(resultado[0]) + 1) if resultado and resultado[0] is not None else 1


def criar_registro(database, nome_tabela, descricao):
    novo_codigo = obter_proximo_codigo(database, nome_tabela)
    novo_codigo_str = str(novo_codigo) 
    query = f"INSERT INTO {nome_tabela} (codigo, descricao) VALUES (?, ?)"
    cursor = database.cursor()
    cursor.execute(query, (novo_codigo_str, descricao))
    database.commit()

def obter_codigo_e_descricao(self, nome_tabela, descricao):
    try:
        cursor = self.database.cursor()
        query = f"SELECT codigo, descricao FROM {nome_tabela} WHERE descricao = ?"
        cursor.execute(query, (descricao,))
        resultado = cursor.fetchone()
        return resultado if resultado else (None, None)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao obter código e descrição: {e}")
        return None, None
