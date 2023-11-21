from tkinter import messagebox
import pyodbc
import socket

def criar_conexao_sql_server(database_sqlserver = None):
    try:
        server = f'{socket.gethostname()}\\SQL2008'
        username = 'SA'
        password = ''
        driver = '{SQL Server Native Client 10.0}'
        
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
    query = f"SELECT MAX(CAST(CODIGO AS INT)) FROM {nome_tabela}"
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

def buscar_empresa_por_descricao(database, nome_tabela, descricao):
    cursor = database.cursor()
    descricao_like = f"%{descricao}%"
    query = f"SELECT codigo, razaosocial FROM {nome_tabela} WHERE razaosocial LIKE ?"
    cursor.execute(query, (descricao_like,))
    resultado = cursor.fetchone()
    return resultado[0] if resultado else None

def verificar_cargo_e_salario(database, cargo, salario_valor):
    try:

        cursor = database.cursor()
        descricao_truncada = cargo[:50]
        query = "SELECT codigo, descricao, faixasalarial FROM funcoes WHERE descricao = ? AND faixasalarial = ?"
        
        cursor.execute(query, (descricao_truncada, salario_valor))
        resultado = cursor.fetchone()
        if resultado:
            print("Resultado encontrado:", resultado)
        else:
            print("Nenhum resultado encontrado.")

        # Se o resultado for None, o cargo com essa faixa salarial não existe
        return resultado is None
    except ValueError as ve:
        messagebox.showerror(f"Valor de salário inválido para a descrição '{cargo}' e valor '{salario_valor}': {ve}")
        return False
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao verificar cargo e salário: {e}")
        return False

def adicionar_funcao(database, descricao, faixa_salarial):
    try:

        codigo = obter_proximo_codigo(database, 'funcoes')

        # Preparar a query SQL com SET IDENTITY_INSERT
        cursor = database.cursor()
        cursor.execute("SET IDENTITY_INSERT funcoes ON")

        insert_query = "INSERT INTO funcoes (codigo, descricao, faixasalarial) VALUES (?, ?, ?)"
        cursor.execute(insert_query, (codigo, descricao, faixa_salarial))

        cursor.execute("SET IDENTITY_INSERT funcoes OFF")
        database.commit()

        return True
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao adicionar a função no banco de dados: {e}")
        return False

def verificar_codigo_funcao(database, cargo, salario_valor):
    try:
        cursor = database.cursor()
        descricao_truncada = cargo[:50]
        query = "SELECT codigo FROM funcoes WHERE descricao = ? AND faixasalarial = ?"
        
        cursor.execute(query, (descricao_truncada, salario_valor))
        resultado = cursor.fetchone()
        
        if resultado:
            print("Resultado encontrado:", resultado)
            return resultado[0]
        else:
            print("Nenhum resultado encontrado.")
            return None  

    except ValueError as ve:
        messagebox.showerror(f"Valor de salário inválido para a descrição '{cargo}' e valor '{salario_valor}': {ve}")
        return None
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao verificar cargo e salário: {e}")
        return None
    
def inserir_funcionario_no_banco(database, dados_funcionario):
    codigo = obter_proximo_codigo(database, "funcionarios")
    with database.cursor() as cursor:
        cursor.execute("SET IDENTITY_INSERT funcionarios ON")
        query = "INSERT INTO funcionarios (codigo, registro, nome, centrocusto, localizacao, funcao, cpf, agencia, contacorrente, tipovinculo, dataadm, nascimento, salario, empresa) VALUES (?,?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        parametros = (codigo, codigo, dados_funcionario[1], dados_funcionario[2], dados_funcionario[3], dados_funcionario[4], dados_funcionario[5], dados_funcionario[6], dados_funcionario[7], dados_funcionario[8], dados_funcionario[9], dados_funcionario[10], dados_funcionario[12], dados_funcionario[13])
        print("Query SQL:", query)
        print("Parâmetros:", parametros)
        try:
            cursor.execute(query, parametros)
            cursor.execute("SET IDENTITY_INSERT funcionarios OFF")
            database.commit()
            messagebox.showinfo("Sucesso", "Funcionário cadastrado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao cadastrar funcionário: {e}")







