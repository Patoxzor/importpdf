from tkinter import messagebox, simpledialog
import pyodbc
import socket
from log_erros import logger

def criar_conexao_sql_server(database_sqlserver = None):
    try:
        server = f'{socket.gethostname()}\\SQL2019'
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

def obter_codigo_e_descricao(nome_tabela, descricao, database):
    try:
        cursor = database.cursor()
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
    
def inserir_funcionarios_no_banco(database, lista_dados_funcionarios, main_window=None):
    escolha = messagebox.askyesno("Escolha do Código", "Deseja usar os códigos existentes na TreeView?", parent=main_window)
    mapeamento_codigos = {}
    funcionarios_nao_adicionados = []
    sucesso = True

    for dados_funcionario in lista_dados_funcionarios:
        codigo_original = dados_funcionario[0]
        if escolha:
            codigo = codigo_original  # Supondo que o código esteja no índice 0
        else:
            codigo = obter_proximo_codigo(database, "funcionarios")
            if codigo != codigo_original:
                mapeamento_codigos[codigo_original] = codigo

        with database.cursor() as cursor:
            cursor.execute("SET IDENTITY_INSERT funcionarios ON")
            query = "INSERT INTO funcionarios (codigo, registro, nome, centrocusto, localizacao, funcao, cpf, agencia, contacorrente, tipovinculo, dataadm, nascimento, salario, empresa) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            parametros = (codigo, codigo, dados_funcionario[1], dados_funcionario[2], dados_funcionario[3], dados_funcionario[4], dados_funcionario[5], dados_funcionario[6], dados_funcionario[7], dados_funcionario[8], dados_funcionario[9], dados_funcionario[10], dados_funcionario[12], dados_funcionario[13])
            try:
                cursor.execute(query, parametros)
                cursor.execute("SET IDENTITY_INSERT funcionarios OFF")
            except Exception as e:
                error_message = f"[inserir_funcionarios_no_banco] Erro ao inserir funcionário {dados_funcionario[1]} no banco de dados: {e}"
                logger.error(error_message)
                funcionarios_nao_adicionados.append(dados_funcionario[1])  # Supondo que o nome esteja no índice 1
                sucesso = False
                continue

    database.commit()
    

    if funcionarios_nao_adicionados:
        nomes_nao_adicionados = ", ".join(funcionarios_nao_adicionados)
        messagebox.showinfo("Funcionários Não Adicionados", f"Não foi possível adicionar os seguintes funcionários: {nomes_nao_adicionados}")
    elif sucesso:
        messagebox.showinfo("Sucesso", "Todos os funcionários foram cadastrados com sucesso!")

    return sucesso, mapeamento_codigos


def buscar_codigos_por_cpfs(database, lista_cpfs):
    codigos = {}
    cursor = database.cursor()
    for cpf in lista_cpfs:
        try:
            query = "SELECT codigo FROM funcionarios WHERE cpf = ?"
            cursor.execute(query, cpf)
            resultados = cursor.fetchall()
            if resultados:
                codigos[cpf] = [resultado[0] for resultado in resultados]
        except pyodbc.Error as e:
            messagebox.showerror(f"Erro ao buscar código para o CPF {cpf}")
    return codigos

def solicitar_periodo(root):
    """
    Solicita ao usuário que insira o período desejado.
    """
    periodo = simpledialog.askstring("Input", "Digite o período desejado (MM/AAAA):", parent=root)
    return periodo

def inserir_acumulos(database, root, lista_codigos):
    try:
        cursor = database.cursor()
        
        # Solicita ao usuário o período desejado
        periodo = solicitar_periodo(root)
        if not periodo:
            raise ValueError("Período não informado.")
        
        motivo = 0
        status = 1
        
        # Prepara a string SQL com os códigos dos funcionários
        codigos_formatados = ', '.join(f"'{codigo}'" for codigo in lista_codigos)
        
        sql = f"""
        INSERT INTO ACUMULOS (EMPRESA, CENTROCUSTO, FUNCIONARIO, PERIODO,
                        DEPENDENTES, FILHOS, LOCALIZACAO, ATIVO,
                        SALARIO, FUNCAO, GFIP_CATEGORIA, OUTRA_PREVIDENCIA,
                        DATAADM, BANCOSALARIO, CONTACORRENTE, DATADEMISSAO, OPERACAO,
                        AGENCIA, NUMEROCONTRATO, PARTICIPAGPS, PARTICIPASEFIP, STATUS)
        (SELECT DISTINCT EMPRESA, CENTROCUSTO, F.CODIGO, ?, DEPENDENTES, FILHOS,
                LOCALIZACAO, ?, SALARIO, FUNCAO, GFIP_CATEGORIA, OUTRA_PREVIDENCIA,
                DATAADM, BANCOSALARIO, CONTACORRENTE, DATADEMISSAO, OPERACAO,
                AGENCIA, NUMEROCONTRATO, PARTICIPAGPS, PARTICIPASEFIP, ?
        FROM FUNCIONARIOS F
        WHERE NOT EXISTS (SELECT 1 FROM ACUMULOS 
                          WHERE FUNCIONARIO = F.CODIGO 
                          AND PERIODO=?)
                          AND F.Codigo IN ({codigos_formatados}))
        """
        
        cursor.execute(sql, (periodo, status, motivo, periodo))
        database.commit()
        messagebox.showinfo("Concluído", "A inserção dos acumulos foi concluída com sucesso.")
    except ValueError as ve:
        messagebox.showerror("Erro", str(ve))
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao inserir acumulos: {e}")

def atualizar_funcionarios_e_empresa(database, lista_codigos, codigo_empresa):
    # Convertendo a lista de códigos para strings para uso na cláusula SQL IN
    codigos_str = ', '.join(f"'{codigo}'" for codigo in lista_codigos)
    query = f"""
    UPDATE FUNCIONARIOS
    SET Ativo = '1', Empresa = {codigo_empresa}
    WHERE Codigo IN ({codigos_str})
    """
    cursor = database.cursor()
    try:
        cursor.execute(query)
        database.commit()
    except Exception as e:
        database.rollback()  # Desfaz as alterações em caso de erro
        messagebox.showerror("Erro", f"Erro ao atualizar funcionários: {e}")


