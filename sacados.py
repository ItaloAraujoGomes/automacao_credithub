import pyodbc
import pandas as pd
from datetime import datetime, timedelta

# Configuração do banco de dados
server = '192.168.1.254'
database = 'POSICAO'
username = 'pwrbi.user'
password = 'Capital@2024Out#'

# Ajustando a string de conexão para usar o driver correto
conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password};TrustServerCertificate=yes'

try:
    # Estabelecendo a conexão
    conn = pyodbc.connect(conn_str)
    print("Conexão bem-sucedida!")

    # Definindo o intervalo de datas (últimos 5 dias)
    hoje = datetime.now()
    cinco_dias_atras = hoje - timedelta(days=5)

    # Consulta SQL
    query = f"""
    SELECT 
        Nome_cedente,
        Nome_sacado,
        Cep_Sacado,
        Endereco_Sacado,
        Bairro_Sacado,
        Cidade_Sacado,
        OP
    FROM 
        POSICAO.View_Titulos_em_Aberto
    WHERE 
        Dt_Vencto_Original >= '{cinco_dias_atras.strftime('%Y-%m-%d')}' AND
        Dt_Vencto_Original <= '{hoje.strftime('%Y-%m-%d')}'
    """

    # Extração dos dados
    df = pd.read_sql(query, conn)

    # Nome do arquivo Excel
    file_name = f"Sacados_Prospecao_{hoje.strftime('%Y-%m-%d')}.xlsx"

    # Salvando em Excel
    df.to_excel(file_name, index=False)

    print(f"Arquivo gerado com sucesso: {file_name}")

    # Fechando a conexão
    conn.close()

except Exception as e:
    print(f"Erro na conexão ou na execução da consulta: {e}")
