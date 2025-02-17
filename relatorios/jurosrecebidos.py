import pdfplumber
import pandas as pd
import os
import re

# Pasta onde estão os PDFs
pdf_folder = r"C:\Users\7981\Desktop\relatorios_fidc"
output_excel = os.path.join(pdf_folder, "juros_recebidos.xlsx")

# Dicionário para armazenar DataFrames separados por tipo de relatório
report_data = {}
processed_files = []

def extract_table_from_pdf(pdf_path):
    """Extrai todas as tabelas de um PDF, combinando múltiplas páginas e limpando os dados"""
    data = []
    header_found = False  # Para identificar quando o cabeçalho já foi encontrado
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            print(f"Tables encontradas na página: {len(tables)}")  # Depuração das tabelas extraídas
            for table in tables:
                if table:
                    df = pd.DataFrame(table)
                    
                    # Verificar se DataFrame está vazio ou não tem colunas suficientes
                    if df.empty or df.shape[1] < 2:
                        continue  # Ignorar páginas sem dados úteis
                    
                    # Encontrar a linha do cabeçalho correto
                    if not header_found:
                        header_row = df[df.apply(lambda row: any("Doctº" in str(cell) for cell in row), axis=1)].index.min()
                        if pd.isna(header_row):
                            continue  # Ignorar páginas sem cabeçalho válido
                        
                        df.columns = df.iloc[int(header_row)]  # Define os nomes das colunas
                        df = df[int(header_row) + 1:]  # Remove a linha original do cabeçalho
                        header_found = True  # Marca que o cabeçalho foi encontrado
                        print(f"Header encontrado: {df.columns.tolist()}")  # Verifique o cabeçalho
                    else:
                        # Usar colunas fixas para todas as páginas
                        fixed_columns = ["Doctº", "Data", "Histórico", "D/C", "Saldo", "Valor D/C"]
                        df.columns = fixed_columns if set(fixed_columns).issubset(df.columns) else df.columns
                    
                    # Remover colunas completamente vazias
                    df = df.dropna(axis=1, how='all')
                    
                    # Remover linhas vazias
                    df = df.dropna(how='all')
                    
                    # Remover linhas que começam com "Movimento do dia:" somente se houver colunas suficientes
                    if df.shape[1] > 0:
                        df = df[~df.iloc[:, 0].astype(str).str.startswith("Movimento do dia:")]
                    
                    # Manter apenas colunas relevantes se elas existirem
                    required_columns = ["Doctº", "Data", "Histórico", "D/C", "Saldo", "Valor D/C"]
                    df = df[required_columns] if set(required_columns).issubset(df.columns) else df
                    
                    # Remover espaços extras
                    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
                    
                    # Verificar o número de colunas
                    if df.shape[1] == 12:
                        # Renomeia apenas as 5 primeiras colunas que são relevantes
                        df.columns = ["Doctº", "Data", "Histórico", "Valor D/C", "Saldo D/C"] + list(df.columns[5:])
                    elif df.shape[1] > 5:
                        # Renomeia as 5 primeiras colunas e mantém as outras
                        df.columns = ["Doctº", "Data", "Histórico", "Valor D/C", "Saldo D/C"] + list(df.columns[5:])
                    data.append(df)
    
    return pd.concat(data, ignore_index=True) if data else pd.DataFrame()

def identify_report_type(filename):
    """Identifica o tipo de relatório com base no nome do arquivo"""
    filename = filename.lower()
    if re.search(r"\bjuros\b|\breceitas\b", filename):
        return "Juros Recebido"
    elif re.search(r"\banalise\b|\boperacional\b", filename):
        return "Resultado Operacional"
    elif "pendencia" in filename:
        return "Pendências em Aberto"
    elif "aberto" in filename:
        return "Relatório em Aberto"
    else:
        return "Outros"

def process_pdf(pdf_path):
    """Processa o PDF e armazena os dados no dicionário"""
    df = extract_table_from_pdf(pdf_path)
    sheet_name = identify_report_type(os.path.basename(pdf_path))
    
    if not df.empty:
        if sheet_name in report_data:
            report_data[sheet_name] = pd.concat([report_data[sheet_name], df], ignore_index=True, sort=False)
        else:
            report_data[sheet_name] = df
        processed_files.append(os.path.basename(pdf_path))

# Percorre todos os PDFs na pasta, mas apenas processa o arquivo específico
total_pdfs = 0
for file in os.listdir(pdf_folder):
    if file.lower().endswith("juros recebido-outras receitas.pdf"):
        pdf_path = os.path.join(pdf_folder, file)
        print(f"Processando: {file}")
        process_pdf(pdf_path)
        total_pdfs += 1

# Salva os dados extraídos em um arquivo Excel
if report_data:
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        for sheet, df in report_data.items():
            df.to_excel(writer, sheet_name=sheet, index=False)
    print(f"Arquivo consolidado salvo como {output_excel}")
else:
    print("Nenhum dado extraído dos PDFs.")

# Exibe relatório de arquivos processados
print(f"Total de PDFs processados: {total_pdfs}")
if processed_files:
    print("Arquivos processados:")
    for file in processed_files:
        print(f" - {file}")
