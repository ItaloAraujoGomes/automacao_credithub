import pdfplumber
import pandas as pd
import os

# 📂 Caminho do PDF e do Excel de saída
pdf_folder = r"C:\Users\7981\Desktop\relatorios_fidc"
pdf_file = "RELAT. PENDENCIA BAIXA.pdf"
pdf_path = os.path.join(pdf_folder, pdf_file)
output_excel = os.path.join(pdf_folder, "resultado_pendencias_baixadas.xlsx")

# Cria o diretório, se não existir
output_folder = os.path.dirname(output_excel)
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 📌 Dicionário para armazenar DataFrames separados por tipo de relatório
report_data = {}
processed_files = []

# Lista com os nomes das colunas que você deseja definir, incluindo colunas com espaços
colunas_personalizadas_pendencias = [
    "Borderô", "SeqDocto.", "  ", "Vencto", "   ", "Vcto Ant", "Sacado", "Dt. Pend.", 
    "Motivo", "    ", "Tipo", "      ", "Vr. Título", "Pendência", "Despesa", "Descto",
     "        ", "Vr. Final"
]

def extract_table_from_pdf(pdf_path):
    """Extrai tabelas do PDF e formata os dados."""
    data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table:
                    df = pd.DataFrame(table)

                    # Verifica se a tabela tem colunas suficientes para ser relevante
                    if df.empty or df.shape[1] < len(colunas_personalizadas_pendencias):
                        continue

                    # Ajusta o número de colunas para garantir que todos os campos sejam mantidos
                    num_cols = df.shape[1]
                    if num_cols <= len(colunas_personalizadas_pendencias):
                        df.columns = colunas_personalizadas_pendencias[:num_cols]
                    else:
                        # Adiciona "colunas extras" para preencher o DataFrame se necessário
                        df.columns = colunas_personalizadas_pendencias + ['Coluna Extra {}'.format(i) for i in range(num_cols - len(colunas_personalizadas_pendencias))]

                    # Remove linhas completamente vazias
                    df = df.dropna(how='all')

                    # Limpa espaços extras (usando apply ao invés de applymap)
                    df = df.apply(lambda x: x.strip() if isinstance(x, str) else x)

                    # Resetando o índice para garantir que seja único
                    df.reset_index(drop=True, inplace=True)

                    data.append(df)

    return pd.concat(data, ignore_index=True) if data else pd.DataFrame()

def process_pdf(pdf_path):
    """Processa o PDF e armazena os dados extraídos."""
    df = extract_table_from_pdf(pdf_path)
    sheet_name = "Pendências Baixadas"

    if not df.empty:
        # Remover palavras indesejadas nas células do DataFrame
        palavras_indesejadas = [
            "Capital", "Finanças", "Fomento", "Mercantil", "LTDA", 
            "Extrato de Conta", "Data do Lançamento", "Página", 
            "Usuário", "ACA", "Pendências Genéricas", "Prorrogação de Títulos", 
            "Origem dos Lançamentos", "Doctº", "Data", "Saldo", "JUROS S/PRORROGACOES","90.03.0003", "Saldo Inicial"
        ]
        
        # Substituir palavras indesejadas por uma string vazia
        df = df.applymap(lambda x: ' ' if isinstance(x, str) and any(palavra in x for palavra in palavras_indesejadas) else x)
        
        # Atualiza os dados
        if sheet_name in report_data:
            # Resetando o índice para garantir que os DataFrames concatenados tenham índices únicos
            df.reset_index(drop=True, inplace=True)
            report_data[sheet_name] = pd.concat([report_data[sheet_name], df], ignore_index=True, sort=False)
        else:
            report_data[sheet_name] = df
        processed_files.append(os.path.basename(pdf_path))

# 📌 Executa apenas se o arquivo existir
if os.path.exists(pdf_path):
    print(f"📄 Processando: {pdf_file}")
    process_pdf(pdf_path)

# 📝 Salva os dados extraídos no Excel
if report_data:
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        for sheet, df in report_data.items():
            df.to_excel(writer, sheet_name=sheet, index=False)
    print(f"✅ Arquivo salvo: {output_excel}")
else:
    print("⚠ Nenhum dado extraído do PDF.")

# 📊 Exibe relatório de processamento
print(f"📂 PDFs processados: {len(processed_files)}")
for file in processed_files:
    print(f" - {file}")
