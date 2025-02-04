import os
import win32com.client
import pandas as pd
from bs4 import BeautifulSoup  # Para processar arquivos HTML
from datetime import datetime, timedelta, timezone

# Configurações
OUTLOOK_FOLDER = "Inbox"  # Pasta do Outlook
FILTER_SUBJECT = "Gerencie Carteira - Consulte as Empresas Monitoradas"  # Filtro de assunto do e-mail
SAVE_PATH = r"C:\Users\7981\Desktop\Anexos"  # Pasta onde salvar os anexos
OUTPUT_EXCEL = r"C:\Users\7981\Desktop\Dados_Processados.xlsx"  # Caminho da base de dados

# Obter data de 7 dias atrás com timezone UTC
seven_days_ago = datetime.now(timezone.utc) - timedelta(days=7)

# Conectar ao Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # Pasta padrão "Caixa de Entrada"
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)  # Ordena os e-mails do mais recente para o mais antigo

# Criar pasta se não existir
if not os.path.exists(SAVE_PATH):
    os.makedirs(SAVE_PATH)

# Processar e-mails
for message in messages:
    if message.Class == 43:  # Garantir que é uma mensagem de e-mail
        received_date = message.ReceivedTime
        # Converter para timezone UTC
        received_date = datetime.fromtimestamp(received_date.timestamp(), tz=timezone.utc)

        # Filtrar por data e assunto
        if received_date >= seven_days_ago and FILTER_SUBJECT in message.Subject:
            email_date = received_date.strftime("%Y-%m-%d")  # Data do e-mail
            print(f"E-mail encontrado: {message.Subject} - Data: {email_date}")

            # Verifica anexos
            if message.Attachments.Count > 0:
                for attachment in message.Attachments:
                    # Adiciona a data ao nome do arquivo
                    filename, ext = os.path.splitext(attachment.FileName)
                    new_filename = f"{filename}_{email_date}{ext}"
                    file_path = os.path.join(SAVE_PATH, new_filename)

                    attachment.SaveAsFile(file_path)
                    print(f"Anexo salvo: {file_path}")

                    # Processar arquivo HTML
                    if file_path.endswith(".html"):
                        with open(file_path, "r", encoding="utf-8") as file:
                            soup = BeautifulSoup(file, "html.parser")

                        # Localizar a tabela no HTML
                        table = soup.find("table")
                        if table:
                            # Extraímos as linhas da tabela
                            rows = table.find_all("tr")
                            data = []

                            # Itera sobre as linhas da tabela, pulando o cabeçalho
                            for row in rows[1:]:
                                cols = row.find_all("td")
                                if len(cols) == 3:  # Verifica se há 3 colunas por linha
                                    cnpj = cols[0].text.strip()
                                    razao_social = cols[1].text.strip()
                                    alteracao = cols[2].text.strip()
                                    data.append([cnpj, razao_social, alteracao])

                            # Cria um DataFrame a partir dos dados extraídos
                            df = pd.DataFrame(data, columns=['CNPJ', 'Razão Social', 'Alteração'])
                            df['Data de Envio'] = email_date  # Adiciona a data do e-mail

                            # Verifica se a planilha já existe
                            if os.path.exists(OUTPUT_EXCEL):
                                df_existing = pd.read_excel(OUTPUT_EXCEL)
                                df = pd.concat([df_existing, df], ignore_index=True)  # Adiciona novos dados

                            # Remove duplicatas com base em CNPJ e Alteração
                            df.drop_duplicates(subset=['CNPJ', 'Alteração'], keep='last', inplace=True)

                            # Salva a planilha
                            
                            df.to_excel(OUTPUT_EXCEL, index=False)
                            print(f"Base de dados atualizada em {OUTPUT_EXCEL}")
                        else:
                            print("Tabela não encontrada no HTML")

print("Processo concluído!")


