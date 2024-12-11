import pandas as pd
import win32com.client
from datetime import datetime

# Conectar ao Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 é o código para "Inbox" ou "Caixa de Entrada"

# Filtros de assunto e data
subject_filter_1 = "Aumento em Protesto Detectado!"
subject_filter_2 = "Redução de Protestos Detectada!"
today = datetime.now().date()

# Lista para armazenar os dados
data = []

for message in inbox.Items:
    # Filtrar por assunto e data
    if (subject_filter_1 in message.Subject or subject_filter_2 in message.Subject) and message.ReceivedTime.date() == today:
        # Extrair informações do corpo do e-mail
        body = message.Body
        lines = body.splitlines()

        razao_social, cpf_cnpj, antes, atual = None, None, None, None

        for line in lines:
            if line.startswith("Nome/Razão Social:"):
                razao_social = line.split("Nome/Razão Social:")[1].strip()
            elif line.startswith("CPF/CNPJ:"):
                cpf_cnpj = line.split("CPF/CNPJ:")[1].strip()
            elif line.startswith("Antes:"):
                try:
                    antes = int(line.split("Antes:")[1].strip())
                except ValueError:
                    antes = None
            elif line.startswith("Atual:"):
                try:
                    atual = int(line.split("Atual:")[1].strip())
                except ValueError:
                    atual = None

        if razao_social and cpf_cnpj and antes is not None and atual is not None:
            # Definir se houve Aumento ou Redução
            if atual > antes:
                alteracao = "Aumento"
            elif atual < antes:
                alteracao = "Redução"
            else:
                alteracao = "Sem Alteração"
            
            # Adicionar os dados à lista, incluindo a data de recebimento no formato brasileiro
            data.append({
                "Razão Social": razao_social,
                "CPF/CNPJ": cpf_cnpj,
                "Antes": antes,
                "Atual": atual,
                "Alteração": alteracao,
                "Data Recebimento": message.ReceivedTime.date().strftime("%d/%m/%Y")  # Formato brasileiro
            })

# Converter a lista de dados em um DataFrame e salvar como Excel
if data:
    # Definir o nome do arquivo com a data atual no formato brasileiro
    file_name = f"C:\\Users\\7932\\Desktop\\CreditHub\\dados_protestos_{today.strftime('%d-%m-%Y')}.xlsx"
    df = pd.DataFrame(data)
    df.to_excel(file_name, index=False)
    print(f"Arquivo salvo com sucesso: {file_name}")
else:
    print("Nenhum dado válido encontrado para salvar.")


import pandas as pd
import win32com.client
import re
from datetime import datetime

# Conectar ao Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 é o código para "Inbox" ou "Caixa de Entrada"

# Filtros de assunto e data
subject_filter_1 = "AUMENTO EM PROCESSOS JUDICIAIS"
subject_filter_2 = "REDUÇÃO DE PROCESSOS JUDICIAIS DETECTADA!"
today = datetime.now().date()

# Lista para armazenar os dados
data = []

for message in inbox.Items:
    if (subject_filter_1 in message.Subject or subject_filter_2 in message.Subject) and message.ReceivedTime.date() == today:
        body = message.Body
        
        # Usar regex para encontrar as informações no corpo do e-mail
        razao_social = re.search(r"Nome/Razão Social:\s*(.+)", body)
        cpf_cnpj = re.search(r"CPF/CNPJ:\s*(\d+)", body)
        antes = re.search(r"Antes:\s*(\d+)", body)
        atual = re.search(r"Atual:\s*(\d+)", body)
        
        # Extrair os valores encontrados
        if razao_social and cpf_cnpj and antes and atual:
            antes_value = int(antes.group(1).strip())
            atual_value = int(atual.group(1).strip())
            
            # Definir se houve Aumento ou Redução
            if atual_value > antes_value:
                alteracao = "Aumento"
            elif atual_value < antes_value:
                alteracao = "Redução"
            else:
                alteracao = "Sem Alteração"
            
            # Adicionar os dados à lista, incluindo a data de recebimento no formato brasileiro
            data.append({
                "Razão Social": razao_social.group(1).strip(),
                "CPF/CNPJ": cpf_cnpj.group(1).strip(),
                "Antes": antes_value,
                "Atual": atual_value,
                "Alteração": alteracao,
                "Data Recebimento": message.ReceivedTime.date().strftime("%d/%m/%Y")  # Formatando a data para o formato brasileiro
            })
        else:
            print("Nem todos os dados foram encontrados para este e-mail.")

# Converter a lista de dados em um DataFrame e salvar como Excel
if data:
    # Definir o nome do arquivo com a data atual no formato brasileiro
    file_name = f"C:\\Users\\7932\\Desktop\\CreditHub\\dados_processos_judiciais_{today.strftime('%d-%m-%Y')}.xlsx"
    df = pd.DataFrame(data)
    df.to_excel(file_name, index=False)
    print(f"Arquivo salvo com sucesso: {file_name}")
else:
    print("Nenhum dado válido encontrado para salvar.")