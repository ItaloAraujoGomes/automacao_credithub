import os
import pandas as pd
import win32com.client
import re
from datetime import datetime, timedelta

# Caminho base para salvar os arquivos
BASE_DIRECTORY = r"C:\Users\7932\Desktop\CreditHub"

# Função para limpar razão social
def clean_razao_social(raw_name):
    return re.sub(r"CPF", "", raw_name.strip())

# Função para converter valores extraídos em inteiros
def extract_int(value):
    try:
        return int(value.strip())
    except ValueError:
        return None

# Função para processar e-mails
def process_emails(subject_filters, file_name, folder_name):
    # Conectar ao Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Código para "Inbox" ou "Caixa de Entrada"

    # Lista para armazenar os dados
    data = []

    # Calcular intervalo de 24 horas
    time_threshold = datetime.now() - timedelta(days=1)

    for item in inbox.Items:
        # Verificar se o item é um e-mail
        if item.Class == 43:  # 43 é o código de `MailItem`
            try:
                # Remover informações de fuso horário de ReceivedTime
                received_time = item.ReceivedTime.replace(tzinfo=None)

                # Verificar se o e-mail está dentro das últimas 24 horas
                if received_time >= time_threshold:
                    # Verificar se o e-mail contém os filtros especificados no assunto
                    if any(re.search(rf"{filter}", item.Subject, re.IGNORECASE) for filter in subject_filters):
                        body = item.Body

                        # Debug: Imprimir o assunto do e-mail
                        print(f"Assunto do e-mail: {item.Subject}")

                        # Usar regex para extrair as informações do corpo do e-mail
                        razao_social = re.search(r"Nome/Razão Social:\s*([\w\s\.\-]+)", body)
                        cpf_cnpj = re.search(r"(\d{14})", body)  # CPF/CNPJ com 14 dígitos
                        antes = re.search(r"Antes:\s*(\d+)", body)
                        atual = re.search(r"Atual:\s*(\d+)", body)

                        # Validar os dados extraídos
                        if razao_social and cpf_cnpj and antes and atual:
                            antes_value = extract_int(antes.group(1))
                            atual_value = extract_int(atual.group(1))

                            # Determinar o tipo de alteração
                            if atual_value > antes_value:
                                alteracao = "Aumento"
                            elif atual_value < antes_value:
                                alteracao = "Redução"
                            else:
                                alteracao = "Inclusão"

                            # Limpar a palavra "CPF" da razão social
                            razao_social_clean = clean_razao_social(razao_social.group(1))

                            # Adicionar dados à lista
                            data.append({
                                "Razão Social": razao_social_clean,
                                "CPF/CNPJ": cpf_cnpj.group(1).strip(),
                                "Antes": antes_value,
                                "Atual": atual_value,
                                "Alteração": alteracao,
                                "Data Recebimento": received_time.strftime("%d/%m/%Y %H:%M:%S")  # Formato brasileiro
                            })
            except Exception as e:
                print(f"Erro ao processar e-mail com assunto '{item.Subject}': {e}")

    # Verificar e criar o diretório para a pasta correspondente se não existir
    folder_path = os.path.join(BASE_DIRECTORY, folder_name)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # Salvar os dados em um arquivo Excel
    if data:
        # Nome do arquivo com data atual no formato brasileiro
        file_name_path = f"{folder_path}\\{file_name}_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
        df = pd.DataFrame(data)
        if not df.empty:
            df.to_excel(file_name_path, index=False)
            print(f"Arquivo salvo com sucesso: {file_name_path}")
        else:
            print(f"Nenhum dado válido encontrado para salvar no arquivo {file_name_path}.")
    else:
        print(f"Nenhum dado encontrado para os filtros especificados.")

# Filtros e prefixos para os e-mails
filters_inclusao_ccf = ["Inclusão de CCF Detectada!"]
filters_aumento_reducao_ccf = ["Aumento de CCF Detectada!", "Redução de CCF Detectada!"]
filters_aumento_reducao_protestos = ["Aumento em Protesto Detectado!", "Redução de Protestos Detectada!"]
filters_aumento_reducao_processos = ["AUMENTO EM PROCESSOS JUDICIAIS", "REDUÇÃO DE PROCESSOS JUDICIAIS DETECTADA!"]

# Processar e-mails de Inclusão de CCF
process_emails(filters_inclusao_ccf, "inclusao_ccf", "Inclusao_CCF")

# Processar e-mails de Aumento e Redução de CCF
process_emails(filters_aumento_reducao_ccf, "ccf", "CCF")

# Processar e-mails de Aumento e Redução de Protestos
process_emails(filters_aumento_reducao_protestos, "protestos", "Protestos")

# Processar e-mails de Aumento e Redução de Processos
process_emails(filters_aumento_reducao_processos, "processos", "Processos")
