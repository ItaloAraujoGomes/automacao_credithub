import os
import pandas as pd
import win32com.client
import re
from datetime import datetime, timedelta

# Função para processar e-mails
def process_emails(subject_filters, file_name_prefix):
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
                    # Verificar se o assunto contém os filtros especificados
                    if any(filter in item.Subject for filter in subject_filters):
                        body = item.Body

                        # Usar regex para extrair informações do corpo do e-mail
                        razao_social = re.search(r"Nome/Razão Social:\s*(.+)", body)
                        cpf_cnpj = re.search(r"CPF/CNPJ:\s*(\d+)", body)
                        antes = re.search(r"Antes:\s*(\d+)", body)
                        atual = re.search(r"Atual:\s*(\d+)", body)

                        # Validar os dados extraídos
                        if razao_social and cpf_cnpj and antes and atual:
                            antes_value = int(antes.group(1).strip())
                            atual_value = int(atual.group(1).strip())

                            # Determinar o tipo de alteração
                            if atual_value > antes_value:
                                alteracao = "Aumento"
                            elif atual_value < antes_value:
                                alteracao = "Redução"
                            else:
                                alteracao = "Sem Alteração"

                            # Adicionar dados à lista
                            data.append({
                                "Razão Social": razao_social.group(1).strip(),
                                "CPF/CNPJ": cpf_cnpj.group(1).strip(),
                                "Antes": antes_value,
                                "Atual": atual_value,
                                "Alteração": alteracao,
                                "Data Recebimento": received_time.strftime("%d/%m/%Y %H:%M:%S")  # Formato brasileiro
                            })
            except AttributeError as e:
                print(f"Erro ao processar item: {e}")

    # Salvar os dados em um arquivo Excel
    if data:
        # Verificar se o diretório existe, criar se não
        directory = r"C:\Users\7932\Desktop\CreditHub"
        if not os.path.exists(directory):
            os.makedirs(directory)

        # Nome do arquivo com data atual no formato brasileiro
        file_name = f"{directory}\\{file_name_prefix}_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
        df = pd.DataFrame(data)
        df.to_excel(file_name, index=False)
        print(f"Arquivo salvo com sucesso: {file_name}")
    else:
        print("Nenhum dado válido encontrado para salvar.")

# Filtros e prefixos para os dois tipos de e-mails
filters_protestos = ["Aumento em Protesto Detectado!", "Redução de Protestos Detectada!"]
filters_processos = ["AUMENTO EM PROCESSOS JUDICIAIS", "REDUÇÃO DE PROCESSOS JUDICIAIS DETECTADA!"]

# Processar e-mails
process_emails(filters_protestos, "dados_protestos")
process_emails(filters_processos, "dados_processos_judiciais")
