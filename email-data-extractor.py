import win32com.client
import re
import pandas as pd

# Conectar ao Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Caixa de Entrada
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)  # Ordena por data (mais recentes primeiro)

# Regex para CNPJ e valores
regex_cnpj = r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}"
regex_valor = r"\d{1,3}(?:\.\d{3})*,\d{2}"  # Ex.: 1.234,56

dados = []

# Iterar pelos e-mails (ex.: últimos 20)
for message in list(messages)[:20]:
    if message.Class == 43:  # Verifica se é e-mail
        corpo = message.Body
        cnpjs = re.findall(regex_cnpj, corpo)
        valores = re.findall(regex_valor, corpo)

        # Assumindo que cada CNPJ corresponde a um valor (mesma ordem)
        for i in range(min(len(cnpjs), len(valores))):
            dados.append({"CNPJ": cnpjs[i], "Valor": valores[i], "Assunto": message.Subject})

# Salvar em Excel provisório
df = pd.DataFrame(dados)
df.to_excel("dados_extraidos.xlsx", index=False)

print("Extração concluída! Arquivo salvo como dados_extraidos.xlsx")
