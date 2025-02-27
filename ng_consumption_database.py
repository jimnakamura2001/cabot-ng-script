

import datetime
import os
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from cryptography.fernet import Fernet
import json
import requests

# Função para criar ou adicionar informações ao arquivo de log
def criar_log(mensagem):
    # Obtém o diretório atual do script
    # diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    # diretorio_atual = r'\\maua-ntf01\Common\BD\ScriptComgas\log'
    diretorio_atual = "./log"

    # Obtém a data e hora atual
    data_hora_atual = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Define o nome do arquivo de log com base na data atual
    nome_arquivo_log = "log_" + datetime.datetime.now().strftime("%Y-%m-%d") + ".txt"

    # Caminho completo para o arquivo de log
    caminho_arquivo_log = os.path.join(diretorio_atual, nome_arquivo_log)

    # Abre o arquivo de log em modo de escrita (append)
    with open(caminho_arquivo_log, "a") as arquivo_log:
        # Escreve a mensagem de log com data e hora
        arquivo_log.write(f"[{data_hora_atual}] {mensagem}\n")
        
    print(mensagem)
    # print("\n")
    
# caminho_bd_excel = r'\\maua-ntf01\Common\BD\ScriptComgas\bd_script_ng.xlsx'

#Declarando caminho para o sharepoint e autenticando

pastagn_url = "https://cabotcorp.sharepoint.com/sites/MauaProd"
pastagn_relative_url = "/sites/MauaProd/Shared%20Documents/General/PRODU%C3%87%C3%83O/Planejamento/Planejamento%20de%20produ%C3%A7%C3%A3o/ScriptGN"
file_name = "ng_consuption_prediction_yest_info.txt"

# Leitura da chave
with open('key.key', 'rb') as key_file:
    key = key_file.read()

cipher = Fernet(key)

# Leitura do arquivo de configuração encriptado
with open('config.enc', 'rb') as config_file:
    encrypted_data = config_file.read()

# Desencriptação dos dados
config_data = cipher.decrypt(encrypted_data).decode()

# Extração das credenciais
# The code is splitting the `config_data` string into a list of lines using the newline character `\n`
# as the delimiter. Each element in the `config_lines` list will represent a line from the original
# `config_data` string.
config_lines = config_data.split('\n')

# The above code is creating a dictionary `config_dict` by splitting each line in the `config_lines`
# list by the '=' character and using the first part as the key and the second part as the value. It
# is also filtering out any empty lines in the `config_lines` list.
config_dict = {line.split('=')[0]: line.split('=')[1] for line in config_lines if line}

username = config_dict['USERNAME']
password = config_dict['PASSWORD']

# URL do site do SharePoint
site_url = config_dict['SITE_URL']

# Caminho do arquivo no SharePoint
relative_url = config_dict['RELATIVE_URL']

# Autenticação
# The above code is attempting to authenticate a user with a given username and password to access a
# site URL using the Microsoft SharePoint Online Client Object Model in Python. It first creates an
# AuthenticationContext object with the site URL, then tries to acquire a token for the user with the
# provided credentials. If the token acquisition is successful, it creates a ClientContext object for
# the site URL using the acquired authentication token. If the token acquisition fails, it prints out
# the last error message from the AuthenticationContext object.
ctx_auth = AuthenticationContext(site_url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(site_url, ctx_auth)
else:
    print(ctx_auth.get_last_error())


ontem_nao_teve_troca=False
ontem_Total_consumoDiario_m3_rounded=0
ontem_Total_consumo_dia_comgas_rounded=0
ontem_ng_yest_consumed_value=0
poder_calorifico = 0
poder_calorifico_backup = 0
input_manual = False
acertou_ontem = False


ctx_auth = AuthenticationContext(pastagn_url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(pastagn_url, ctx_auth)
else:
    print(ctx_auth.get_last_error())

# Caminho do arquivo no SharePoint
file_url = pastagn_relative_url + "/ng_consuption_prediction_yest_info.txt"
# Obter o arquivo do SharePoint
file_yest = ctx.web.get_file_by_server_relative_url(file_url)

local_path = "ng_consuption_prediction_yest_info.txt"
with open(local_path, "wb") as local_file:
    file_yest.download(local_file)
    ctx.execute_query()
print(f"Arquivo salvo localmente em: {local_path}")

# Lendo o conteúdo do stream
# Processar o arquivo localmente
with open(local_path, "r", encoding="utf-8") as arquivo:
    for linha in arquivo:
        exec(linha)  # Executar cada linha do arquivo

print(ontem_nao_teve_troca)
print(ontem_Total_consumoDiario_m3_rounded)
print(ontem_Total_consumo_dia_comgas_rounded)
print(ontem_ng_yest_consumed_value)
print(poder_calorifico)
print(poder_calorifico_backup)
print(input_manual)
print(acertou_ontem)


criar_log("Iniciando processo de inserir valor no site \n")

# The above Python code is making a POST request to the specified URL
# "https://industrial.comgas.com.br/api/PI/login-industrial" with a JSON payload containing
# certain data such as "documento", "codUsuario", "uniqueId", etc. It sets specific headers
# for the request like 'accept', 'content-type', 'user-agent', etc.
url = "https://industrial.comgas.com.br/api/PI/login-industrial"

payload = json.dumps({
"documento": "61741690000124",
"codUsuario": "000003751112",
"email": "",
"senha": "",
"segmento": "",
"uniqueId": "NjE3NDE2OTAwMDAxMjQxNzQwMTU5NDExOTc2"
})
headers = {
'accept': 'application/json, text/plain, */*',
'accept-language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7,pt-PT;q=0.6',
'content-type': 'application/json',
'origin': 'https://industrial.comgas.com.br',
'priority': 'u=1, i',
'sec-ch-ua': '"Not(A:Brand";v="99", "Microsoft Edge";v="133", "Chromium";v="133"',
'sec-ch-ua-mobile': '?0',
'sec-ch-ua-platform': '"Windows"',
'sec-fetch-dest': 'empty',
'sec-fetch-mode': 'cors',
'sec-fetch-site': 'same-origin',
'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36 Edg/133.0.0.0',
'x-msg-timestamp': '1740159411977'
}

# The above Python code is making a POST request to a specified URL with headers and payload
# data. It then checks if the response status code is 202 (indicating a successful POST
# request). If the status code is 202, it converts the response to JSON format and extracts the
# "jwt" token from the response data.
response = requests.request("POST", url, headers=headers, data=payload)

# Verifica se a resposta foi bem-sucedida (código 202 para POST)
if response.status_code == 202:
    data = response.json()  # Converte a resposta para JSON
    auth_token = data.get("jwt")  # Obtém o token 'jwt'

    if auth_token:
        # Simulando o armazenamento da variável (substitua conforme necessário)
        collection_variables = {"Auth-industrial": auth_token}
        # print("Token armazenado:", collection_variables["Auth-industrial"])
        criar_log("Auth-industrial armazenado com sucesso.")
    else:
        criar_log("Erro: Chave 'jwt' não encontrada na resposta.")
else:
    criar_log(f"Erro na requisição: {response.status_code}, {response.text}")
        
# CÓDIGO PARA ACESSAR A API COM O TOKEN GERADO

# The above Python code snippet is setting up a request to a specific URL endpoint with the
# given payload and headers. The URL being accessed is
# "https://industrial.comgas.com.br/api/PI/login?uniqueId=NjE3NDE2OTAwMDAxMjQxNzQwMTU5NDExOTc2&uniqueId=NjE3NDE2OTAwMDAxMjQxNzQwMTU5NDExOTc2".
url = "https://industrial.comgas.com.br/api/PI/login?uniqueId=NjE3NDE2OTAwMDAxMjQxNzQwMTU5NDExOTc2&uniqueId=NjE3NDE2OTAwMDAxMjQxNzQwMTU5NDExOTc2"

payload = {}
headers = {
'accept': 'application/json, text/plain, */*',
'accept-language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7,pt-PT;q=0.6',
'priority': 'u=1, i',
'authorization': f'ComgasToken {collection_variables["Auth-industrial"]}',
'sec-ch-ua': '"Not(A:Brand";v="99", "Microsoft Edge";v="133", "Chromium";v="133"',
'sec-ch-ua-mobile': '?0',
'sec-ch-ua-platform': '"Windows"',
'sec-fetch-dest': 'empty',
'sec-fetch-mode': 'cors',
'sec-fetch-site': 'same-origin',
'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36 Edg/133.0.0.0',
'x-msg-timestamp': '1740159412388'
}

# The above code is making a GET request to a specified URL with custom headers and payload
# data. It then checks if the response status code is 200 (indicating a successful request),
# converts the response to JSON format, and extracts the "jwt" token from the response data.
# If the token is found, it stores it in a collection variable and prints a message confirming
# the token storage. If the token is not found in the response, it prints an error message. If
# the response status code is not 200, it prints an error message with the status code and
# response text.
response = requests.request("GET", url, headers=headers, data=payload)

# Verifica se a resposta foi bem-sucedida (código 200 para GET)
if response.status_code == 200:
    data = response.json()  # Converte a resposta para JSON
    auth_token = data.get("jwt")  # Obtém o token 'jwt'

    if auth_token:
        # Simulando o armazenamento da variável (substitua conforme necessário)
        collection_variables = {"Auth-consumo": auth_token}
        # print("Token armazenado:", collection_variables["Auth-consumo"])
        criar_log("Auth-consumo armazenado com sucesso. \n")
    else:
        criar_log("Erro: Chave 'jwt' não encontrada na resposta.")
else:
    criar_log(f"Erro na requisição: {response.status_code}, {response.text}")
    
# CÓDIGO PARA ACESSAR O CONSUMO DO DIA DE ONTEM
    
url = "https://industrial.comgas.com.br/api/portal-industrial/consumo-cliente"

payload = ""
headers = {
'accept': 'application/json, text/plain, */*',
'accept-language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7,pt-PT;q=0.6',
'authorization': f'ComgasToken {collection_variables["Auth-consumo"]}',
'priority': 'u=1, i',
'sec-ch-ua': '"Not(A:Brand";v="99", "Microsoft Edge";v="133", "Chromium";v="133"',
'sec-ch-ua-mobile': '?0',
'sec-ch-ua-platform': '"Windows"',
'sec-fetch-dest': 'empty',
'sec-fetch-mode': 'cors',
'sec-fetch-site': 'same-origin',
'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36 Edg/133.0.0.0'
}

# The above Python code is making a GET request to a specified URL with headers and payload. It
# then checks if the response status code is 200 (indicating a successful response). If the
# response is successful, it converts the response data to a dictionary and retrieves the value
# associated with the key 'consumoProgramadoHoje' from the JSON data. This value is stored in the
# variable `ng_yest_consumed_value` and printed out. If the response status code is not 200, it
# prints an error message with the status code and response text.
ng_yest_consumed_value = 0

response_yesterday = requests.request("GET", url, headers=headers, data=payload)

# Verifique se a resposta foi bem-sucedida antes de acessar os dados
if response_yesterday.status_code == 200:
    data = response_yesterday.json()  # Converte para um dicionário
    ng_yest_consumed_value = data.get('consumoProgramadoHoje', "Chave não encontrada")
else:
    criar_log(f"Erro na requisição: {response_yesterday.status_code}, {response_yesterday.text}")
criar_log(f"Valor calculado pelo script: {ontem_Total_consumo_dia_comgas_rounded}")
criar_log(f"Valor inserido no portal: {ng_yest_consumed_value}")
aprox_m3_result = 0
if int(ng_yest_consumed_value) == int(ontem_Total_consumo_dia_comgas_rounded):
    criar_log("Os valores estão correspondentes, finalizando o código...")
    input_manual = False
    aprox_m3_result = ontem_Total_consumoDiario_m3_rounded
else:
    criar_log("A previsão de NG sofreu um input manual. \n")
    novo_poder_calorifico = int(ng_yest_consumed_value)/float(poder_calorifico)
            
    aprox_m3_result = int(novo_poder_calorifico)
    
    input_manual = True
    
    criar_log("O valor em m3 calculado foi de aproximadamente: ")
    criar_log(aprox_m3_result)
    criar_log("\n")          
    
    file_url = f"{pastagn_relative_url}/{file_name}"
    local_path = file_name # Caminho local para download temporário do arquivo

file_write_yest = ctx.web.get_file_by_server_relative_url(file_url)
with open(local_path, "wb") as local_file:
    file_write_yest.download(local_file)
    ctx.execute_query()
print(f"Arquivo baixado para edição local: {local_path}")
    
# Editar o arquivo localmente
with open(local_path, "w", encoding="utf-8") as arquivo:
    arquivo.write(f'ontem_nao_teve_troca={ontem_nao_teve_troca}\n')
    arquivo.write(f'ontem_Total_consumoDiario_m3_rounded={aprox_m3_result}\n')
    arquivo.write(f'ontem_Total_consumo_dia_comgas_rounded={ng_yest_consumed_value}\n')
    arquivo.write(f'ontem_ng_yest_consumed_value={ontem_ng_yest_consumed_value}\n')
    arquivo.write(f'poder_calorifico={poder_calorifico}\n')
    arquivo.write(f'poder_calorifico_backup={poder_calorifico_backup}\n')
    arquivo.write(f'input_manual={input_manual}\n')
    arquivo.write(f'acertou_ontem={acertou_ontem}\n')
print(f"Arquivo editado localmente: {local_path}")

# Fazer upload do arquivo atualizado de volta para o SharePoint
with open(local_path, "rb") as updated_file:
    target_folder = ctx.web.get_folder_by_server_relative_url(pastagn_relative_url)
    target_file = target_folder.upload_file(file_name, updated_file.read())
    ctx.execute_query()
print(f"Arquivo atualizado com sucesso no SharePoint: {file_url}")
    
print("Arquivos atualizados com sucesso no documento, fechando navegador. \n")


