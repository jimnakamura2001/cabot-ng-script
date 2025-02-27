import os
import datetime
import tagreader
import pandas as pd
import traceback
import time
import sys
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from cryptography.fernet import Fernet
import requests
import json

# Função para criar ou adicionar informações ao arquivo de log
def criar_log(mensagem):
    # Obtém o diretório atual do script
    # diretorio_atual = os.path.dirname(os.path.abspath(__file__))
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

# Calcular diferença percentual entre dois valores
def dentro_da_margem(a, b, margem_percentual=10):
    # Calcula a diferença percentual entre A e B
    diferenca_percentual = abs((int(a) - int(b)) / int(a)) * 100

    # Verifica se a diferença percentual está dentro da margem
    return diferenca_percentual <= margem_percentual

#Converte uma string para datetime
def converter_para_datetime(data_str):
    # Obtém o ano atual
    ano_atual = datetime.datetime.now().year

    # Adiciona o ano atual à string de data
    data_com_ano = f"{data_str}-{ano_atual}"
    
    # Retorna o objeto datetime
    return datetime.datetime.strptime(data_com_ano, '%d-%b-%Y')

# Cores de texto ANSI
# The above Python code defines escape sequences for text color and formatting in the terminal. Each
# variable represents a different color or style, such as red, green, yellow, blue, magenta, cyan,
# white, bold, and underline. These escape sequences can be used to change the color or style of text
# output in the terminal.
RED = '\033[91m'
GREEN = '\033[92m'
YELLOW = '\033[93m'
BLUE = '\033[94m'
MAGENTA = '\033[95m'
CYAN = '\033[96m'
WHITE = '\033[97m'
BOLD = '\033[1m'
UNDERLINE = '\033[4m'
END = '\033[0m'

# # Usando as cores
# print(f"{BOLD}{RED}Esta linha é muito importante!{END}")
# print(f"{BOLD}{BLUE}Esta linha também é importante, mas menos dramática!{END}")
# print(f"{GREEN}Esta linha é apenas uma informação adicional.{END}")

#Declarando caminho para o sharepoint e autenticando

pastagn_url = "https://cabotcorp.sharepoint.com/sites/MauaProd"
pastagn_relative_url = "/sites/MauaProd/Shared%20Documents/General/PRODU%C3%87%C3%83O/Planejamento/Planejamento%20de%20produ%C3%A7%C3%A3o/ScriptGN"

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
print(relative_url)

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

# Obter o arquivo
# The above code is using the SharePoint client-side object model in Python to get a file by its
# server-relative URL. It loads the file object and then executes the query to retrieve the file from
# the SharePoint server.
file = ctx.web.get_file_by_server_relative_url(relative_url)
ctx.load(file)
ctx.execute_query()

# Obter a data de modificação do arquivo
last_modified = file.properties['TimeLastModified']
print(f"Data de modificação do arquivo: {last_modified} \n")

arquivo_atualizado = False

# The above Python code snippet is checking if the date of the last modification of a file
# (`last_modified`) is not equal to the current date. 
if last_modified.date() != datetime.datetime.today().date():
    criar_log(f"{BOLD}{RED}O arquivo de troca de graus não está atualizado. {END}")
    arquivo_atualizado = False
    criar_log(arquivo_atualizado)
    print("\n")
    
else:
    # Obtendo o arquivo pelo caminho relativo no SharePoint
    file = ctx.web.get_file_by_server_relative_url(relative_url)
    ctx.load(file)
    ctx.execute_query()

    # Lendo os dados binários do arquivo
    file_data = file.read()

    # Salvando o arquivo localmente
    with open("localfile.csv", "wb") as local_file:
        local_file.write(file_data)

    print("File downloaded successfully")  
    print("A data é hoje.\n")
    arquivo_atualizado = True
    print(arquivo_atualizado)
    
nao_teve_troca = False

if __name__ == "__main__":
    
    # Chama a função para criar o log com uma mensagem
    criar_log("Script iniciado com sucesso! \n")
    
    try:
        
# This Python code snippet is reading information from a file specified by the variable
# `caminho_info_ontem` and then executing each line of the file as Python code using the `exec()`
# function. The code then logs various variables and their values using a function called
# `criar_log()` and prints a newline character.
        ontem_nao_teve_troca=False
        ontem_Total_consumoDiario_m3_rounded=0
        ontem_Total_consumo_dia_comgas_rounded=0
        ontem_ng_yest_consumed_value=0
        poder_calorifico = 0
        poder_calorifico_backup = 0
        input_manual = False
        acertou_ontem = False

        # caminho_info_ontem = "./ng_consuption_prediction_yest_info.txt"
        
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

        criar_log(f"Ontem não teve troca: {ontem_nao_teve_troca}")
        criar_log(f"Consumo diário sem o poder calorífico de ontem: {ontem_Total_consumoDiario_m3_rounded}")
        criar_log(f"Consumo diário com o poder calorífico de ontem: {ontem_Total_consumo_dia_comgas_rounded}")
        criar_log(ontem_ng_yest_consumed_value)
        criar_log(f"Poder calorífico: {poder_calorifico}")
        criar_log(f"Poder calorífico de backup: {poder_calorifico_backup}")
        criar_log(f"Houve input manual ontem: {input_manual}")
        criar_log(f"O script acertou ontem: {acertou_ontem}")
        print("\n")
            
        # Conexão com banco de dados Aspen
        c = tagreader.IMSClient(datasource="maua-ntp01",
                                imstype="aspenone",
                                tz="Brazil/East",
                                url="http://maua-ntp01/ProcessData/AtProcessDataREST.dll")
        c.connect()
        
        # Tags que precisam ser lidas
        tags = ["1.BURN1.NG.FC.01.PV", 
                "2.BURN1.NG.FC.01.PV", 
                "2.REAC1.NG.FC.01.PV", 
                "3.BURN1.NG.FC.01.PV",
                ]
        
        tags_status = ["1.REAC1.STATUS",
                       "2.REAC1.STATUS",
                       "3.REAC1.STATUS",
                       "1.REAC1.DCS.GRADE",
                       "2.REAC1.DCS.GRADE",
                       "3.REAC1.DCS.GRADE"]

        # Horario de início
        hStart_raw = datetime.datetime.now()
        hStart_midnight = hStart_raw.replace(hour=1, minute=0, second=0, microsecond=0)

        # Minutos alterados para 9h30
        hStart_status_raw = hStart_raw.replace(hour=9, minute=0, second=0, microsecond=0)

        # The code snippet is converting a datetime object `hStart_midnight` to a string format with
        # the day, month, year, hour, minute, and second components formatted as "dd.mm.YYYY
        # HH:MM:SS".
        hStart = hStart_midnight.strftime("%d.%m.%Y %H:%M:%S")
        
        hStart_status = hStart_status_raw.strftime("%d.%m.%Y %H:%M:%S")

        # Horário de fim
        hEnd_raw = datetime.datetime.now()
        hEnd_midnight = hEnd_raw.replace(hour=10, minute=0, second=0, microsecond=0)

        hEnd = hEnd_midnight.strftime("%d.%m.%Y %H:%M:%S")

        # Intervalo de tempo
        interval = 3600

        # Leitura dos dados
        df = c.read(tags, hStart, hEnd, interval)

        # Cálculo para obter Nm³/h (Cabot) Current
        #Horário de agora
        # hNow_raw = hStart_status_raw
        hNow_raw = datetime.datetime.now()
        hNow = hNow_raw.strftime("%d.%m.%Y %H:%M:%S")

        dfnm3Current = []

        for i in range(4):
            dfCurrent = c.read(tags[i], None, hNow, read_type=tagreader.ReaderType.SNAPSHOT)
            dfnm3Current.append(dfCurrent[tags[i]].mean())
        
        df_reac_status = []
        
        for i in range(0, 6):
            # The code snippet is reading data from a DataFrame `df_reacst` using the `c.read` method
            # with specific parameters such as `tags_status[i]`, `hStart_status`, `hEnd`, `interval`,
            # and `read_type=tagreader.ReaderType.INTERPOLATED`. The data is being read from the
            # specified tags within the given time interval and using interpolation for missing
            # values.
            df_reacst = c.read(tags_status[i], hStart_status, hEnd, interval, read_type=tagreader.ReaderType.INTERPOLATED)
            # print(type(df_reacst))
            df_reac_status.append(df_reacst[tags_status[i]].mean())
            # print(df_reacst)
        
        df_reac_status = list(map(int, df_reac_status))

        # print(df_reac_status)
        
        # The code is creating a new DataFrame `df_reac_grades` by selecting rows starting from the
        # 4th row (index 3) of the DataFrame `df_reac_status`.
        df_reac_grades = df_reac_status[3:]
        df_reac_STATUS = df_reac_status[:3]
        
        # print("Últimos 3 itens:", df_reac_grades)
        # print("Primeiros 3 itens:", df_reac_STATUS, "\n")
                
        # The code is reading a CSV file named 'unit_status.csv' into a pandas DataFrame called
        # 'mapping_df'.
        mapping_df = pd.read_csv('unit_status.csv')

        # Criar um dicionário de mapeamento a partir do DataFrame
        status_mapping = pd.Series(mapping_df.status_text.values, index=mapping_df.status_id).to_dict()
        
        # The code is creating a list called `status_text_list` by mapping each value in the
        # `df_reac_STATUS` column to its corresponding value in the `status_mapping` dictionary.
        status_text_list = [status_mapping[status] for status in df_reac_STATUS]
        
        print(status_text_list)
        
        #Convertendo o valor dos graus para nomes
        
        df_grade_naming = pd.read_csv('grade_codes.csv')
        
        # Criar dicionários de mapeamento para cada par de colunas de código e nome
        # The above code snippet in Python is creating a dictionary `ma1_mapping` using pandas
        # library. It is taking values from the 'ma1name' column of the DataFrame `df_grade_naming`
        # and using 'ma1code' column as the index. The `pd.Series` function is used to create a Series
        # object, and then `to_dict()` method is called to convert this Series object into a
        # dictionary. The resulting dictionary will have 'ma1code' as keys and 'ma1name' as values.
        ma1_mapping = pd.Series(df_grade_naming.ma1name.values, index=df_grade_naming.ma1code).to_dict()
        ma2_mapping = pd.Series(df_grade_naming.ma2name.values, index=df_grade_naming.ma2code).to_dict()
        ma3_mapping = pd.Series(df_grade_naming.ma3name.values, index=df_grade_naming.ma3code).to_dict()
     
        # Converter cada item da lista de acordo com os dicionários de mapeamento
        converted_grades = [
            ma1_mapping.get(df_reac_grades[0], "Unknown"),
            ma2_mapping.get(df_reac_grades[1], "Unknown"),
            ma3_mapping.get(df_reac_grades[2], "Unknown")
        ]

        # Exibir a lista resultante
        print(converted_grades)            
           
        # Cálculo para obter Nm³/h (Cabot) Passed
        dfSum = [df["1.BURN1.NG.FC.01.PV"].sum(),
                df["2.BURN1.NG.FC.01.PV"].sum(),
                df["2.REAC1.NG.FC.01.PV"].sum(), 
                df["3.BURN1.NG.FC.01.PV"].sum()]

        # Cálculo para obter NG até a troca
        # This Python code is calculating the time remaining in the day. It first extracts the current hour
        # and minute from a given time `hNow_raw`. Then, it converts the current time to minutes and
        # calculates the hours and minutes that have already passed in the day. Finally, it calculates the
        # hours remaining in the day by subtracting the hours passed from 24.
        hAtual = hNow_raw.hour
        mAtual = hNow_raw.minute

        mPassados = (hAtual * 60) + mAtual
        hPassados = mPassados // 60
        hRestantes = 24 - hPassados
        
        # The above code snippet is written in Python and performs the following operations:
        dfOperations = pd.DataFrame({"Nm³/h (Cabot) Passed": dfSum, "Nm³/h (Cabot) Current": dfnm3Current})

        rename_df = {0: '1.BURN1.NG.FC.01.PV', 
                    1: '2.BURN1.NG.FC.01.PV', 
                    2: '2.REAC1.NG.FC.01.PV', 
                    3: '3.BURN1.NG.FC.01.PV'}

        dfOperations = dfOperations.rename(index=rename_df)

        total_nm3hPassed = sum([dfSum[0], dfSum[1], dfSum[2], dfSum[3]])

        # Coletando as trocas de grau
        
        # Inicializar variáveis
        MA01_hora_troca = 0
        MA02_hora_troca = 0
        MA03_hora_troca = 0

        MA01_grau_troca = None
        MA02_grau_troca = None
        MA03_grau_troca = None

        arquivo_troca_grau = "localfile.csv"
        
        # Realiza o tratamento dos nomes das colunas para que seja possível trabalhar com elas
        #ADICIONAR TRUE ADICIONAR TRUE ADICIONAR TRUE ADICIONAR TRUE ADICIONAR TRUE ADICIONAR TRUE ADICIONAR TRUE ADICIONAR TRUE ADICIONAR TRUE ADICIONAR TRUE ADICIONAR TRUE  
        if arquivo_atualizado == True:
        # Realiza a leitura do arquivo aonde contém todas as trocas de grau
            try:
                df = pd.read_csv(arquivo_troca_grau, encoding='latin-1')
                
            except UnicodeDecodeError:
                criar_log("Não foi possível ler o arquivo CSV com a codificação Latin-1.")
            df = df.rename(columns={" REPLACE( CONCAT( LEFT(Time, 2), 'h'), ':', '')": 'Time_Formatted'})
            
            # print(df.columns)
            
            colunas_to_remove = ['ï»¿WD', 'Date_ValidaÃ§Ã£o', 'Locationdesc', 'Workorder Number', 'Quantity',
                                 'JDE Comment', 'Aspen Comment', ' CAST (Workorder Number AS INTEGER )']
            
            # The code is dropping columns specified in the list `colunas_to_remove` from the
            # DataFrame `df`.
            df = df.drop(columns=colunas_to_remove)
            
            df['Date'] = pd.to_datetime(df['Date'])
            
            data_hoje = datetime.datetime.today().date()
            
            criar_log(f"df.date: {df['Date'].dt.date}")
            criar_log(f"data_hoje: {data_hoje}")
            
            # Verificar se todas as datas na coluna "Date" são correspondentes ao dia de hoje
            # The code is checking if any date in the 'Date' column of the DataFrame `df` is equal to
            # the variable `data_hoje`. If any date matches, the condition will return `True`,
            # otherwise `False`.
            if arquivo_atualizado:
                print("O arquivo está atualizado. \n")

                # Filtrar as linhas onde Time_Formatted é diferente de "0h"
                linhas_troca = df[df['Time_Formatted'] != '0h']

                # Atribuir o horário da coluna "Time_Formatted" às variáveis baseadas na "Resource Description"
                for _, row in linhas_troca.iterrows():
                    hora_troca = int(row['Time_Formatted'].replace('h', ''))
                    material = row['Material'].split('_')[0]
                    
                    # The code is checking if the value in the 'Resource Description' column of a row
                    # is equal to "MA - UNIT 1".
                    if row['Resource Description'] == "MA - UNIT  1" and status_text_list[0] != "SHUTDOWN":
                        MA01_hora_troca = hora_troca
                        MA01_grau_troca = material
                        nao_teve_troca = False
                        criar_log("\n")
                        criar_log(f"{BOLD}{YELLOW}Na MA01 está planejado troca de grau{END}")
                        criar_log(f"Horário da troca: (h) {MA01_hora_troca}")
                        criar_log(f"Grau atual: {converted_grades[0]}")
                        criar_log(f"Grau da troca: {MA01_grau_troca}")
                        criar_log("\n")
                    elif row['Resource Description'] == "MA - UNIT  2" and status_text_list[1] != "SHUTDOWN":
                        MA02_hora_troca = hora_troca
                        MA02_grau_troca = material
                        nao_teve_troca = False
                        criar_log("\n")
                        criar_log(f"{BOLD}{YELLOW}Na MA02 está planejado troca de grau{END}")
                        criar_log(f"Horário da troca: (h) {MA02_hora_troca}")
                        criar_log(f"Grau atual: {converted_grades[1]}")
                        criar_log(f"Grau da troca: {MA02_grau_troca}")
                        criar_log("\n")
                    elif row['Resource Description'] == "MA - UNIT  3" and status_text_list[2] != "SHUTDOWN":
                        MA03_hora_troca = hora_troca
                        MA03_grau_troca = material
                        nao_teve_troca = False
                        criar_log("\n")
                        criar_log(f"{BOLD}{YELLOW}Na MA03 está planejado troca de grau{END}")
                        criar_log(f"Horário da troca: (h) {MA03_hora_troca}")
                        criar_log(f"Grau atual: {converted_grades[2]}")
                        criar_log(f"Grau da troca: {MA03_grau_troca}")
                        criar_log("\n")
                    else:
                        nao_teve_troca = True
                        criar_log("\n")            
                        criar_log(f"{BOLD}{RED}ATENÇÃO: NÃO ESTÁ PREVISTO NENHUMA TROCA DE GRAU PARA HOJE. {END}")            
                        criar_log("\n")  

                # Exibir os resultados
                criar_log(f"MA01_hora_troca: {MA01_hora_troca}, MA01_grau_troca: {MA01_grau_troca}")
                criar_log(f"MA02_hora_troca: {MA02_hora_troca}, MA02_grau_troca: {MA02_grau_troca}")
                criar_log(f"MA03_hora_troca: {MA03_hora_troca}, MA03_grau_troca: {MA03_grau_troca}")
                
            else:
                
                nao_teve_troca = True
                criar_log("\n")            
                criar_log(f"{BOLD}{RED}ATENÇÃO: NÃO ESTÁ PREVISTO NENHUMA TROCA DE GRAU PARA HOJE. {END}")            
                criar_log("\n")            
                
        # Realiza a leitura do arquivo aonde está a relação de consumo dos grades de forma manual
        arquivo_grades_consumo_url = pastagn_relative_url + "/grades_consumo.csv"

        # Obter o arquivo do SharePoint
        file_grade = ctx.web.get_file_by_server_relative_url(arquivo_grades_consumo_url)
        
        arquivo_grades_consumo = "grades_consumo.csv"
        with open(arquivo_grades_consumo, "wb") as local_file:
            file_grade.download(local_file)
            ctx.execute_query()
        print(f"Arquivo salvo localmente em: {arquivo_grades_consumo}")
        
        try:
            df_grades_consumo = pd.read_csv(arquivo_grades_consumo, encoding='latin-1')
        except UnicodeDecodeError:
            # print("Não foi possível ler o arquivo CSV com a codificação Latin-1.")
            criar_log("Não foi possível ler o arquivo CSV com a codificação Latin-1.")

        NG_ate_troca = []
        
        # Cálculo de NG até a troca
        
        # The code is checking if the value of the variable `MA01_hora_troca` is greater than the
        # value of the variable `hPassados`. If the condition is true, the code block following the
        # `if` statement will be executed.
        print(dfnm3Current)
        
        if MA01_hora_troca > hPassados:
            # The code is appending a tuple to the list `NG_ate_troca`. The tuple contains the result
            # of the expression `(MA01_hora_troca - hPassados) * dfnm3Current[0]`.
            NG_ate_troca.append((MA01_hora_troca - hPassados)*dfnm3Current[0])
        else:
            NG_ate_troca.append(0)
            
        if MA02_hora_troca > hPassados:
            NG_ate_troca.append((MA02_hora_troca - hPassados)*dfnm3Current[1])
            if converted_grades[1] == "SNS1":
                NG_ate_troca.append((MA02_hora_troca - hPassados)*dfnm3Current[2])
            else:
                NG_ate_troca.append(0)
        else: 
            NG_ate_troca.append(0)
            NG_ate_troca.append(0)
            
        if MA03_hora_troca > hPassados:
            NG_ate_troca.append((MA03_hora_troca - hPassados)*dfnm3Current[3])
        else:
            NG_ate_troca.append(0)
            
        # print(NG_ate_troca)
            
        dfOperations["NG até a troca"] = NG_ate_troca

        #NG do Grade após a troca
        NG_grade_apos_troca = []

        # This Python code snippet is checking if the time difference between the current time and the time of
        # a specific event (MA01, MA02, MA03) is less than 24 hours. If the condition is met, it retrieves a
        # consumption value from a DataFrame based on the grade associated with the event and appends it to a
        # list called `NG_grade_apos_troca`. If the condition is not met, it appends a value of 0 to the list.
        if (24 - MA01_hora_troca) < 24:
            consumo_grau = df_grades_consumo.loc[df_grades_consumo['Grades MA1'] == MA01_grau_troca]
            consumo_value =  consumo_grau.iat[0, 1]
            NG_grade_apos_troca.append(consumo_value)
        else:
            NG_grade_apos_troca.append(0)
            
        if (24 - MA02_hora_troca) < 24:
            consumo_grau = df_grades_consumo.loc[df_grades_consumo['Grades MA2'] == MA02_grau_troca]
            consumo_value =  consumo_grau.iat[0, 3]
            NG_grade_apos_troca.append(consumo_value)
            NG_grade_apos_troca.append(0)
        else:
            NG_grade_apos_troca.append(0)
            NG_grade_apos_troca.append(0)

        if (24 - MA03_hora_troca) < 24:
            consumo_grau = df_grades_consumo.loc[df_grades_consumo['Grade MA3'] == MA03_grau_troca]
            print(df_grades_consumo)
            print(MA03_grau_troca)
            consumo_value =  consumo_grau.iat[0, 5]
            NG_grade_apos_troca.append(consumo_value)
        else:
            NG_grade_apos_troca.append(0)
            
        dfOperations['NG do Grade após troca'] = NG_grade_apos_troca

        # Cálculo para o valor Comgas (Nm³) Estimado
        ComgasNm3Estimado = []

        # The code is checking if the result of subtracting the value of `MA01_hora_troca` from 24 is
        # equal to 24.
        if (24 - MA01_hora_troca) == 24:
            # The code is appending the result of multiplying the first element of the `dfnm3Current`
            # list by the value of `hRestantes` to the `ComgasNm3Estimado` list in Python.
            ComgasNm3Estimado.append(dfnm3Current[0] * hRestantes)
        else: 
            # The code snippet is appending a calculated value to the list `ComgasNm3Estimado`. The
            # calculated value is obtained by multiplying the first element of the list
            # `NG_grade_apos_troca` by the difference between 24 and the value of `MA01_hora_troca`,
            # and then adding the first element of the list `NG_ate_troca`.
            ComgasNm3Estimado.append((NG_grade_apos_troca[0] * (24 - MA01_hora_troca)) + NG_ate_troca[0])
            
        if (24 - MA02_hora_troca) == 24:
            ComgasNm3Estimado.append(dfnm3Current[1] * hRestantes)
            ComgasNm3Estimado.append(dfnm3Current[2] * hRestantes)
        else: 
            ComgasNm3Estimado.append((NG_grade_apos_troca[1] * (24 - MA02_hora_troca)) + NG_ate_troca[1])
            ComgasNm3Estimado.append((NG_grade_apos_troca[2] * (24 - MA02_hora_troca)) + NG_ate_troca[2])
            
        if (24 - MA03_hora_troca) == 24:
            ComgasNm3Estimado.append(dfnm3Current[3] * hRestantes)
        else: 
            ComgasNm3Estimado.append((NG_grade_apos_troca[3] * (24 - MA03_hora_troca)) + NG_ate_troca[3])

        dfOperations['Comgas (Nm³) Estimado'] = ComgasNm3Estimado

        # The code is calculating the sum of the elements in the list `ComgasNm3Estimado` and
        # assigning the result to the variable `totalComgasNm3Estimado`.
        totalComgasNm3Estimado = sum(ComgasNm3Estimado)

        # Cálculo para o Consumo diário (m³)
        consumoDiario_m3 = []

        for i in range(0, 4):
            # The code snippet is appending the sum of `dfSum[i]` and `ComgasNm3Estimado[i]` to the
            # list `consumoDiario_m3`.
            consumoDiario_m3.append(dfSum[i] + ComgasNm3Estimado[i])

        dfOperations['Consumo diário (m³)'] = consumoDiario_m3

        Total_consumoDiario_m3 = sum(consumoDiario_m3)
        
        Total_consumoDiario_m3_0 = round(Total_consumoDiario_m3, 0)
        
        # The code is converting the value of `Total_consumoDiario_m3_0` to an integer and then
        # rounding it to the nearest whole number. The result is stored in the variable
        # `Total_consumoDiario_m3_rounded`.
        Total_consumoDiario_m3_rounded = int(Total_consumoDiario_m3_0)
        
        criar_log("Previsão do consumo do dia sem o poder calorífico: ")
        criar_log(Total_consumoDiario_m3_rounded)
        criar_log('\n')

        # print(valor_com_ponto_str)
        # print(Total_consumo_dia_comgas_rounded)
        
        # Script para colocar o valor de consumo no site da Comgas
        
        criar_log(dfOperations)
        criar_log("\n")

        criar_log("Total Nm³/h (Cabot) Passed: ")
        criar_log(total_nm3hPassed)
        criar_log("\n")

        criar_log("Total de Nm³ estimado para o resto do dia: ")
        criar_log(totalComgasNm3Estimado)
        criar_log("\n")

        criar_log("Total de consumo de NG diário (m³): ")
        criar_log(Total_consumoDiario_m3)
        criar_log("\n")
        
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
        # associated with the key 'consumoRealOntem' from the JSON data. This value is stored in the
        # variable `ng_yest_consumed_value` and printed out. If the response status code is not 200, it
        # prints an error message with the status code and response text.
        ng_yest_consumed_value = 0

        response_yesterday = requests.request("GET", url, headers=headers, data=payload)

        # Verifique se a resposta foi bem-sucedida antes de acessar os dados
        if response_yesterday.status_code == 200:
            data = response_yesterday.json()  # Converte para um dicionário
            ng_yest_consumed_value = data.get('consumoRealOntem', "Chave não encontrada")
            criar_log(f"Valor de consumo ontem: {ng_yest_consumed_value}\n")
        else:
            criar_log(f"Erro na requisição: {response_yesterday.status_code}, {response_yesterday.text}")

        # Calculating the calorific factor based on the yesterday's consumption
        
        #Cálculo para ver se ele acertou ontem:
        # The above Python code is calculating different variations of a value stored in the variable
        # `ng_yest_consumed_value`. It calculates `variacao_maior` by adding 5% to the value, `variacao_menor`
        # by subtracting 5% from the value, `variacao_tres_maior` by adding 3% to the value, and
        # `variacao_tres_menor` by subtracting 3% from the value.
        variacao_maior = int(ng_yest_consumed_value) + (int(ng_yest_consumed_value)* 0.05)
        variacao_menor = int(ng_yest_consumed_value) - (int(ng_yest_consumed_value)* 0.05)
        variacao_tres_maior = int(ng_yest_consumed_value) + (int(ng_yest_consumed_value)* 0.03)
        variacao_tres_menor = int(ng_yest_consumed_value) - (int(ng_yest_consumed_value)* 0.03)
        
        criar_log(f"A varição maior é de: {variacao_maior}")
        criar_log(f"A varição menor é de: {variacao_menor} \n")
        
        # The code is checking if the value of `ontem_Total_consumo_dia_comgas_rounded` falls within a certain
        # range defined by the variables `variacao_tres_maior`, `variacao_menor`, `variacao_maior`, and
        # `variacao_tres_menor`. The variable `acertou` is set to `True` if the value is within the range
        # defined by `variacao_tres_maior` and `variacao_menor`, and `acertou_tres_porc` is set to `True` if
        # the value is within
        acertou = int(ontem_Total_consumo_dia_comgas_rounded) <= int(variacao_tres_maior) and int(ontem_Total_consumo_dia_comgas_rounded) >= int(variacao_menor)
        acertou_tres_porc = int(ontem_Total_consumo_dia_comgas_rounded) <= int(variacao_maior) and int(ontem_Total_consumo_dia_comgas_rounded) >= int(variacao_tres_menor)
        
        criar_log(f"{BLUE}Conferência se o script acertou ontem: {acertou}{END}\n")           
        
        #Cálculo para pegar o poder calorífico:       
        
        # The above Python code snippet is checking multiple conditions using nested `if` statements. Here is
        # a breakdown of what the code is doing:
        if acertou:
            if nao_teve_troca:
                if acertou_tres_porc:
                    criar_log(f"{BLUE}Devido ao acerto do script no dia de ontem, o poder calorífico se mantém em: {poder_calorifico_backup}{END}")
                    poder_calorifico = poder_calorifico_backup
                else:
                    criar_log(f"{BLUE}Devido a divergência maior que 3% do esperado no dia de ontem, o poder calorífico será reajustado.{END}")                    
                    poder_calorifico = int(ng_yest_consumed_value) / int(ontem_Total_consumoDiario_m3_rounded)
                    poder_calorifico_backup = poder_calorifico
                    criar_log("O novo poder calorífico é: ")
                    criar_log("{:.2f}".format(poder_calorifico))
                    criar_log(poder_calorifico)
                    criar_log("\n")
        else:
            if nao_teve_troca:
                poder_calorifico = int(ng_yest_consumed_value) / int(ontem_Total_consumoDiario_m3_rounded)
                poder_calorifico_backup = poder_calorifico
                criar_log("O poder calorífico é: ")
                criar_log("{:.2f}".format(poder_calorifico))
                criar_log(poder_calorifico)
                criar_log("\n")
        
        # The above Python code snippet is checking if the variable `poder_calorifico` is less than or equal
        # to 1.02. If it is, a log message is created indicating that the calculated calorific power is too
        # low, and then the `poder_calorifico` variable is adjusted to 1.03. Additionally, the original value
        # of `poder_calorifico` is stored in the variable `poder_calorifico_backup`.
        if poder_calorifico <= 1.02:
            criar_log(f"{YELLOW}Devido ao poder calorífico calculado estar muito baixo, foi reajustado para 1.03{END}")
            poder_calorifico = 1.02
            poder_calorifico_backup = poder_calorifico
        # The code snippet is part of a Python script. It is an `elif` block that checks if a variable
        # `poder_calorifico` is greater than or equal to 1.06. If this condition is true, it logs a message
        # indicating that the calculated calorific power is too high, adjusts the `poder_calorifico` variable
        # to 1.06, and stores the original value of `poder_calorifico` in the `poder_calorifico_backup`
        # variable.
        elif poder_calorifico >= 1.06:
            criar_log(f"{YELLOW}Devido ao poder calorífico calculado estar muito alto, foi reajustado para 1.06{END}")
            poder_calorifico = 1.06
            poder_calorifico_backup = poder_calorifico
    
        # Cálculo para conseguir o valor de Consumo do dia (Comgas)
        # The above Python code is calculating the total daily consumption in comgas units based on the
        # calorific power and the daily consumption in cubic meters. It first calculates the total consumption
        # in comgas units, then rounds it to the nearest whole number, and finally converts it to an integer.
        Total_consumo_dia_comgas = poder_calorifico * Total_consumoDiario_m3
        
        Total_consumo_dia_comgas_rounded_0 = round(Total_consumo_dia_comgas, 0)
        
        Total_consumo_dia_comgas_rounded = int(Total_consumo_dia_comgas_rounded_0)
        
        criar_log("Previsão de consumo para o dia: ")
        criar_log(f'{Total_consumo_dia_comgas_rounded}\n')
        
        # Cálculo para conferir se o valor não está muito discrepante
        # The above Python code snippet is checking if there was no exchange yesterday and if the
        # current consumption value is within a certain margin compared to the total consumption value
        # from yesterday. If the current consumption value is within the margin, it logs a message
        # stating that the consumption forecast is within the margin of the last few days. If the
        # current consumption value is not within the margin, it logs a warning message indicating
        # that the values for today are significantly different from yesterday's value. The warning
        # message is displayed in bold blue text to draw attention to the discrepancy.
        
        if ontem_nao_teve_troca:
            if dentro_da_margem(ng_yest_consumed_value, Total_consumo_dia_comgas_rounded):
                criar_log("A previsão de consumo está dentro da margem dos últimos dias")
            else:
                criar_log(f"{BOLD}{BLUE}ATENÇÃO!!{END}")
                criar_log(f"{BOLD}{BLUE}OS VALORES DE HOJE ESTÃO MUITO DISCREPANTES DO VALOR DE ONTEM{END}")
                criar_log(f"{BOLD}{BLUE}ATENÇÃO!!{END}")
                criar_log(f"{BOLD}{BLUE}OS VALORES DE HOJE ESTÃO MUITO DISCREPANTES DO VALOR DE ONTEM{END}")
                
        # Convertemos para string
        valor_com_ponto_str = str(Total_consumo_dia_comgas_rounded)
        # Inserimos um ponto na terceira casa decimal da direita para a esquerda
        posicao_inserir_ponto = len(valor_com_ponto_str) - 3
        valor_com_ponto_str = valor_com_ponto_str[:posicao_inserir_ponto] + '.' + valor_com_ponto_str[posicao_inserir_ponto:]
            
        # The above Python code is opening a file in write mode and writing several variables and
        # their values to the file. The variables and values being written include information related
        # to consumption, gas values, calorific power, manual input, and a flag indicating if a
        # certain condition was met yesterday.

        # Caminho do arquivo no SharePoint
        file_name = "ng_consuption_prediction_yest_info.txt"
        file_url = f"{pastagn_relative_url}/{file_name}"
        local_path = file_name # Caminho local para download temporário do arquivo

        # Fazer o download do arquivo para edição local
        file_write_yest = ctx.web.get_file_by_server_relative_url(file_url)
        with open(local_path, "wb") as local_file:
            file_write_yest.download(local_file)
            ctx.execute_query()
        print(f"Arquivo baixado para edição local: {local_path}")
        
        # Editar o arquivo localmente
        with open(local_path, "w", encoding="utf-8") as arquivo:
            arquivo.write(f'ontem_nao_teve_troca={nao_teve_troca}\n')
            arquivo.write(f'ontem_Total_consumoDiario_m3_rounded={Total_consumoDiario_m3_rounded}\n')
            arquivo.write(f'ontem_Total_consumo_dia_comgas_rounded={Total_consumo_dia_comgas_rounded}\n')
            arquivo.write(f'ontem_ng_yest_consumed_value={ng_yest_consumed_value}\n')
            arquivo.write(f'poder_calorifico={poder_calorifico}\n')
            arquivo.write(f'poder_calorifico_backup={poder_calorifico_backup}\n')
            arquivo.write(f'input_manual={input_manual}\n')
            arquivo.write(f'acertou_ontem={acertou}\n')
        print(f"Arquivo editado localmente: {local_path}")
        
        # Fazer upload do arquivo atualizado de volta para o SharePoint
        with open(local_path, "rb") as updated_file:
            target_folder = ctx.web.get_folder_by_server_relative_url(pastagn_relative_url)
            target_file = target_folder.upload_file(file_name, updated_file.read())
            ctx.execute_query()
        print(f"Arquivo atualizado com sucesso no SharePoint: {file_url}")


        criar_log("Informações salvas para base de cálculo para amanhã!")
        criar_log("\n")    
            
        # Inserting the predicted value of NG consumption for today
        
        # The above Python code is making a POST request to the URL
        # "https://industrial.comgas.com.br/api/PI/previsaoConsumo" with a JSON payload containing user
        # information, consumption forecast data for the current date, and a unique ID. The request is being
        # sent with specific headers including accept, authorization, content-type, user-agent, and others.
        # The response from the server is then printed out.
        
        today_date = datetime.datetime.today().strftime("%Y%m%d")
        
# The above Python code is sending a POST request to the URL
# "https://industrial.comgas.com.br/api/PI/previsaoConsumo" with a JSON payload containing the user
# code, email, and a list of consumption forecast data. The consumption forecast data includes the
# date of the forecast (stored in the variable `today_date`) and the predicted consumption value for
# that date (stored in the variable `Total_consumo_dia_comgas_rounded`).
        url = "https://industrial.comgas.com.br/api/PI/previsaoConsumo"

        payload = json.dumps({
        "codUsuario": "000003751112",
        "email": "ewerton.araujo@cabotcorp.com",
        "previsaoConsumo": [
            {
            "dataPrevisao": today_date,
            "valorPrevisao": Total_consumo_dia_comgas_rounded
            }
        ],
        "uniqueId": "NjE3NDE2OTAwMDAxMjQxNzM2OTQ0MTg5Nzc3"
        })
        headers = {
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7,pt-PT;q=0.6',
        'authorization': f'ComgasToken {collection_variables["Auth-consumo"]}',
        'content-type': 'application/json',
        'cookie': '_ga_WF5KF8E9JC=GS1.1.1702557914.49.1.1702557979.60.0.0; _ga=GA1.1.1308956571.1696957003; _ce.irv=notReturning; cebs=1; _ce.clock_data=938%2C200.218.172.232%2C1%2Cffc3218438300d069a0fd5dfa5c6e851%2CEdge%2CBR; _ga_X8F0YFKT06=GS1.1.1736943896.94.1.1736943914.0.0.0; cebsp_=178; OptanonConsent=isGpcEnabled=0&datestamp=Wed+Jan+15+2025+09%3A29%3A47+GMT-0300+(Brasilia+Standard+Time)&version=202407.2.0&browserGpcFlag=0&isIABGlobal=false&hosts=&consentId=ddaa7d8d-94a7-4d1d-9768-e8d6f17fad13&interactionCount=1&isAnonUser=1&landingPath=NotLandingPage&groups=C0001%3A1&intType=1&geolocation=BR%3BSP&AwaitingReconsent=false; OptanonAlertBoxClosed=2025-01-15T12:29:47.514Z; _ga_VY7D8X5KGL=GS1.1.1736943904.46.1.1736944244.1.0.0; _ce.s=v~e5fb0347b3554b8d3b2f760c511a2124ac1748db~lcw~1736944269581~vpv~3~ir~1~lva~1736862153321~as~false~v11ls~a79834a0-49ab-11ef-b336-97dd2eb26962~vi~ewerton.araujo%40cabotcorp.com~v11.fhb~1736943896633~v11.lhb~1736944247390~v11slnt~1736862155371~vir~returning~v11.cs~425289~v11.s~bb0c4750-d33b-11ef-b28e-23192a580cd1~v11.sla~1736944269581~gtrk.la~m5xvrrwu~lcw~1736944269582',
        'origin': 'https://industrial.comgas.com.br',
        'priority': 'u=1, i',
        'sec-ch-ua': '"Microsoft Edge";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0',
        'x-msg-timestamp': '1736944269675'
        }
# The above Python code is making a POST request to a specified URL using the `requests` library. It
# includes headers and payload data in the request.

        response = requests.request("POST", url, headers=headers, data=payload)

        print(response.text)
        print("Valor inseridos com sucesso.")
        print("Valor inseridos com sucesso.")
        print("Valor inseridos com sucesso.")
        
        time.sleep(10) 
        
        # ====================================================================================================================================================================================
        
        criar_log(dfOperations)
        criar_log("\n")

        criar_log("Total Nm³/h (Cabot) Passed: ")
        criar_log(total_nm3hPassed)
        criar_log("\n")

        criar_log("Total de Nm³ estimado para o resto do dia: ")
        criar_log(totalComgasNm3Estimado)
        criar_log("\n")

        criar_log("Total de consumo de NG diário (m³): ")
        criar_log(Total_consumoDiario_m3)
        criar_log("\n")

        criar_log("//---------------------------------------------------------//")
        criar_log(f"   TOTAL DE CONSUMO DO DIA (COMGAS): {Total_consumo_dia_comgas_rounded}")
        # criar_log(Total_consumo_dia_comgas_rounded)
        criar_log("//---------------------------------------------------------//\n")

        criar_log("Horas passadas: ")
        criar_log(hPassados)
        criar_log("Horas restantes: ")
        criar_log(hRestantes)
        criar_log("\n")

        criar_log(df_grades_consumo.head(10))
        criar_log("\n")

    except Exception as e:
        # Registra o erro no log
        criar_log(f"Erro: {str(e)}")
        traceback.print_exc()

    # Exemplo: Simulando a conclusão bem-sucedida do script
    criar_log("Script finalizado.")
    criar_log("\n")

    criar_log("Prestes e encerrar o código...")
    sys.exit("Encerrando o código!")