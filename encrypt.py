from cryptography.fernet import Fernet

# Leitura da chave
with open('key.key', 'rb') as key_file:
    key = key_file.read()

cipher = Fernet(key)

# Dados a serem encriptados (exemplo com usuário e senha)
# The `config_data` variable is storing a byte string that contains sensitive configuration data in
# the format of key-value pairs. In this case, it includes information such as a username, password,
# site URL, and relative URL. This data is going to be encrypted using the Fernet encryption algorithm
# before being saved to a file for secure storage or transmission.
config_data = "USERNAME=mau.operation@cabotcorp.com\nPASSWORD=Sala%50qualy\nSITE_URL=https://cabotcorp.sharepoint.com/sites/MauaProd\nRELATIVE_URL=/Shared Documents/General/PRODUÇÃO/Planejamento/Planejamento de produção/Troca de Grau - GN.csv".encode("utf-8")

# https://cabotcorp.sharepoint.com/sites/MauaProd/Shared%20Documents/General/PRODU%C3%87%C3%83O/Planejamento/Planejamento%20de%20produ%C3%A7%C3%A3o/Troca%20de%20Grau%20-%20GN.csv

# Encriptação dos dados
encrypted_data = cipher.encrypt(config_data)

# Salve os dados encriptados em um arquivo
with open('config.enc', 'wb') as config_file:
    config_file.write(encrypted_data)
  