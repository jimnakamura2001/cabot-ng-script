from cryptography.fernet import Fernet

# Leitura da chave
with open('key.key', 'rb') as key_file:
    key = key_file.read()

cipher = Fernet(key)

# Dados a serem encriptados (exemplo com usuário e senha)
config_data = b'CNPJ=61.741.690/0001-24\nCRED=24512370\nMAIL=ewerton.araujo@cabotcorp.com\nPASSW=Tom@r@ce2010'

# Encriptação dos dados
encrypted_data = cipher.encrypt(config_data)

# Salve os dados encriptados em um arquivo
with open('configcg.enc', 'wb') as config_file:
    config_file.write(encrypted_data)
