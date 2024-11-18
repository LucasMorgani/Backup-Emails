import os
import shutil
import requests
from datetime import datetime


# Pegando a data atual no formato mmddyy
data_atual = datetime.now()
data_formatada = data_atual.strftime("%m%d%y")

# Setando os caminhos possíveis para a pasta do Outlook
caminho1 = "Arquivos do Outlook"

# Setando o usuário da máquina atualmente logado
user = os.environ.get('USERNAME')

# Setando caminho para a pasta Documents
documentos = f'C:/Users/{user}/Documents'

# Setando caminho para a pasta da rede (usando uma string bruta)
rede = r'\\192.168.0.230\Volume_1\2024\SEDE'

# Buscando a pasta de arquivos do outlook na pasta documentos
pastas_documentos = [pasta for pasta in os.listdir(documentos) if os.path.isdir(os.path.join(documentos, pasta))]

# Caso encontre a pasta "Arquivos do Outlook"
if caminho1 in pastas_documentos:
    print(f'Foi encontrado a pasta: {caminho1}')
    
    # Criando uma pasta de backup com o nome de usuário e data
    nome_pasta = f'{user}-{data_formatada}'
    caminho_pasta_backup = os.path.join(documentos, nome_pasta)
    
    # Verificando se a pasta de backup já existe, caso não, cria
    if not os.path.exists(caminho_pasta_backup):
        os.mkdir(caminho_pasta_backup)  # Criando a pasta de backup
    
    # Copiando a pasta do Outlook para o novo diretório de backup
    caminho_origem_outlook = os.path.join(documentos, caminho1)
    caminho_destino_outlook = caminho_pasta_backup
    print('Foi criado a pasta que será enviada para a rede')
    
    # Usando dirs_exist_ok=True, se já existir a pasta de destino, irá sobrescrever
    shutil.copytree(caminho_origem_outlook, caminho_destino_outlook, dirs_exist_ok=True)  # Copiando a pasta inteira
    
    # Movendo a pasta de backup para a rede
    caminho_origem_rede = caminho_pasta_backup
    caminho_destino_rede = rede
    shutil.move(caminho_origem_rede, caminho_destino_rede)  # Movendo a pasta criada para a rede

    print(f'A pasta {nome_pasta} foi movida para {rede}')

    ############ BLOCO TELEGRAM ############
    bot_token = "7665194645:AAH5MyExs_gP1XBAFEGQsCsA69cSXVK0dlM"  # Substitua pelo token do seu bot
    chat_id = 6947234667  # O chat_id obtido
    message = f'O backup do usuário {user} foi concluído'

    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    params = {
        "chat_id": chat_id,
        "text": message
    }

    response = requests.get(url, params=params)

    if response.status_code == 200:
        print("Mensagem enviada com sucesso!")
    else:
        print("Falha ao enviar mensagem. Status:", response.status_code)

else:
    print(f'A pasta {caminho1} não foi encontrada em {documentos}. Sem backups para acontecer. Em caso de dúvidas ler o READ ME')