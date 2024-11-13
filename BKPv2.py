import os
import shutil
from datetime import datetime


# Pegando a data atual no formato mmddyy
data_atual = datetime.now()
data_formatada = data_atual.strftime("%m%d%y")

# Setando os caminhos possíveis para a pasta do Outlook
caminho1 = "Arquivos do Outlook"
caminho2 = "Outlook files"

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
    nome_pasta = f'{user}-bkp-{data_formatada}'
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

    # Removendo a pasta criada no menu em "Documentos"
    try:
        shutil.rmtree(os.path.join(documentos, nome_pasta))
        print(f'A pasta {nome_pasta} foi removida com sucesso.')
    except PermissionError:
        print(f'Permissão negada para remover a pasta {nome_pasta}.')
    except Exception as e:
        print(f'Ocorreu um erro ao remover a pasta: {e}')

elif caminho2 in pastas_documentos:
    print(f'Foi encontrado a pasta {caminho2}')

    # Criando uma pasta de backup com o nome de usuário e data
    nome_pasta = f'{user}-bkp-{data_formatada}'
    caminho_pasta_backup = os.path.join(documentos, nome_pasta)
    
    # Verificando se a pasta de backup já existe, caso não, cria
    if not os.path.exists(caminho_pasta_backup):
        os.mkdir(caminho_pasta_backup) # Criando a pasta de backup

    # Copiando a pasta do Outlook para o novo diretório de backup
    caminho_origem_outlook = os.path.join(documentos, caminho2)
    caminho_destino_outlook = caminho_pasta_backup

    # Usando dirs_exist_ok=True, se já existir a pasta de destino, irá sobrescrever
    shutil.copytree(caminho_origem_outlook, caminho_destino_outlook, dirs_exist_ok=True) # Copiando a pasta inteira

    # Movendo a pasta de backup para a rede
    caminho_origem_rede = caminho_pasta_backup
    caminho_destino_rede = rede
    shutil.move(caminho_origem_rede, caminho_destino_rede) # Movendo a pasta criada para a rede

    print(f'A pasta {nome_pasta} foi movida para {rede}')
else:
    print(f'A pasta {caminho1} ou {caminho2} não foi encontrada em {documentos}. Sem backups para acontecer. Em caso de dúvidas ler o READ ME')