import os
import shutil
from datetime import datetime


# Pegando a data atual no formato mmddyy
data_atual = datetime.now()
data_formatada = data_atual.strftime("%m%d%y")

# Setando os caminhos possiveis para a pasta
caminho1 = "Arquivos do Outlook"
caminho2 = "Outlook files"

# Setando o usuário da máquina atualmente logado
user = os.environ.get('USERNAME')

# Setando caminho para a pasta Documents
documentos = f'C:/Users/{user}/Documents'

# Setando caminho para a pasta da rede
rede = '\\SRVSP089\Backup'

# Buscando a pasta de arquivos do outlook na pasta documentos
pastas_documentos = [pasta for pasta in os.listdir(documentos) if os.path.isdir(os.path.join(documentos, pasta))]
# Caso encontre a pasta "Arquivos do Outlook"
if caminho1 in pastas_documentos:
    print(f'Foi encontrado a pasta{caminho1}')
    # Acessando a pasta encontrada
    # Juntando os caminhos documentos e caminho1
    #os.chdir(os.path.join(documentos, caminho1))
    # Verificando se rolou
    #diretorio_atual = os.getcwd()
    #print(diretorio_atual)
    # Criando uma pasta com o nome de usuário da pessoa
    nome_pasta = f'{user}-bkp-{data_formatada}'
    os.mkdir(nome_pasta)
    # Copiando os arquivos do outlook para a pasta criada dentro de documentos
    caminho_origem_outlook = os.path.join(documentos, caminho1)
    caminho_destino_outlook = os.path.join(documentos, nome_pasta)
    shutil.copytree(caminho_origem_outlook, caminho_destino_outlook)
    # Movendo a pasta criada para a rede
    caminho_origem_rede = os.path.join(documentos, caminho1)
    caminho_destino_rede = rede
    shutil.move(caminho_origem_rede, caminho_destino_rede)
