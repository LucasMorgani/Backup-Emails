from datetime import datetime

data_atual = datetime.now()
# Y retorna ano completo - y retorna somente os ultimos 2 digitos do ano
data_formatada = data_atual.strftime("%m%d%y")
print(data_formatada)