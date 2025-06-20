import os
import datetime
import pandas as pd

# Defina o caminho do diretório raiz
diretorio = "D:\Clientes\0046-CLAW\\"

# Lista para armazenar os dados coletados
dados = []

# Utiliza os.walk para percorrer recursivamente o diretório
for root, dirs, files in os.walk(diretorio):
    for arquivo in files:
        caminho_completo = os.path.join(root, arquivo)
        # Obtém a data de modificação em formato timestamp e converte para datetime
        tempo_modificacao = os.path.getmtime(caminho_completo)
        data_modificacao = datetime.datetime.fromtimestamp(tempo_modificacao)
        dados.append({
            'Arquivo': arquivo,
            'Caminho Completo': caminho_completo,
            'Data de Modificação': data_modificacao
        })

# Cria um DataFrame com os dados
df = pd.DataFrame(dados)

# Exporta os dados para um arquivo Excel
df.to_excel('dados_modificacao.xlsx', index=False)

print("Exportação concluída com sucesso!")