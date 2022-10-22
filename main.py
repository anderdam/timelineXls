# Imports
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook

# Início da conversão de csv para xls.
# Foi utilizado openpyxl ao invés de pandas.to_excel() devido a mensagem de erro indicando que a função
# será depreciada.

# Caminho e arquivo de entrada e saída
input_path = 'C:\\Users\\adapp\\Desktop\\'
input_file = 'agend.csv'
out_path = 'C:\\Users\\adapp\\Desktop\\'
out_file = 'agendXls.xls'

# Cria a planilha (Workbook) e a deixa ativa para alterações
wb = Workbook()
ws = wb.active

# Utilizando pandas é realizada a leitura do arquivo em csv
df = pd.read_csv(input_path + input_file)

# Loop para adicionar cada linha do csv na planilha criada
for r in dataframe_to_rows(df, index=True, header=True):
    ws.append(r)

# Salvar a nova planilha, sendo que caso já exista a planilha, a nova sobrescreve a anterior
wb.save(out_path + out_file)

print(f'O arquivo {out_file} foi gerado com sucesso em {out_path}')

# Fim da conversão de csv para xls e início da formatação dos dados

df = pd.read_excel(out_path + out_file)
df.dropna()
print(df.head())
