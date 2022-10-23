"""Criação, formatação e gráfico de Gantt para gerar uma timeline de processos"""

import pandas as pd
import matplotlib.pyplot as plt

# Início da conversão de csv para xls.

# Caminho e arquivo de entrada e saída
INPUT_FILE = 'C:\\Users\\adapp\\Desktop\\agend.csv'
OUT_FILE = 'C:\\Users\\adapp\\Desktop\\agendXls.xls'

# Utilizando pandas é realizada a leitura do arquivo em csv e com
df = pd.read_csv(INPUT_FILE)
writer = pd.ExcelWriter(OUT_FILE, engine='xlsxwriter')  # pylint: disable=abstract-class-instantiated
df.to_excel(writer, sheet_name='Timeline')

# Salvar a nova planilha, sendo que caso já exista a planilha, a nova sobrescreve a anterior
# Exibição de mensagem e amostra confirmando a criação da planilha.

writer.save()

print(f'O arquivo {OUT_FILE} foi gerado com sucesso.')
print()
print('Amostra:')
print(df.head())

# Fim da conversão de csv para xls e início da formatação dos dados

df = pd.read_excel(OUT_FILE)

# Gantt Chart

# Project start date

df['end_time'] = df.Horas + 1
print(df)

fig, ax = plt.subplots(1, figsize=(16, 6))
ax.barh(df.mensagem, df.end_time, left=df.Horas)
plt.show()
