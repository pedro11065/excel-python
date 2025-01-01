import openpyxl
import datetime

from openpyxl import load_workbook

caminho_arquivo = "C:/Users/Quixabeira/Documents/Projetos/excel-python/dados.xlsx"
workbook = load_workbook(caminho_arquivo)

# Selecionar uma planilha (por nome ou ativa)
worksheet =   workbook['VENDAS'] # ou workbook.active

line_count = 0

names = []
total = []

# Ler valores das linhas
for line in worksheet:

    # Obter os valores das células, ignorando None
    line = [cell.value for cell in line if cell.value is not None]

    line_info = []

    line_count = line_count + 1 #conta as linhas
    cell_count = 0
    
    if line != [] and line_count > 1: #ignora lists vazias e a primeira linha

        for cell in line: #ler os valores das celulas e agrupalas em uma lista
            cell_count = cell_count + 1

            if cell_count >= 2:
                line_info.append(cell) 

#-----------------------------------------------------------------------
#Apuração dos dados
    
        info_sum = 0

        for info in line_info:
            try:
                int(info)
                info_sum += info
            except:
                None

        if info_sum > 0:

            names.append(line_info[0])
            total.append(info_sum)


print("\nForam vendidos:\n")
for name, num in zip(names, total):
    print(f" {name}: {num}")
print('')

most_sold_index = -1
most_sold = 0

for index, value in enumerate(total):
    if value > most_sold:
        most_sold = value
        most_sold_index = index

print(f"O produto mais vendido foi: {names[most_sold_index]} \nTotal vendido: {most_sold}")

# Fechar o arquivo (boa prática)
workbook.close()