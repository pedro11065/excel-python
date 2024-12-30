import openpyxl
import datetime

from openpyxl import load_workbook

caminho_arquivo = "C:/Users/Quixabeira/Documents/dados.xlsx"
workbook = load_workbook(caminho_arquivo)

# Selecionar uma planilha (por nome ou ativa)
worksheet =   workbook['VENDAS'] # ou workbook.active

line_count = 0

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
            print(f'\nForam vendidos {info_sum} {line_info[0]} ')

# Fechar o arquivo (boa prática)
workbook.close()