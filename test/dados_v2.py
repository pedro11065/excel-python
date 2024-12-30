import openpyxl
import datetime

from openpyxl import load_workbook

caminho_arquivo = "C:/Users/Quixabeira/Documents/dados.xlsx"
workbook = load_workbook(caminho_arquivo)

# Selecionar uma planilha (por nome ou ativa)
worksheet =   workbook['VENDAS'] # ou workbook.active

line_count = 0

data = []

# Ler valores das linhas
for line in worksheet:

    # Obter os valores das células, ignorando None
    line = [cell.value for cell in line if cell.value is not None]

    line_info = []

    line_count = line_count + 1 #conta as linhas
    cell_count = 0
    
    if line != [] and line_count > 1: #ignora lists vazias e a primeira linha
        line_info.append(line_count)

        for cell in line: #ler os valores das celulas e agrupalas em uma lista
            cell_count = cell_count + 1

            if cell_count >= 2:
                line_info.append(cell) 

        data.append(line_info)

workbook.close()

#A variável //data// contem uma Lista com listas dos dados de cada linha, o primeiro valor é a linha dessa dado no excel, ex:
#[2, 'Samsung Galaxy S23 Ultra', 10, 9, 12, 5] -> Nesse exemplo, essa dado está na linha 2 no excel

for i in data:
    print(i)

