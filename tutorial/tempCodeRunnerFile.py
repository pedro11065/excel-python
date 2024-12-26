import openpyxl
import datetime

from openpyxl import load_workbook

# Carregar o arquivo Excel
caminho_arquivo = "C:/Users/Quixabeira/Documents/dados.xlsx"
workbook = load_workbook(caminho_arquivo)

# Selecionar uma planilha (por nome ou ativa)
planilha =   workbook['VENDAS'] # ou workbook.active

# Ler valores das células
for linha in planilha:
    # Obter os valores das células, ignorando None
    linha_valores = [celula.value for celula in linha if celula.value is not None]
    if linha_valores != []:
        for valor in linha_valores:
            if "datetime" in valor: