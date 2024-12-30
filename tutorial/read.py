import openpyxl
import datetime

from openpyxl import load_workbook

caminho_arquivo = "C:/Users/Quixabeira/Documents/dados.xlsx"
workbook = load_workbook(caminho_arquivo)

# Selecionar uma planilha (por nome ou ativa)
planilha =   workbook['VENDAS'] # ou workbook.active

# Ler valores das células
for linha in planilha:
    # Obter os valores das células, ignorando None
    linha = [celula.value for celula in linha if celula.value is not None]
    if linha != []:
        for celula in linha:
            if celula == "PRODUTO":
                celula = "PRODUTO                      "

            if isinstance(celula, datetime.datetime):
                    celula = celula.strftime("%Y-%m-%d") 
            print(celula, end='\t') 
        print("\b")
        

# Fechar o arquivo (boa prática)
workbook.close()