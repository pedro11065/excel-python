import openpyxl

# Carregar o arquivo Excel
caminho_arquivo = "C:/Users/Quixabeira/Documents/dados.xlsx"
workbook = openpyxl.load_workbook(caminho_arquivo)

# Selecionar a planilha
planilha = workbook['VENDAS']  # ou workbook.active

# Editar o valor de uma célula
planilha['A1'] = "ID"  # A célula A1 recebe o novo valor

# Para editar uma célula baseada em uma coordenada específica (linha e coluna)

planilha.cell(row=1, column=2).value = "PRODUTO"

print("Dados atualizados com sucesso!!")

# Salvar o arquivo após as edições
workbook.save(caminho_arquivo)

# Fechar o arquivo (boa prática)
workbook.close()
