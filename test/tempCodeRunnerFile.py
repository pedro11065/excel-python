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
