from openpyxl import Workbook
import os

icms = float(input('Qual é a tarifa de ICMS do seu estado? Digite a porcentagem:\n'))
icms = icms / 100
icms_total_geral = 0

n_vezes = int(input('Digite a quantidade de produtos que deseja calcular o ICMS:\n'))

# Cria uma planilha nova
wb = Workbook()

# Seleciona a planilha principal
ws = wb.active

# Define o nome da aba
ws.title = "Relatorio ICMS"

# Cabeçalhos 
ws.append(["Produto", "Quantidade", "Valor Unitário", "Valor Total", "ICMS"])

for i in range(n_vezes):
    print(f'----- PRODUTO {i + 1} -----')
    
    produto = input('Digite o nome do produto:\n')
    valor_unitario = float(input('Digite o valor unitário do produto:\n'))
    qtd_vendida = int(input('Digite a quantidade de produtos que foi vendida:\n'))
    
    valor_total = valor_unitario * qtd_vendida
    valor_icms = valor_total * icms
    icms_total_geral += valor_icms
    
    print(f'\n➡ Produto: {produto}')
    print(f'  Valor total vendido: R$ {valor_total:.2f}')
    print(f'  ICMS gerado: R$ {valor_icms:.2f}\n')

    # Grava no Excel
    ws.append([produto, qtd_vendida, valor_unitario, valor_total, valor_icms])

# Salva o arquivo
wb.save("relatorio_icms.xlsx")

print("\nArquivo Excel criado com sucesso!")
print("Nome do arquivo: relatorio_icms.xlsx")
print(f"Total geral de ICMS dos produtos: R$ {icms_total_geral:.2f}")

os.startfile("relatorio_icms.xlsx")