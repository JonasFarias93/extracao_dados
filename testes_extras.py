import pandas as pd
import re
import openpyxl

pastas = ['LISTAS - (01) JANEIRO', 'RELATORIOS', 'DADOS']
arquivos = ['equipamentos loja - DS 2196.xlsx', 'vendas_mensal.xlsx', 'clientes.xlsx']

for pasta in pastas:
    for arquivo in arquivos:
        caminho_completo = criar_caminho_completo(pasta, arquivo)
        workbooks.append(caminho_completo)
        # Processar o arquivo (abrir, modificar, etc.)