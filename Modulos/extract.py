import pandas as pd
import openpyxl 
import os as os


def listar_nomes_arquivos():
    # Lista o nome dos arquivos planilha
    pasta = 'C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas\\LISTAS - (01) JANEIRO'
    arquivo = os.listdir(pasta)
    workbooks = [arquivo for arquivo in arquivo if arquivo.endswith(('.xlsx', '.xls'))]
    nomes_sheets = []
    for workbooks in workbooks:
        nomes_sheets.append(workbooks)
        print(workbooks)

import os

def listar_nomes_pastas():
    # Lista o nome das pastas
    pasta = 'C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas'
    arquivos_e_pastas = os.listdir(pasta)  # Lista todos os itens da pasta
    pastas = [item for item in arquivos_e_pastas if os.path.isdir(os.path.join(pasta, item))]
    nome_pastas = []
    for pasta in pastas:
        nome_pastas.append(pasta)
    print(pastas)

listar_nomes_pastas()
listar_nomes_arquivos()

