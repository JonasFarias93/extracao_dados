import os
import pandas as pd
import re
import openpyxl


pasta_raiz = 'C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas'


pastas = []
arquivos = []


def puxando_java():
    sheet = workbooks.active
    cell  = sheet['C1']
    valor_java = cell.value
    valor_java_tratado = valor_java.split(":")
    valor_tratado = valor_java_tratado[1]
    print(valor_tratado)
    print(type(valor_tratado))

def listar_arquivos_em_pastas(pasta_raiz):
    for raiz, diretorios, arquivos in os.walk(pasta_raiz):
        nome_pasta = os.path.basename(raiz)
        if nome_pasta.startswith("LISTAS"):
            pastas.append(nome_pasta)
        #print(f"Pasta: {raiz}")
        for arquivo in arquivos:
            if arquivo.endswith(('.xlsx', '.xls')):  # Filtra por arquivos Excel
                print(f"  Arquivo: {arquivo}")
                

# Exemplo de uso:

listar_arquivos_em_pastas(pasta_raiz)
print(pastas)