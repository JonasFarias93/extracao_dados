import os
import pandas as pd
import re
import openpyxl



#puxa o endereço da pasta
pasta_raiz = 'C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas\\'
# recebe os nomes das pastas
pastas = ['1','2','3','4']
arquivos = []

def Concatenando_nomes_e_pastas():
    # Concatenando os caminhos completos
    caminhos_completos = [os.path.join(pasta_raiz, pasta) for pasta in pastas]
    # Imprimindo os caminhos completos
    for caminho in caminhos_completos:
        print(caminho)

Concatenando_nomes_e_pastas()