import os
import pandas as pd
import re
import openpyxl



#puxa o endereço da pasta
pasta_raiz = 'C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas\\'
# recebe os nomes das pastas
pastas = []
arquivos = []

'''#puxa o valor java da planilha em uma celula expesifica
def puxando_java():
    sheet = workbooks.active
    cell  = sheet['C1']
    valor_java = cell.value
    valor_java_tratado = valor_java.split(":")
    valor_tratado = valor_java_tratado[1]
    print(valor_tratado)
    print(type(valor_tratado))'''

#lista o nome das pastas e adiciona na lista pastas
def listar_arquivos_em_pastas(pasta_raiz):
    for raiz, diretorios, arquivos in os.walk(pasta_raiz):
        nome_pasta = os.path.basename(raiz)
        if nome_pasta.startswith("LISTAS"):
            pastas.append(nome_pasta)
            print(f"Pasta: {nome_pasta}")
            for arquivo in arquivos:
                if arquivo.endswith(('.xlsx', '.xls')):
                    print(f"  Arquivo: {arquivo}")
listar_arquivos_em_pastas(pasta_raiz)


def Concatenando_nomes_e_pastas(pasta_raiz, pastas):
    # Concatenando os caminhos completos
    caminhos_completos = [os.path.join(pasta_raiz, pasta) for pasta in pastas]
    return caminhos_completos
    # Imprimindo os caminhos completos
    for caminho in caminhos_completos:
        pass
        #print(type(caminho))


def selecionando_colunas_desejadas(caminhos_completos):
    for caminho in caminhos_completos:
        df = pd.read_excel(caminho, header=1)
        #seleciona as colunas desejadas
        df_selecionado = df.iloc[:, :7]
        print(df_elecionado)


caminhos = Concatenando_nomes_e_pastas(pasta_raiz, pastas)
selecionando_colunas_desejadas(caminhos)


