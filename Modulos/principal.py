import os
import pandas as pd
import openpyxl
import re

# Define o endereço da pasta
pasta_raiz = 'C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas\\'
javas = []

# Função para puxar o valor Java (somente numérico) de uma célula
def puxar_valor_java(caminho_arquivo, celula='C1'):
    try:
        # Carrega a planilha
        workbook = openpyxl.load_workbook(caminho_arquivo)
        sheet = workbook.active
        
        # Obtém o valor da célula especificada
        cell_value = sheet[celula].value
        
        # Verifica se a célula não está vazia
        if cell_value:
            # Usa regex para encontrar o primeiro valor numérico na célula
            valor_numerico = re.search(r'\d+', str(cell_value))
            
            if valor_numerico:
                valor_tratado = valor_numerico.group()  # Extrai o valor numérico encontrado
                return valor_tratado
            else:
                print(f"Nenhum valor numérico encontrado na célula {celula}.")
        else:
            print(f"Célula {celula} está vazia.")
        
    except Exception as e:
        print(f"Erro ao processar o arquivo {caminho_arquivo}: {e}")
        return None

# Função para listar os arquivos em pastas que começam com "LISTAS"
def listar_arquivos_em_pastas(pasta_raiz):
    arquivos_encontrados = []
    for raiz, diretorios, arquivos in os.walk(pasta_raiz):
        nome_pasta = os.path.basename(raiz)
        if nome_pasta.startswith("LISTAS"):
            for arquivo in arquivos:
                if arquivo.endswith(('.xlsx', '.xls')):
                    arquivos_encontrados.append(os.path.join(raiz, arquivo))
    return arquivos_encontrados

# Função para selecionar as primeiras 7 colunas de cada arquivo e exibir as primeiras linhas
def selecionando_colunas_desejadas(caminhos_completos):
    for caminho in caminhos_completos:
        try:
            df = pd.read_excel(caminho, header=1)
            df_selecionado = df.iloc[:, :7]
            #print(f"Arquivo: {os.path.basename(caminho)}")
            #print(df_selecionado.head())  # Exibe as primeiras linhas do DataFrame selecionado

            # Puxa o valor Java e o utiliza na função de ordenação de colunas
            valor_java = puxar_valor_java(caminho, celula='C1')
            if valor_java:
                ordenando_as_colunas(df_selecionado, valor_java, caminho)
        except Exception as e:
            print(f"Erro ao processar o arquivo {caminho}: {e}")

def ordenando_as_colunas(df_selecionado, valor_java, caminho):
    # Manipula o DataFrame conforme o solicitado
    df_selecionado = df_selecionado.drop('GRUPO', axis=1, errors='ignore')
    df_selecionado = df_selecionado.rename(columns={'NF': 'NF ENTRADA'})
    df_selecionado.insert(0, 'JAVA', valor_java)
    df_selecionado.insert(1, 'FILIAL', ' ')
    df_selecionado.insert(2, 'BANDEIRA', ' ')
    df_selecionado.insert(3, 'UF', ' ')
    
    # Exibe o DataFrame resultante
    #print("DataFrame após ordenação das colunas:")
    #print(df_selecionado.head())  # Exibe as primeiras linhas do DataFrame
    # Salva o DataFrame modificado em um novo arquivo
    novo_caminho = caminho.replace('.xlsx', '_modificado.xlsx')  # Define o caminho do novo arquivo
    df_selecionado.to_excel(novo_caminho, index=False)
    #print(f"Arquivo salvo com sucesso em: {novo_caminho}")

# Puxa o valor Java de cada arquivo e armazena na lista javas
arquivos = listar_arquivos_em_pastas(pasta_raiz)
selecionando_colunas_desejadas(arquivos)

for arquivo in arquivos:
    java = puxar_valor_java(arquivo, celula='C1')
    if java:
        javas.append(java)

print("Valores Java extraídos:", javas)

def concatenar_arquivos_excel(diretorio_entrada, caminho_saida):
    # Cria a pasta de saída, se não existir
    caminho_saida = os.path.join(diretorio_entrada, nome_pasta_saida)
    os.makedirs(caminho_saida, exist_ok=True)

    # Cria uma lista para armazenar todos os DataFrames
    lista_dfs = []

    # Percorre todos os arquivos no diretório de entrada
    for nome_arquivo in os.listdir(diretorio_entrada):
        if nome_arquivo.endswith('.xlsx'):
            caminho_arquivo = os.path.join(diretorio_entrada, nome_arquivo)
            # Lê o arquivo Excel e adiciona o DataFrame à lista
            df = pd.read_excel(caminho_arquivo)
            lista_dfs.append(df)

    # Salva o DataFrame concatenado na nova pasta
    nome_arquivo_saida = "arquivo_concatenado.xlsx"  # Você pode personalizar o nome
    caminho_completo = os.path.join(caminho_saida, nome_arquivo_saida)
    df_concatenado.to_excel(caminho_completo, index=False)
    print(f"Arquivo concatenado salvo com sucesso em: {caminho_completo}")

# uso
#concatenar_arquivos_excel(pasta_raiz, 'C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados')
