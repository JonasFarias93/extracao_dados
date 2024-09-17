import os
import pandas as pd

# Define o endereço da pasta
pasta_raiz = 'C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas\\'

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
            print(f"Arquivo: {os.path.basename(caminho)}")
            print(df_selecionado.head())  # Exibe as primeiras linhas do DataFrame selecionado
        except Exception as e:
            print(f"Erro ao processar o arquivo {caminho}: {e}")

# Exemplo de uso
arquivos = listar_arquivos_em_pastas(pasta_raiz)
selecionando_colunas_desejadas(arquivos)
