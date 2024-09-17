import os

import os

def listar_arquivos_em_pastas(pasta_raiz):
    """Lista recursivamente os arquivos em uma pasta e suas subpastas,
    imprimindo o nome da pasta e os arquivos Excel encontrados.

    Args:
        pasta_raiz (str): Caminho da pasta raiz.
    """

    for raiz, diretorios, arquivos in os.walk(pasta_raiz):
        nome_pasta = os.path.basename(raiz)
        if nome_pasta.startswith("LISTAS"):
            print(f"Pasta: {nome_pasta}")
            arquivos_excel = []
            for arquivo in arquivos:
                if arquivo.endswith(('.xlsx', '.xls')):
                    arquivos_excel.append(arquivo)
                    print(f"  Arquivo: {arquivo}")

# Exemplo de uso:
pasta_raiz = 'C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas'
listar_arquivos_em_pastas(pasta_raiz)