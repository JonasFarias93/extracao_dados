import os

def listar_arquivos_em_pastas(pasta_raiz):
    """
    Lista recursivamente os arquivos em uma pasta e suas subpastas.

    Args:
        pasta_raiz (str): Caminho da pasta raiz a ser explorada.
    """

    for raiz, diretorios, arquivos in os.walk(pasta_raiz):
        print(f"Pasta: {raiz}")
        for arquivo in arquivos:
            if arquivo.endswith(('.xlsx', '.xls')):  # Filtra por arquivos Excel
                print(f"  Arquivo: {arquivo}")

# Exemplo de uso:
pasta_raiz = 'C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas'
listar_arquivos_em_pastas(pasta_raiz)