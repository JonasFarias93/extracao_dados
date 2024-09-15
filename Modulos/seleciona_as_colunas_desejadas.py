import pandas  as pd
import numpy as np 
import os as os



arquivo_excel = pd.read_excel('C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas\\LISTAS - (01) JANEIRO\\equipamentos loja - DS 2196.xlsx', header=1)

#seleciona as colunas desejadas
DataFrame_Selecionado = arquivo_excel.iloc[:, :7]
print(DataFrame_Selecionado)
