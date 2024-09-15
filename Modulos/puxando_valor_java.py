import pandas as pd
import re
import openpyxl


workbooks = openpyxl.load_workbook('C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas\\LISTAS - (01) JANEIRO\\equipamentos loja - DS 2196.xlsx')
sheet = workbooks.active

cell  = sheet['C1']
valor_java = cell.value
valor_java_tratado = valor_java.split(":")
valor_tratado = valor_java_tratado[1]
print(valor_tratado)
print(type(valor_tratado))