import pandas as pd
import re
import openpyxl


#CAMINHO
local = 'C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas\\LISTAS - (01) JANEIRO\\equipamentos loja - DS 2196.xlsx'
local2 = local
workbooks = openpyxl.load_workbook(local)
workbooks_read = pd.read_excel(local2, header=1)


print('**********************************************************')
print('workbooks')
print(workbooks)

# PUXANDO VALOR JAVA
sheet = workbooks.active
cell  = sheet['C1']
valor_java = cell.value
valor_java_tratado = valor_java.split(":")
valor_tratado = valor_java_tratado[1]
print(valor_tratado)
print(type(valor_tratado))

print('**********************************************************')
print('workbooks_read')
print(workbooks_read)


#seleciona as colunas desejadas
DataFrame_Selecionado = workbooks_read.iloc[:, :7]
print(DataFrame_Selecionado)


print(DataFrame_Selecionado)
valor_java = valor_tratado
DataFrame_Selecionado = DataFrame_Selecionado.drop('GRUPO', axis=1)
DataFrame_Selecionado = DataFrame_Selecionado.rename(columns = {'NF': 'NF ENTRADA'})
DataFrame_Selecionado.insert(0, 'JAVA', valor_java)
DataFrame_Selecionado.insert(1, 'FILIAL', ' ')
DataFrame_Selecionado.insert(2, 'BANDEIRA   ', ' ')
DataFrame_Selecionado.insert(3, 'UF', ' ')
print(DataFrame_Selecionado)



excel = DataFrame_Selecionado.to_excel('teste_modelo.xlsx', index=False)