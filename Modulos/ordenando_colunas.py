import pandas as pd

arquivo_excel = pd.read_excel('C:\\Users\\jfgfilho\\OneDrive - rd.com.br\\Área de Trabalho\\extracao_dados\\planilhas\\LISTAS - (01) JANEIRO\\equipamentos loja - DS 2196.xlsx', header=1)


#seleciona as colunas desejadas
DataFrame_Selecionado = arquivo_excel.iloc[:, :7]

print(DataFrame_Selecionado)
def ordenando_as_colunas():
    valor_java = 1234
    DataFrame_Selecionado = DataFrame_Selecionado.drop('GRUPO', axis=1)
    DataFrame_Selecionado = DataFrame_Selecionado.rename(columns = {'NF': 'NF ENTRADA'})
    DataFrame_Selecionado.insert(0, 'JAVA', valor_java)
    DataFrame_Selecionado.insert(1, 'FILIAL', ' ')
    DataFrame_Selecionado.insert(2, 'BANDEIRA   ', ' ')
    DataFrame_Selecionado.insert(3, 'UF', ' ')
    print(DataFrame_Selecionado)



excel = DataFrame_Selecionado.to_excel('teste_modelo.xlsx', index=False)
