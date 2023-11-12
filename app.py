import openpyxl
# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos'] 
# Copiar informação de um campo e colar no seu campo correspondente 
for linha in sheet_produtos.iter_rows(min_row=2):
    print(linha[0].value)
# Repetir esses passos para outros campos atée preencher os campos daquela pagina
# Clicar na proxima 
# Repetir os mesmos passos e ir para a proxima pagina (pagina 2)
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir 
# Clicar em Ok, para finalizar o processo 
# Clicar em Ok mais uma vez na mensagem de confirmação de salvamento do banco de dados 
# Clicar em "adicionar mais um e reetir o processo até finalizar a planilha"