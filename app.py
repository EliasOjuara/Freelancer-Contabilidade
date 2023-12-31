import openpyxl
import pyperclip

# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos'] 
# Copiar informação de um campo e colar no seu campo correspondente 
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    
    descricao = linha[1].value
    categoria = linha[2].value
    codigo_produto = linha[3].value
    peso = linha[4].value
    dimensoes = linha[5].value
    preco = linha[6].value
    quantidade_em_estoque = linha[7].value
    data_de_validade = linha[8].value
    cor = linha[9].value
    tamanho = linha[10].value
    material = linha[11].value
    fabricante = linha[12].value
    pais_origem = linha[13].value
    observacoes = linha[14].value
    codigo_de_barras = linha[15].value
    localizacao_armazem = linha[16].value

# Repetir esses passos para outros campos atée preencher os campos daquela pagina
# Clicar na proxima 
# Repetir os mesmos passos e ir para a proxima pagina (pagina 2)
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir 
# Clicar em Ok, para finalizar o processo 
# Clicar em Ok mais uma vez na mensagem de confirmação de salvamento do banco de dados 
# Clicar em "adicionar mais um e reetir o processo até finalizar a planilha"