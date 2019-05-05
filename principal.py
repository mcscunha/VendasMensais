'''
    Arquivo principal para execucao do sistema de construcao
de planilhas resumo de vendas do mes


Criador : MuriloCunha
Data    : 01/05/2019
'''
#!/usr/bin/env python3
# -*- coding: utf-8 -*-


from funConexaoOracle import ConexaoOracle
from funArquivoTexto import ArquivoTexto
from dados_confidenciais import dicConexaoOracle

HOST = dicConexaoOracle['host']
PORTA = dicConexaoOracle['port']
USUARIO = dicConexaoOracle['user']
SENHA = dicConexaoOracle['pass']
SID = dicConexaoOracle['sid']
ARQ_SQL = 'Cubo_8003.sql'
DATA_INICIO = '01/04/2019'
DATA_FIM = '02/04/2019'
CODEMITENTE = { 'colunas': [],   # Vendedores
                46: [],      
                8888: [] }

# Recuperar o codigo SQL gravado no arquivo .sql
arqSql =  ArquivoTexto(ARQ_SQL, 'r', 'utf-8')
strCubo = arqSql.carregarTexto()

# Fazer conexao com o banco de dados e recuperar os dados
conOracle = ConexaoOracle(USUARIO, SENHA, HOST, PORTA, SID)
curCursor = conOracle.criarCursor()

for vendedor in CODEMITENTE.keys():
    strCubo = strCubo.format(varDataInicio=DATA_INICIO,
                            varDataFim=DATA_FIM,
                            varCodEmitente=CODEMITENTE)
    linhas = conOracle.recuperarTodasLinhas(curCursor, strCubo)
    if CODEMITENTE.get('colunas').count == 0:
        CODEMITENTE['colunas'] = conOracle.extrairNomeColunas(curCursor)
    oraColTipos = conOracle.guardarCaracteristicasColunas(curCursor)

curCursor.close()
conOracle.fecharConexao()

# if linhas == None: 
#     print("Nenhum Resultado")
#     exit
# else:
#     for linha in linhas:
#         print(linha)

# Gravacao dos dados no XLS
from openpyxl import Workbook

# Criando a pasta de trabalho
book = Workbook()
# Ativando a planilha da pasta de trabalho
sheet = book.active

# Acrescentar, na primeira linha, o cabecalho de colunas
sheet.append(cabecalho)

# Acrescentar os dados recuperados do banco na planilha
for linha in linhas:
    for c, celula in enumerate(linha):
        print('\tInserindo linha:', c+1, '\tTipo em analise:', oraColTipos['coltypes'][c])
        #
        # A linha abaixo insere EXATAMENTE como o dado se apresenta no PRINT
        # Se for um NUMBER e o separador de milhar for "." (ponto), a planilha
        # nao fará os calculos como numericos, apenas como STRING.
        # sheet.cell(row=idx+2, column=c+2).value = str(celula)
        
        #
        # Na linha abaixo, a biblioteca se encarrega de formatar o tipo correto
        # para cada coluna (STRING, NUMBER, DATE...) de acordo com o dado.
        # Entao, esta forma é preferida para inserir varias info sem se preocupar
        # com os tipos
        sheet.append(celula)
book.save('cubo.xlsx')
# FIM - Gravacao do XLS

print('\n')
import csv
with open('cubo.csv', 'w', newline='') as arq_csv:
    c = csv.writer(arq_csv, delimiter=';')
    c.writerows(linhas)

print('Arquivo criado com sucesso')



