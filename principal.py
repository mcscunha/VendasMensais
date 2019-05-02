'''
    Arquivo principal para execucao do sistema de construcao
de planilhas resumo de vendas do mes


Criador : MuriloCunha
Data    : 01/05/2019
'''
#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import xlsxwriter
from funConexaoOracle import ConexaoOracle
from funArquivoTexto import ArquivoTexto


HOST    = '179.190.39.14'
PORTA   = '45826'
USUARIO = 'schippers'
SENHA   = 'SCHIPPESPRD1711'
SID     = 'SCHP5102'
ARQ_SQL = 'Cubo_8003.sql'

datInicio = '01/04/2019'
datFim = '01/04/2019'

arqSql =  ArquivoTexto(ARQ_SQL, 'r', 'utf-8')
strCubo = arqSql.carregarTexto()
strCubo = strCubo.format(varDataInicio=datInicio, varDataFim=datFim)

conOracle = ConexaoOracle(USUARIO, SENHA, HOST, PORTA, SID)
linhas = conOracle.enviarSelect(strCubo)



# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Expenses01.xlsx')
worksheet = workbook.add_worksheet()

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for linha in linhas:
    worksheet.write_row(row=row, col=col, data=linha)
    row += 1

# Write a total using a formula.
worksheet.write(row, 0, 'Total')
worksheet.write(row, 1, '=SOMA(A1:A5)')

workbook.close()
print('Arquivo criado com sucesso')

if linhas == None: 
    print("Nenhum Resultado")
    exit
else:
    for linha in linhas:
        print(linha)



