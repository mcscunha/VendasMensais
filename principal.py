'''
    Arquivo principal para execucao do sistema de construcao
de planilhas resumo de vendas do mes


Criador : MuriloCunha
Data    : 01/05/2019
'''
# -*- coding: utf-8 -*-


from datetime import date, datetime, timezone
import os
from openpyxl import Workbook
from funConexaoOracle import ConexaoOracle
from funArquivoTexto import ArquivoTexto
from dados_confidenciais import dicConexaoOracle


# DADOS A SEREM ALTERADOS PELO USUARIO
VENDEDORES   = [40]
bolFiltroNot = False         # True = NOT IN  |  False = IN
lstEstado    = ['SP', 'MG']
datInicio    = '01/05/2019'
datFim       = '31/05/2019'
# FIM DOS DADOS ALTERADOS PELO USUARIO


# Constantes de conexao com o banco
HOST = dicConexaoOracle['host']
PORTA = dicConexaoOracle['port']
USUARIO = dicConexaoOracle['user']
SENHA = dicConexaoOracle['pass']
SID = dicConexaoOracle['sid']
ARQ_SQL = 'Cubo_8003.sql'


dicCodUsur = {}
lstCabecalho = []

# Recuperar o codigo SQL gravado no arquivo .sql
arqSql =  ArquivoTexto(ARQ_SQL, 'r', 'utf-8')
strCubo = arqSql.carregarTexto()

# Fazer conexao com o banco de dados
conOracle = ConexaoOracle(USUARIO, SENHA, HOST, PORTA, SID)
curCursor = conOracle.criarCursor()

# Recuperar info dos vendedores e alimentar o dicionario
for vendedor in VENDEDORES:
    print('VENDEDOR:', vendedor)

    if bolFiltroNot:
        strEstado = "and pcclient.estent not in (" + str(lstEstado)[1:-1] + ")"
        indice = 0
    else:
        if len(lstEstado) > 0:
            for indice, estado in enumerate(lstEstado):
                strEstado = "and pcclient.estent in ('" + estado + "')"
        else:
            strEstado = ''

    strSql = strCubo.format(varDataInicio=datInicio,
                            varDataFim=datFim,
                            varCodUsur=vendedor,
                            varFiltroEstado=strEstado)

    print(strSql)
    exit(0)

    # Nao trocar esta ordem de execucao
    # Primeiro deve-se recuperar as linhas e depois o cabecalho
    strItemDic = str(vendedor) + '-' + str(indice)
    dicCodUsur[strItemDic] = conOracle.recuperarTodasLinhas(curCursor, strSql)

    if len(lstCabecalho) == 0:
        # Nao usar APPEND para acrescentar somente uma vez E dados tipo LISTA
        # O APPEND cria uma LISTA DENTRO DE OUTRA
        lstCabecalho = conOracle.extrairNomeColunas(curCursor)

# Fechar a conexao
curCursor.close()
conOracle.fecharConexao()

#
# Gravacao dos dados no XLS

# Acrescentar os dados recuperados do banco na planilha
for vendedor, linhas in dicCodUsur.items():
    # Criando a pasta de trabalho
    book = Workbook()

    # Ativando a planilha da pasta de trabalho
    sheet = book.active

    # Acrescentar, na primeira linha, o cabecalho de colunas
    sheet.append(lstCabecalho)

    c = 1
    for linha in dicCodUsur[vendedor]:
        print('Vendedor: {}\t-Inserindo linha: {}'.format(vendedor, c))
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
        sheet.append(linha)
        c += 1

    # Gravando o arquivo XLS
    if not os.path.exists('./Resultado'):
        os.makedirs('./Resultado')
    if os.path.isfile('./Resultado/Cubo_8003_{}.xlsx'.format(vendedor)):
        os.remove('./Resultado/Cubo_8003_{}.xlsx'.format(vendedor))
    
    datAgora = datetime.now()
    strData = datAgora.strftime('%Y%m%d')
    strHora = datAgora.strftime('%H%M%S')
    #print(agora.year, '-', agora.month, '-', agora.day, ' | ',
    #      agora.hour, ':', agora.minute, ':', agora.second)
    book.save('./Resultado/Cubo_8003_{}_{}_{}.xlsx'.format(
        vendedor, strData, strHora))
# FIM - Gravacao do XLS
#

print('Arquivo(s) criado(s) com sucesso')
print('Encerrando a aplicação')
exit(0)
