# -*- coding: utf-8 -*-
'''
    Arquivo principal para execucao do sistema de construcao
de planilhas resumo de vendas do mes


Criador : MuriloCunha
Data    : 01/05/2019

LOG de alteracoes:
-------------------------------------------------
04/08/2019- Trocando variavel VENDEDORES para uma lista de listas para fazer todo o 
            processo de exportacao de planilhas em uma execucao

-------------------------------------------------
'''


from funArquivoTexto import ArquivoTexto
from funConexaoOracle import ConexaoOracle
from funPlanilhaXLSX import GravarResultadoEmPlanilha
from dados_confidenciais import dicConexaoOracle


# DADOS A SEREM ALTERADOS PELO USUARIO
VENDEDORES = [
    # [ ID_Vendedor, bolFiltroNot ] , [Estados]
    # bolFiltroNot ==>> False = IN (inclusao do estado)
    #                   True = NOT IN (exclusao dos estados)
    [ [3, False], ['MG'] ],
    # [ [3, False], ['SP'] ],
    # [ [3, True], ['MG', 'SP'] ],
    # [ [10, False], [] ],
    # [ [13, False], [] ],
    # [ [16, False], [] ],
    # [ [19, False], ['MS'] ],
    # [ [19, True], ['MS'] ],
    # [ [22, False], ['SP'] ],
    # [ [22, False], ['MG'] ],
    # [ [22, True], ['SP', 'MG'] ],
    # [ [24, False], [] ],
    # [ [34, False], ['PR'] ],
    # [ [34, False], ['MS'] ],
    # [ [34, True], ['PR', 'MS'] ],
    # [ [35, False], [] ],
    # [ [38, False], [] ],
    # [ [40, False], [] ],
]
datInicio    = '01/07/2019'
datFim       = '31/07/2019'
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
    intVendedor = vendedor[0][0]
    lstEstado = vendedor[1]
    bolFiltroNot = vendedor[0][1]

    print('VENDEDOR:', vendedor[0][0])

    if len(lstEstado) > 0:
        strEstado = "and pcclient.estent "
        if bolFiltroNot:
            strEstado += "not in (" 
            print('[INFO] Usando NOT IN - Estados NAO procurados')
        else:
            strEstado += "in ("
            print('[INFO] Usando IN - Estado PROCURADO:', str(lstEstado)[1:-1])
        
        strEstado += str(lstEstado)[1:-1] + ")"
        # strEstado += str(lstEstado) + ")"
    else:
        strEstado = ''

    strSql = strCubo.format(varDataInicio=datInicio,
                            varDataFim=datFim,
                            varCodUsur=intVendedor,
                            varFiltroEstado=strEstado)

    # Nao trocar esta ordem de execucao
    # Primeiro deve-se recuperar as linhas e depois o cabecalho
    strItemDic = str(intVendedor) + '-' + \
                 ('Exceto_' if bolFiltroNot else '') + \
                 str(lstEstado)[1:-1].replace("'", "").replace(', ', '_')
    
    
    print(strSql)
    dicCodUsur[strItemDic] = conOracle.recuperarTodasLinhas(curCursor, strSql)

    if len(lstCabecalho) == 0:
        # Nao usar APPEND para acrescentar somente uma vez E dados tipo LISTA
        # O APPEND cria uma LISTA DENTRO DE OUTRA
        lstCabecalho = conOracle.extrairNomeColunas(curCursor)

    GravarResultadoEmPlanilha(dicCodUsur[strItemDic], lstCabecalho, strItemDic)
    # Limpar da memoria o conteudo do dicionario usado
    del dicCodUsur[strItemDic]
    
# Fechar a conexao
curCursor.close()
conOracle.fecharConexao()

print('\nEncerrando a aplicação\n')
