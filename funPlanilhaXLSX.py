# -*- coding: utf-8 -*-
'''
    Arquivo com funcoes para trabalhar com planilhas XLSX


Criador : MuriloCunha
Data    : 04/08/2019

LOG de alteracoes:
-------------------------------------------------
04/08/2019- Criacao deste arquivo

-------------------------------------------------
'''

from datetime import date, datetime
import os
from openpyxl import Workbook
from funArquivoTexto import ArquivoXlsx


def GravarResultadoEmPlanilha(dicDadosAGravar, lstCabecalho):
    '''
        Gravacao dos dados no XLS
    '''
    # Acrescentar os dados recuperados do banco na planilha
    for vendedor, linhas in dicDadosAGravar.items():
        # Criando a pasta de trabalho
        book = Workbook()

        # Ativando a planilha da pasta de trabalho
        sheet = book.active

        # Acrescentar, na primeira linha, o cabecalho de colunas
        sheet.append(lstCabecalho)
        
        c = 1
        for linha in dicDadosAGravar[vendedor]:
            print('[INFO] Vendedor: {}\t-Inserindo linha: {}'.format(
                vendedor, c))
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

        #print(agora.year, '-', agora.month, '-', agora.day, ' | ',
        #      agora.hour, ':', agora.minute, ':', agora.second)
        datAgora = datetime.now()
        strData = datAgora.strftime('%Y%m%d')
        strHora = datAgora.strftime('%H%M%S')
        strArqXlsx = 'Cubo_8003_{}_{}_{}.xlsx'.format(
            vendedor, strData, strHora)

        #exit(0)

        # Criar classe para manipulacao de arquivos XLSX
        arqXlsx = ArquivoXlsx('./Resultado', strArqXlsx)
        # Criar diretorio se nao existir
        arqXlsx.PrepararDiretorioParaGravacao()
        # Apagar arquivo se este existir
        arqXlsx.ApagarArquivoExistente()
        
        strCaminhoCompleto = os.path.join('./Resultado', strArqXlsx)
        # Gravando o arquivo XLS
        book.save(strCaminhoCompleto)
    # FIM - Gravacao do XLS
    #

    print('Arquivo(s) criado(s) com sucesso:', strArqXlsx, '\n')

