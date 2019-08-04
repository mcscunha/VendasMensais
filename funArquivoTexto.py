#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'''
    Classe para manipulacao de arquivos texto


Criador : MuriloCunha
Data    : 01/05/2019
'''


import os


class ArquivoTexto():
    def __init__(self, strArquivo, strModo, strEncoding):
        self.strArquivo = strArquivo
        self.strModo = strModo
        self.strEncoding = strEncoding

    def verificaExistenciaArquivo(self):
        if not self.strModo in ('w', 'W'):
            if not os.path.exists(self.strArquivo):
                return 0
        return 1
        
    def carregarTexto(self):
        if self.verificaExistenciaArquivo():
            with open(file=self.strArquivo,
                      mode=self.strModo,
                      encoding=self.strEncoding) as arq:
                return arq.read()
        else:
            print('Arquivo não encontrado.')
            return []


class ArquivoXlsx():
    def __init__(self, strDiretorio, strNomeArquivoXlsx):
        self.strDiretorio = strDiretorio
        self.strNomeArquivo = strNomeArquivoXlsx

    def PrepararDiretorioParaGravacao(self):
        if not os.path.exists(self.strDiretorio):
            os.makedirs(self.strDiretorio)
    
    def ApagarArquivoExistente(self):
        strCaminhoCompleto = os.path.join(self.strDiretorio, self.strNomeArquivo)
        if os.path.isfile(strCaminhoCompleto):
            os.remove(strCaminhoCompleto)
