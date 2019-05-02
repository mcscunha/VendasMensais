'''
    Classe para manipulacao de arquivos texto


Criador : MuriloCunha
Data    : 01/05/2019
'''
#!/usr/bin/env python3
# -*- coding: utf-8 -*-


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