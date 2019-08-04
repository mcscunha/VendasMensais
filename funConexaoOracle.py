#!/usr/bin/env python3
# -*- coding: utf-8 -*-
'''
    Classe para conexao e manipulacao de dados do Oracle


Criador : MuriloCunha
Data    : 01/05/2019
'''


import cx_Oracle


class ConexaoOracle():
    def __init__(self, strUsuario, strSenha, strHost, strPorta, strSid):
        self.strUsuario = strUsuario
        self.strSenha = strSenha
        self.strHost = strHost
        self.strPorta = strPorta
        self.strSid = strSid
        self.conexao = None

    def conectarBD(self):
        # strConexao = connstr = uid + '/' + pwd + '@' + host + ':' + port + '/' + sid
        # connection = cx_Oracle.connect(strConexao) #cria a conexão
        dsn = cx_Oracle.makedsn(self.strHost, self.strPorta, self.strSid)
        self.conexao = cx_Oracle.connect(dsn=dsn, user=self.strUsuario, password=self.strSenha)
        return self.conexao


    def criarCursor(self):
        self.conectarBD()
        return self.conexao.cursor() # cria uma instancia de cursor


    def fecharConexao(self):
        self.conexao.close()
        print('[INFO] Conexao fechada com sucesso')


    def guardarCaracteristicasColunas(self, curCursor):
        dicTipoColunas = {'colnames': [],
                          'coltypes': [],
                          'coldisplay_sz': [],
                          'colinternal_sz': [],
                          'colprecision': [],
                          'colscale': [],
                          'colnullok': [] }
        if (curCursor is not None) and (isinstance (curCursor, cx_Oracle.Cursor)):
            [ [dicTipoColunas['colnames'].append(row[0]),
               dicTipoColunas['coltypes'].append(row[1]),
               dicTipoColunas['coldisplay_sz'].append(row[2]),
               dicTipoColunas['colinternal_sz'].append(row[3]),
               dicTipoColunas['colprecision'].append(row[4]),
               dicTipoColunas['colscale'].append(row[5]),
               dicTipoColunas['colnullok'].append(row[6]) ] for row in curCursor.description]
        print("[INFO] Nome(s) da(s) Coluna(s): ", dicTipoColunas['colnames'])
        print("[INFO] Tipo(s) das Coluna(s)  : ", dicTipoColunas['coltypes'])
        print("[INFO] Tamanho Visualização   : ", dicTipoColunas['coldisplay_sz'])
        print("[INFO] Tamanho Interno        : ", dicTipoColunas['colinternal_sz'])
        print("[INFO] Precisão da Coluna     : ", dicTipoColunas['colprecision'])
        print("[INFO] Escala da Coluna       : ", dicTipoColunas['colscale'])
        print("[INFO] Aceita NULO            : ", dicTipoColunas['colnullok'])
        print('[INFO] ATENÇÃO: Se estiver NULA as info acima, pode ser que a consulta retornou vazio')
        return dicTipoColunas

    def extrairNomeColunas(self, curCursor):
        return [lstColuna[0] for lstColuna in curCursor.description]


    def recuperarTodasLinhas(self, curCursor, strSql):
        linhas = []
        curCursor.execute(strSql)       # consulta sql
        linhas = curCursor.fetchall()   # busca o resultado da consulta
        return linhas

'''
Retirar uma LISTA dentro de outra:
    flat_list = []
    for sublist in l:
        for item in sublist:
            flat_list.append(item)

Outra forma mais otimizada:
    flatten = [item for sublist in l for item in sublist]
'''