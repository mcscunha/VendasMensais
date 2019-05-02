'''
    Classe para conexao e manipulacao de dados do Oracle


Criador : MuriloCunha
Data    : 01/05/2019
'''
#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import cx_Oracle


class ConexaoOracle():
    def __init__(self, strUsuario, strSenha, strHost, strPorta, strSid):
        self.strUsuario = strUsuario
        self.strSenha = strSenha
        self.strHost = strHost
        self.strPorta = strPorta
        self.strSid = strSid
    
    def conectarBD(self):
        #strConexao = connstr = uid + '/' + pwd + '@' + host + ':' + port + '/' + sid
        #connection = cx_Oracle.connect(strConexao) #cria a conexão
        dsn = cx_Oracle.makedsn(self.strHost, self.strPorta, self.strSid)
        return cx_Oracle.connect(dsn=dsn, user=self.strUsuario, password=self.strSenha)

    def enviarSelect(self, strSql):
        linhas = []
        conexao = self.conectarBD()
        cur = conexao.cursor()      # cria um cursor
        cur.execute(strSql)         # consulta sql
        linhas.append(cur.description)
        linhas.append(cur.fetchall())        # busca o resultado da consulta
        conexao.close()
        return linhas