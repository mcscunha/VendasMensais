select pcnfsaid.codfilial,
       trunc(pcnfsaid.dtfat) DATA_SAIDA,
       pcnfsaid.numnota,
       pcpedc.codemitente,
       (select nome from schippers.pcempr where pcpedc.codemitente = pcempr.matricula) as emitente,
       pcpedc.codusur,
       (select nome from schippers.pcusuari where pcpedc.codusur = pcusuari.codusur) as vendedor,
       pcnfsaid.codplpag,
       (select descricao
          from schippers.pcplpag
         where pcnfsaid.codplpag = pcplpag.codplpag) as planopagamento,
       (select pcativi.ramo
          from schippers.pcclient, schippers.pcativi
         where pcclient.codatv1 = pcativi.codativ
           and pcpedc.codcli = pcclient.codcli) as ramoativi,
       pcpedc.codcli,
       (select cliente from schippers.pcclient where pcpedc.codcli = pcclient.codcli) as cliente,
       (select municent from schippers.pcclient where pcpedc.codcli = pcclient.codcli) as cidade,
       (select estent from schippers.pcclient where pcpedc.codcli = pcclient.codcli) as estado,
       pcmov.codprod,
       (select codfab from schippers.pcprodut where pcmov.codprod = pcprodut.codprod) as codfabrica,
       pcmov.descricao,
       pcmov.embalagem,
       (select pcmarca.marca
          from schippers.pcprodut, schippers.pcmarca
         where pcprodut.codmarca = pcmarca.codmarca
           and pcmov.codprod = pcprodut.codprod) as marca,
       (select codepto from schippers.pcprodut where pcmov.codprod = pcprodut.codprod) as codepto,
       (select descricao
          from schippers.pcdepto
         where pcdepto.codepto =
               (select codepto
                  from schippers.pcprodut
                 where pcmov.codprod = pcprodut.codprod)) departamento,
       (select codsec from schippers.pcprodut where pcmov.codprod = pcprodut.codprod) as codsecao,
       (select descricao
          from schippers.pcsecao
         where pcsecao.codsec =
               (select codsec
                  from schippers.pcprodut
                 where pcmov.codprod = pcprodut.codprod)) secao,
       (select nvl(custofornec, 0)
          from schippers.pcest
         where pcnfsaid.codfilial = pcest.codfilial
           and pcmov.codprod = pcest.codprod) as custofornec,
       pcmov.qt,
       (select pcpraca.numregiao || '-' || pcregiao.regiao
          from schippers.pcpraca, schippers.pcclient, schippers.pcregiao
         where pcpraca.codpraca = pcclient.codpraca
           and pcpraca.numregiao = pcregiao.numregiao
           and pcclient.codcli = pcmov.codcli) NUMREGIAO_PRACA,
       (select pctabprcli.numregiao || '-' || pcregiao.regiao
          from schippers.pctabprcli, schippers.pcclient, schippers.pcregiao
         where pctabprcli.codcli = pcmov.codcli
           and pctabprcli.numregiao = pcregiao.numregiao
           and pctabprcli.codfilialnf in ('1')
         group by pctabprcli.codcli, pctabprcli.numregiao, pcregiao.regiao) NUMREGIAO_3314,
       (SELECT PCCLIENT.vip
          FROM schippers.PCCLIENT
         WHERE PCCLIENT.CODCLI = PCMOV.CODCLI) Classificacao_cliente,
       (select classevenda
          from schippers.pcclient
         where pcpedc.codcli = pcclient.codcli) as classevenda,
       (select qtcheckout
          from schippers.pcclient
         where pcpedc.codcli = pcclient.codcli) as qtcheckout,
       pcmov.punit,
       pcmov.punit * pcmov.qt as totalvenda,
       (pcmov.punit - (select nvl(custofornec, 0)
                         from schippers.pcest
                        where pcnfsaid.codfilial = pcest.codfilial
                          and pcmov.codprod = pcest.codprod
                          and pcest.codfilial in ('1'))) * pcmov.qt as Margem_total
  from schippers.pcmov, schippers.pcnfsaid, schippers.pcpedc
 where pcnfsaid.numtransvenda = pcmov.numtransvenda
   and pcnfsaid.numtransvenda = pcpedc.numtransvenda
   and pcpedc.posicao = 'F'
   and pcnfsaid.numnota = pcmov.numnota
   and pcnfsaid.condvenda in (1, 5)
   and pcnfsaid.codfilial in ('1')
   and to_date(pcnfsaid.dtfat, 'DD-MM-YY') between 
        to_date( '{varDataInicio}', 'DD/MM/YYYY') and 
        to_date( '{varDataFim}', 'DD/MM/YYYY')
   and pcpedc.codemitente = {varCodEmitente}
order by 
   totalvenda