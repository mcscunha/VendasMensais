select 
   trunc(pcnfsaid.dtfat) DTFAT,
   (select nome from schippers.pcempr where pcpedc.codemitente = pcempr.matricula) as EMITENTE,
   (select nome from schippers.pcusuari where pcpedc.codusur = pcusuari.codusur) as VENDEDOR,
   (select descricao
         from schippers.pcplpag
      where pcnfsaid.codplpag = pcplpag.codplpag) as PLANOPAGAMENTO,
   pcpedc.codcli,
   (select cliente from schippers.pcclient where pcpedc.codcli = pcclient.codcli) as cliente,
   (select municent from schippers.pcclient where pcpedc.codcli = pcclient.codcli) as cidade,
   (select estent from schippers.pcclient where pcpedc.codcli = pcclient.codcli) as estado,
   pcmov.descricao,
   pcmov.qt,
   pcmov.punit,
   pcmov.punit * pcmov.qt as totalvenda,
   (select nvl(custofornec, 0)
         from schippers.pcest
      where pcnfsaid.codfilial = pcest.codfilial
         and pcmov.codprod = pcest.codprod) as custofornec,
   pcnfsaid.numnota,
   (select descricao
         from schippers.pcdepto
      where pcdepto.codepto =
            (select codepto
               from schippers.pcprodut
               where pcmov.codprod = pcprodut.codprod)) departamento
from schippers.pcmov
join schippers.pcnfsaid 
   on pcnfsaid.numtransvenda = pcmov.numtransvenda
   and pcnfsaid.numnota = pcmov.numnota
join schippers.pcpedc
   on pcnfsaid.numtransvenda = pcpedc.numtransvenda
join schippers.pcclient
   on pcpedc.codcli = pcclient.codcli
where 
   pcpedc.posicao = 'F'
   and pcnfsaid.condvenda in (1, 5)
   and pcnfsaid.codfilial in ('1')
   and to_date(pcnfsaid.dtfat, 'DD-MM-YY') between 
      to_date( '{varDataInicio}', 'DD/MM/YYYY') and 
      to_date( '{varDataFim}', 'DD/MM/YYYY')
   and pcpedc.codusur = {varCodUsur}
   {varFiltroEstado}
   
order by 
   totalvenda