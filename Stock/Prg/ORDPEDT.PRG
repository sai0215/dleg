CLOS DATA
report ='ordredt'
DO ordsedt.spr
CLOS DATA

PROCEDURE vblank
****************
= checkzapfile("ordtmp")
SELECT (alias)
m.supdes   = ''
m.saldes   = ''
m.bandes   = ''
m.curdes1  = ''
m.curdes2  = ''
m.bandes   = ''
m.lcurdes  = ''
m.totprice = 0
m.bnktotal = 0
m.curtotal = 0
SCATTER MEMVAR BLANK
IF isadding
   torder = ORDER()
   trecno = RECNO()
   SET ORDER TO 0
   GO BOTTOM
   m.serial = serial+1
   SET ORDER TO &torder
   IF trecno > 0 AND trecno <= RECCOUNT()
      GO RECORD trecno
   ENDIF   
   IF isadding
      m.date = DATE()
   ENDIF   
ENDIF
SELECT ordtmp
APPEND BLANK
REPLACE ordhea   WITH m.serial
REPLACE ordtype  WITH 'N'
SHOW GETS
SELECT (alias)

PROCEDURE getdata
*****************
m.supdes = suptab.name
SELECT curtab
LOCATE FOR serial = ordhea.curtab1
m.curdes1 = curtab.number
m.bandes  = bantab.name
SELECT curtab
LOCATE FOR serial = ordhea.curtab2
m.curdes2 = curtab.number
m.lcurdes = curtab.number
m.saldes = saltab.name
SELECT (alias)
SCATTER MEMVAR
ttserial = m.serial
=checkzapfile("ordtmp")
SELECT ordrow
SEEK ordhea.serial
DO WHILE ordhea = ordhea.serial AND !EOF()
   SCATTER MEMVAR
   SELECT ordtmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE ordhea   WITH m.serial
   REPLACE name     WITH protab.name
   REPLACE des      WITH protab.number
   REPLACE alldesc  WITH IIF(!EMPTY(protab.field2),'/'+ALLTRIM(protab.field2),'')+IIF(!EMPTY(fl3tab.name),'/'+ALLTRIM(fl3tab.name),'')+IIF(!EMPTY(fl4tab.name),'/'+ALLTRIM(fl4tab.name),'')
   REPLACE ordtype  WITH 'U'
   REPLACE ordrecno WITH RECNO('ordrow')
   SELECT ordrow
   SKIP
ENDDO
SELECT ordtmp
APPEND BLANK
REPLACE ordhea  WITH m.serial
REPLACE ordtype WITH 'N'

=checkzapfile("oextmp")
SELECT exttab
SET FILTER TO dest='O' OR dest='B'
GO TOP
DO WHILE !EOF()
	SCATTER MEMVAR
	SELECT oextmp
	APPEND BLANK
	GATHER MEMVAR
	REPLACE exttab WITH m.serial
	REPLACE curr   WITH curtab.number
	SELECT exttab
	SKIP
ENDDO
SELECT oextmp
REPLACE ordhea WITH ordhea.serial ALL
SELECT oextab
SEEK ordhea.serial
DO WHILE ordhea=ordhea.serial AND !EOF()
	SCATTER MEMVAR
	wrec = RECNO()
	SELECT oextmp
	LOCATE FOR ordhea=ordhea.serial AND exttab=oextab.exttab
	IF FOUND()
		REPLACE wrec    WITH m.wrec
		REPLACE price   WITH oextab.price
		REPLACE curr    WITH curtab.number
		REPLACE curprix WITH oextab.curprix
		REPLACE choix   WITH .T.
	ENDIF
	SELECT oextab
	SKIP
ENDDO
SELECT oextmp
SUM curprix TO m.curtotal
DO caldisc
SELECT ordtmp
m.serial = ttserial
SHOW GETS
DO getprdt
SELECT (alias)
m.serial = ttserial

PROCEDURE savedata
******************
SELECT (alias)
IF isadding
   torder = ORDER()
   SET ORDER TO 0
   GO BOTTOM
   m.serial = serial+1
   SET ORDER TO &torder
   APPEND BLANK
ENDIF
SET DELETED OFF
SELECT ordtmp
GO TOP
DO WHILE !EOF()
   DO CASE
      CASE ordtype = 'A'
           SCATTER MEMVAR
           SELECT ordrow
           APPEND BLANK
           GATHER MEMVAR
           REPLACE ordhea WITH m.serial
      CASE ordtype = 'U'
           IF ordrecno <> 0
              SCATTER MEMVAR
              SELECT ordrow
              GO RECORD m.ordrecno
              GATHER MEMVAR
              REPLACE ordhea WITH m.serial
           ENDIF
      CASE ordtype = 'D'
           IF ordrecno <> 0
              SCATTER MEMVAR
              SELECT ordrow
              GO RECORD m.ordrecno
              DELETE
           ENDIF
   ENDCASE
   SELECT ordtmp
   REPLACE ordtype WITH IIF(ordtype = 'A' , 'U', ordtype )
   SKIP        
ENDDO
SET DELETED ON
SELECT (alias)
GATHER MEMVAR
isadding = .F.

PROCEDURE search
*****************
SELECT (alias)
m.sfound = .F.
m.srecno = 0
IF issearch
   m.srecno = RECNO()
ENDIF
filesearch = "suptab"
DO ordsearch.spr
IF m.srecno<>0 AND !m.sfound
   GO TOP
   IF NOT EOF()
      GO RECORD m.srecno
   ENDIF   
ENDIF
IF m.sfound
   m.issearch = .T.
   DO getdata
ENDIF

PROCEDURE okdel
***************
SELECT ordrow
DELETE FOR ordhea = m.serial
SELECT (alias)
RETURN .T.

PROCEDURE prtrec
****************
SELECT (alias)
DO print WITH report,IIF(m.criteria=1,'RECNO()=m.saverec',IIF(m.criteria=2,'NOT EMPTY(bidon)','.T.')),IIF(m.print=1,'PREVIEW','TO PRINT')

PROCEDURE caldisc
*****************
SELECT ordtmp
m.ttordrec = RECNO()
GO TOP
SUM price TO m.totprice
m.totprice = (m.totprice*(100-m.discount))/100
IF .F.
SELECT rathea 
SET ORDER TO date
GO BOTTOM
tserial = serial
SELECT ratrow
SET FILTER TO rathea=tserial
SELECT curtab
LOCATE FOR number=m.curdes1
m.ttotal = m.totprice*ratrow.drate1
SELECT curtab
LOCATE FOR number=m.curdes2
m.bnktotal = m.ttotal/ratrow.drate1
SELECT ratrow
SET FILTER TO
ENDIF
m.bnktotal = cal_cur(m.totprice,m.date,m.curtab1,m.curtab2)
SHOW GET m.bnktotal
SELECT oextmp
SUM curprix TO m.curtotal

SELECT ordtmp
IF m.ttordrec > 0 AND m.ttordrec <=RECCOUNT()
   GO RECORD m.ttordrec
ENDIF

PROCEDURE getprdt
*****************
=checkzapfile("protmp1")
SELECT protab1
SET FILTER TO suptab=m.suptab
GO TOP
DO WHILE !EOF()
	m.price  = price
	m.curtab = curtab
	SELECT protab
	m.serial = serial
	m.name   = name
	m.number = number
	SELECT protmp1
	APPEND BLANK
	GATHER MEMVAR
	SELECT protab1
	SKIP
ENDDO

PROCEDURE prtdata
*****************
m.out = 1
SELECT (alias)
wrec    = RECNO()
m.okbtn = 1
DO printrec.spr
IF m.okbtn=2
	RETURN
ENDIF	
SELECT ordtmp
*SET ORDER TO iprotab
SET ORDER TO des
IF m.out=1
	REPORT FORM ordredt2 FOR ordtype <> 'N' PREVIEW
ELSE
	REPORT FORM ordredt2 FOR ordtype <> 'N' TO PRINT OFF
ENDIF	
SELECT ordtmp
SET ORDER TO 0
SELECT (alias)
IF wrec > 0 AND wrec <= RECCOUNT()
   GO RECORD wrec
ENDIF   
RETURN	

PROCEDURE somme
***************
tselect = SELECT()
SELECT ordtmp
SUM qty TO n
SELECT (tselect)
RETURN n