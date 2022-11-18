CLOS DATA
report ='arrredt'
report1 ='arrredt1'
DO arrsedt.spr
CLOS DATA

PROCEDURE vblank
****************
=checkzapfile("sertmp1")
=checkzapfile("arrtmp")
SELECT (alias)
m.wardes = ''
m.orddes = ''
m.supdes  = ''
m.curdes  = ''
m.curdes1 = ''
m.curdes2 = ''
m.curdes3 = ''
m.agedes1 = ''
m.agedes2 = ''
m.agedes3 = ''
m.totprice = 0
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
SELECT arrtmp
APPEND BLANK
REPLACE arrhea   WITH m.serial
REPLACE arrtype  WITH 'N'
SHOW GETS
SELECT (alias)

PROCEDURE getdata
*****************
m.orddes = ordhea.ref_no
m.supdes = suptab.name
m.wardes = warhea.name
SELECT curtab
LOCATE FOR serial = arrhea.curtab
m.curdes  = number
LOCATE FOR serial = arrhea.curtab1
m.curdes1 = number
LOCATE FOR serial = arrhea.curtab2
m.curdes2 = number
LOCATE FOR serial = arrhea.curtab3
m.curdes3 = number
SELECT agetab
LOCATE FOR serial = arrhea.agetab1
m.agedes1 = name
LOCATE FOR serial = arrhea.agetab2
m.agedes2 = name
LOCATE FOR serial = arrhea.agetab3
m.agedes3 = name
SELECT (alias)
SCATTER MEMVAR
ttserial = m.serial
m.lfcost   = m.fcost
m.lfcharge = m.fcharge
m.lffees   = m.ffees
m.lcost    = m.cost
m.lcharge  = m.charge
m.lfees    = m.fees
m.lcurtab1 = m.curtab1
m.lcurtab2 = m.curtab2
m.lcurtab3 = m.curtab3
SELECT arrtmp
DO fillrow WITH .T.
DO caldisc
SELECT arrtmp
SHOW GETS

=checkzapfile("aextmp")
SELECT exttab
SET FILTER TO dest='A' OR dest='B'
GO TOP
DO WHILE !EOF()
	SCATTER MEMVAR
	SELECT aextmp
	APPEND BLANK
	GATHER MEMVAR
	REPLACE exttab WITH m.serial
	REPLACE curr   WITH curtab.number
	SELECT exttab
	SKIP
ENDDO
SELECT aextmp
REPLACE arrhea WITH arrhea.serial ALL
SELECT aextab
SEEK arrhea.serial
DO WHILE arrhea=arrhea.serial AND !EOF()
	SCATTER MEMVAR
	wrec = RECNO()
	SELECT aextmp
	LOCATE FOR arrhea=arrhea.serial AND exttab=aextab.exttab
	IF FOUND()
		REPLACE wrec    WITH m.wrec
		REPLACE price   WITH aextab.price
		REPLACE curr    WITH curtab.number
		REPLACE curprix WITH aextab.curprix
		REPLACE choix   WITH .T.
	ENDIF
	SELECT aextab
	SKIP
ENDDO
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
SELECT (alias)
GATHER MEMVAR
SELECT arrrow
SEEK arrhea.serial
DO WHILE arrhea = arrhea.serial AND !EOF()
   SELECT sertab
   DELETE FOR serial = RECNO('arrrow')
   SELECT arrrow
   SKIP
ENDDO
SET DELETED OFF
SELECT arrtmp
GO TOP
DO WHILE !EOF()
   DO CASE
      CASE arrtype = 'A'
           SCATTER MEMVAR
           SELECT arrrow
           APPEND BLANK
           GATHER MEMVAR
           REPLACE arrhea WITH m.serial
           SELECT sertmp1
           SEEK RECNO('arrtmp')
           DO WHILE serial = RECNO('arrtmp') AND !EOF()
              SELECT sertab
              APPEND BLANK
              REPLACE serial  WITH RECNO('arrrow')
              REPLACE ser_nbr WITH sertmp1.ser_nbr
              SELECT sertmp1
              SKIP
           ENDDO
           SELECT arrtmp
      CASE arrtype = 'U'
           IF arrrecno <> 0
              IF qty = 0
                 SCATTER MEMVAR
                 SELECT arrrow
                 GO RECORD m.arrrecno
                 DELETE
                 SELECT arrtmp
              ELSE
                 SCATTER MEMVAR
                 SELECT arrrow
                 GO RECORD m.arrrecno
                 GATHER MEMVAR
                 REPLACE arrhea WITH m.serial
                 SELECT sertmp1
                 SEEK RECNO('arrtmp')
                 DO WHILE serial = RECNO('arrtmp') AND !EOF()
                    SELECT sertab
                    APPEND BLANK
                    REPLACE serial  WITH RECNO('arrrow')
                    REPLACE ser_nbr WITH sertmp1.ser_nbr
                    SELECT sertmp1
                    SKIP
                 ENDDO
                 SELECT arrtmp
              ENDIF   
           ENDIF
      CASE arrtype = 'D'
           IF arrrecno <> 0
              SCATTER MEMVAR
              SELECT arrrow
              GO RECORD m.arrrecno
              DELETE
           ENDIF
   ENDCASE
   SELECT arrtmp
   REPLACE arrtype WITH IIF(arrtype = 'A' , 'U', arrtype )
   SKIP
ENDDO
SET DELETED ON
SELECT (alias)
isadding = .F.

PROCEDURE search
*****************
SELECT (alias)
m.sfound = .F.
m.srecno = 0
IF issearch
   m.srecno = RECNO()
ENDIF
m.filesearch = 'suptab'
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
SELECT arrrow
DELETE FOR arrhea = m.serial
SELECT (alias)
RETURN .T.


PROCEDURE caldisc
*****************
SELECT arrtmp
m.ttarrrec = RECNO()
GO TOP
REPLACE ALL perc1 WITH 0
REPLACE ALL perc2 WITH 0
SUM weight TO m.ttotweight
SUM price  TO m.ttotprice
SUM price  TO m.totprice
m.totprice = m.totprice + m.fcost+m.fcharge+m.ffees
GO TOP
DO WHILE !EOF()
   m.value1 = 0
   m.value2 = 0
   IF m.ttotweight <> 0 AND m.fcost <> 0 AND price <> 0
      m.value1=(m.fcost*weight)/ttotweight
      REPLACE perc1 WITH (m.value1*100)/price
   ENDIF   
   IF m.ttotprice <> 0 AND m.fcharge+m.ffees <> 0 AND price <> 0
      m.value2=((m.fcharge+m.ffees)*price)/ttotprice
      REPLACE perc2 WITH (ttotprice)/(m.fcharge+m.ffees)
   ENDIF   
   IF qty <> 0
      REPLACE fprice WITH (price+m.value1+m.value2)/qty
   ELSE   
      REPLACE fprice WITH 0
   ENDIF   
   SKIP
ENDDO
SELECT arrtmp
IF m.ttarrrec > 0 AND m.ttarrrec <=RECCOUNT()
   GO RECORD m.ttarrrec
ENDIF   


PROCEDURE prtrec
***************
SELECT prttab
DELETE FOR EMPTY(bidon)
SELECT (alias)
SET RELATION TO serial INTO Prttab 
DO print WITH IIF(m.criteria=1,report,report1),IIF(m.criteria=1,'RECNO()=m.saverec',IIF(m.criteria=2,'FOUND("prttab")','.T.')),IIF(m.print=1,'PREVIEW','TO PRINT')
SELECT prttab
USE
SELECT (alias)
IF m.saverec <> 0
	GO m.saverec
ENDIF	

PROCEDURE prtdata
*****************
WAIT WINDOW 'Feature not available' NOWAIT
RETURN
SELECT (alias)
STORE RECNO() TO m.saverec
SELECT 0
USE prttab  
= checkzapfile("prttab")
SET ORDER TO serial
SELECT (alias)
GO TOP
DO WHILE !EOF()	
	SELECT prttab
	APPEND BLANK
	REPLACE serial WITH arrhea.serial
	REPLACE name   WITH suptab.name
**	REPLACE number WITH ALLTRIM(DTOC(arrhea.date))
	SELECT arrhea
	SKIP
ENDDO	
IF m.saverec <> 0
	GO m.saverec
ENDIF	
SELECT prttab
DO printall.spr

PROCEDURE fillrow
*****************
PARAMETERS tgetdata

=checkzapfile("sertmp1")
=checkzapfile("arrtmp")
IF m.ordhea =0
   RETURN
ENDIF   
SELECT arrhea
tarrrecno = RECNO()
tarrorder = ORDER()
SET ORDER TO ordhea
SELECT arrrow
SEEK m.serial
DO WHILE arrhea = arrhea.serial AND !EOF() AND tgetdata
   SCATTER MEMVAR
   SELECT arrtmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE arrhea   WITH m.serial
   REPLACE name     WITH protab.name
   REPLACE des      WITH protab.number
   REPLACE arrtype  WITH 'U'
   REPLACE arrrecno WITH RECNO('arrrow')
   SELECT sertab
   SEEK RECNO('arrrow')
   DO WHILE serial = RECNO('arrrow') AND !EOF()
      SELECT sertmp1
      APPEND BLANK
      REPLACE serial WITH RECNO('arrtmp')
      REPLACE ser_nbr WITH sertab.ser_nbr
      SELECT sertab
      SKIP
   ENDDO
   SELECT arrrow
   SKIP
ENDDO
SELECT ordhea
SEEK m.ordhea
SELECT ordrow
SEEK m.ordhea
DO WHILE ordhea = m.ordhea AND !EOF()
   SELECT arrtmp
   LOCATE FOR protab = ordrow.protab
   IF !FOUND()
      SELECT protab
      SEEK ordrow.protab
      SELECT arrtmp
      APPEND BLANK
      REPLACE arrhea WITH m.serial
      REPLACE protab WITH ordrow.protab
      REPLACE des    WITH protab.number
      REPLACE name   WITH protab.name
      REPLACE arrtype WITH 'A'
      REPLACE disc WITH (100*ordrow.disc+100*ordhea.discount-ordrow.disc*ordhea.discount)/100
   ENDIF
   REPLACE uprice WITH ordrow.uprice
   REPLACE allqty WITH allqty+ordrow.qty
   SELECT ordrow
   SKIP
ENDDO
*** GET QTY TAKEN *******
SELECT arrhea
SEEK m.ordhea
DO WHILE ordhea = m.ordhea AND !EOF()
   IF serial <> m.serial
      SELECT arrrow
      SEEK arrhea.serial
      DO WHILE arrhea = arrhea.serial
         SELECT arrtmp
         LOCATE FOR protab = arrrow.protab
         REPLACE qtytaken WITH qtytaken+arrrow.qty
         SELECT arrrow
         SKIP
      ENDDO
   ENDIF   
   SELECT arrhea
   SKIP
ENDDO
SELECT arrtmp
REPLACE ALL qtyleft WITH IIF(allqty-qtytaken<0,0,allqty-qtytaken)
IF !tgetdata
   REPLACE ALL qty    WITH qtyleft
   REPLACE ALL price WITH (qty*uprice*(100-disc))/100
ENDIF
GO TOP
*************************
SELECT arrhea
SET ORDER TO (tarrorder)
IF tarrrecno>0 AND tarrrecno<=RECCOUNT()
   GO RECORD tarrrecno
ENDIF
