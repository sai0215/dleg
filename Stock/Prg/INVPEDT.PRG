CLOS DATA
report ='invredt'
report1 ='invredt1'
DO invsedt.spr
CLOS DATA

PROCEDURE vblank
****************
=checkzapfile("invtmp")
SELECT (alias)
m.payprice  = 0
m.restprice = 0
m.totprice  = 0
m.isquohea  = .F.
m.quodes = ''
m.moudes = ''
m.clides = ''
m.quodes = ''
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
SELECT invtmp
APPEND BLANK
REPLACE invhea   WITH m.serial
REPLACE invtype  WITH 'N'
SHOW GETS
SELECT (alias)

PROCEDURE getdata
*****************
m.quodes = IIF(invhea.quomou=1,quohea.ref_no,'')
m.moudes = IIF(invhea.quomou=2,mouhea.ref_no,'')
m.curdes = curtab.number
m.clides = clitab.name
SELECT (alias)
SCATTER MEMVAR
=checkzapfile("invtmp")
SELECT invrow
SEEK invhea.serial
DO WHILE invhea = invhea.serial AND !EOF()
   SCATTER MEMVAR
   SELECT invtmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE invhea   WITH m.serial
   REPLACE name     WITH protab.name
   REPLACE des      WITH protab.number
   REPLACE invtype  WITH 'U'
   REPLACE invrecno WITH RECNO('invrow')
   SELECT invrow
   SKIP
ENDDO
m.isquohea = (m.quohea<>0)
SELECT paytab
SUM amount  TO m.payprice FOR invhea = m.serial
SELECT invtmp
APPEND BLANK
REPLACE invhea  WITH m.serial
REPLACE invtype WITH 'N'
DO fillrow WITH .T.
DO caldisc
SHOW GETS
SELECT (alias)

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
SELECT invtmp
GO TOP
DO WHILE !EOF()
   DO CASE
      CASE invtype = 'A' AND qty > 0
           SCATTER MEMVAR
           SELECT invrow
           APPEND BLANK
           GATHER MEMVAR
           REPLACE invhea WITH m.serial
      CASE invtype = 'U'
           IF invrecno <> 0
              IF qty = 0
                 SCATTER MEMVAR
                 SELECT invrow
                 GO RECORD m.invrecno
                 DELETE
                 SELECT invtmp
              ELSE
                 SCATTER MEMVAR
                 SELECT invrow
                 GO RECORD m.invrecno
                 GATHER MEMVAR
                 REPLACE invhea WITH m.serial
              ENDIF
           ENDIF   
      CASE invtype = 'D'
           IF invrecno <> 0
              SCATTER MEMVAR
              SELECT invrow
              GO RECORD m.invrecno
              DELETE
           ENDIF
   ENDCASE
   SELECT invtmp
   REPLACE invtype WITH IIF(invtype = 'A' , 'U', invtype )
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
DO trasearch.spr
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
SELECT invrow
DELETE FOR invhea = m.serial
SELECT paytab
DELETE FOR invhea = m.serial
SELECT invhea
RETURN .T.


PROCEDURE caldisc
*****************
SELECT invtmp
m.ttinvrec = RECNO()
GO TOP
SUM fprice  TO m.totprice
*m.totprice = (m.totprice*(100-m.discount))/100
SELECT invtmp
IF m.ttinvrec > 0 AND m.ttinvrec <=RECCOUNT()
   GO RECORD m.ttinvrec
ENDIF   
m.restprice = m.totprice - m.payprice

PROCEDURE calpay
*****************
SELECT paytmp
m.ttpayrec = RECNO()
GO TOP
SUM amount  TO m.payprice
SELECT paytmp
IF m.ttpayrec > 0 AND m.ttpayrec <=RECCOUNT()
   GO RECORD m.ttpayrec
ENDIF   
m.restprice = m.totprice - m.payprice

PROCEDURE prtrec
***************
SELECT prttab
DELETE FOR EMPTY(bidon)
SELECT (alias)
SET RELATION TO serial INTO Prttab 
DO print WITH IIF(m.criteria = 1,report,report1),IIF(m.criteria=1,'RECNO()=m.saverec',IIF(m.criteria=2,'FOUND("prttab")','.T.')),IIF(m.print=1,'PREVIEW','TO PRINT')
SELECT prttab
USE
SELECT (alias)
IF m.saverec <> 0
	GO m.saverec
ENDIF	

PROCEDURE prtdata
*****************
m.out=1
SELECT (alias)
wrec = RECNO()
DO printrec.spr
IF m.out=1
	REPORT FORM invredt FOR RECNO() = wrec PREVIEW
ELSE
	REPORT FORM invredt FOR RECNO() = wrec TO PRINT
ENDIF	
return
SELECT (alias)
STORE RECNO() TO m.saverec
SELECT 0
USE prttab  
SET ORDER TO serial
=checkzapfile("prttab")
SELECT (alias)
GO TOP
DO WHILE !EOF()	
	SELECT prttab
	APPEND BLANK
	REPLACE serial WITH invhea.serial
	REPLACE name WITH clitab.name
	REPLACE number WITH ALLTRIM(DTOC(invhea.date))
	SELECT invhea
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
=checkzapfile("invtmp")
IF m.quomou =0
   RETURN
ENDIF   
SELECT invhea
tinvrecno = RECNO()
tinvorder = ORDER()
SET ORDER TO quohea
SELECT invrow
SEEK m.serial
DO WHILE invhea = invhea.serial AND !EOF() AND tgetdata
   SCATTER MEMVAR
   SELECT invtmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE invhea   WITH m.serial
   REPLACE name     WITH protab.name
   REPLACE des      WITH protab.number
   REPLACE invtype  WITH 'U'
   REPLACE invrecno WITH RECNO('invrow')
   SELECT invrow
   SKIP
ENDDO
IF m.quomou =1
   SELECT quohea
   SEEK m.quohea
   SELECT quorow
   SEEK m.quohea
   DO WHILE quohea = m.quohea AND !EOF()
      SELECT invtmp
      LOCATE FOR protab = quorow.protab
      IF !FOUND()
         SELECT protab
         SEEK quorow.protab
         SELECT invtmp
         APPEND BLANK
         REPLACE invhea WITH m.serial
         REPLACE protab WITH quorow.protab
         REPLACE des    WITH protab.number
         REPLACE name   WITH protab.name
         REPLACE invtype WITH 'A'
         REPLACE price WITH quorow.price
         REPLACE disc WITH (100*quorow.disc+100*quohea.discount-quorow.disc*quohea.discount)/100
      ENDIF
      REPLACE allqty WITH allqty+quorow.qty
      SELECT quohea
      SKIP
   ENDDO
ENDIF
IF m.quomou =2
   SELECT mouhea
   SEEK m.quohea
   SELECT mourow
   SEEK m.quohea
   DO WHILE mouhea = m.quohea AND !EOF()
      SELECT invtmp
      LOCATE FOR protab = mourow.protab
      IF !FOUND()
         SELECT protab
         SEEK mourow.protab
         SELECT invtmp
         APPEND BLANK
         REPLACE invhea WITH m.serial
         REPLACE des    WITH protab.number
         REPLACE protab WITH mourow.protab
         REPLACE name   WITH protab.name
         REPLACE invtype WITH 'A'
         REPLACE price WITH mourow.uprice
         REPLACE disc WITH (100*mourow.disc+100*mouhea.discount-mourow.disc*mouhea.discount)/100
      ENDIF
      REPLACE allqty WITH allqty+mourow.qty
      SELECT mourow
      SKIP
   ENDDO
ENDIF
*** GET QTY TAKEN *******
SELECT invhea
SEEK m.quohea
DO WHILE quohea = m.quohea AND !EOF()
   IF serial <> m.serial AND quomou = m.quomou
      SELECT invrow
      SEEK invhea.serial
      DO WHILE invhea = invhea.serial
         SELECT invtmp
         LOCATE FOR protab = invrow.protab
         REPLACE qtytaken WITH qtytaken+invrow.qty
         SELECT invrow
         SKIP
      ENDDO
   ENDIF   
   SELECT invhea
   SKIP
ENDDO
*************************
SELECT invtmp
REPLACE ALL qtyleft WITH IIF(allqty-qtytaken<0,0,allqty-qtytaken)
IF !tgetdata
   REPLACE ALL qty    WITH qtyleft
   REPLACE ALL fprice WITH (qty*price*(100-disc))/100
ENDIF
GO TOP
SELECT invhea
SET ORDER TO (tinvorder)
IF tinvrecno>0 AND tinvrecno<=RECCOUNT()
   GO RECORD tinvrecno
ENDIF

