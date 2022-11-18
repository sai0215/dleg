CLOS DATA
report ='traredt'
report1='traredt1'
DO trasedt.spr
CLOS DATA

PROCEDURE vblank
****************
SELECT sertmp1
ZAP
SELECT tratmp
ZAP
SELECT (alias)
m.isquohea = .F.
m.wardes = ''
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
   m.inout = 1
   IF isadding
      m.date = DATE()
   ENDIF   
ENDIF
SELECT tratmp
APPEND BLANK
REPLACE trahea   WITH m.serial
REPLACE tratype  WITH 'N'
SHOW GETS
SELECT (alias)

PROCEDURE getdata
*****************
SELECT quohea
LOCATE FOR serial = trahea.quohea
m.quodes = ref_no
SELECT warhea
m.wardes = name
SELECT clitab
m.clides = name
SELECT (alias)
SCATTER MEMVAR
SELECT sertmp1
ZAP
SELECT tratmp
ZAP
SELECT trarow
SEEK trahea.serial
DO WHILE trahea = trahea.serial AND !EOF()
   SCATTER MEMVAR
   SELECT tratmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE trahea   WITH m.serial
   REPLACE name     WITH protab.name
   REPLACE tratype  WITH 'U'
   REPLACE trarecno WITH RECNO('trarow')
   SELECT srntab
   SEEK RECNO('trarow')
   DO WHILE serial = RECNO('trarow') AND !EOF()
      SELECT sertmp1
      APPEND BLANK
      REPLACE serial WITH RECNO('tratmp')
      REPLACE ser_nbr WITH srntab.ser_nbr
      SELECT srntab
      SKIP
   ENDDO
   SELECT trarow
   SKIP
ENDDO
m.isquohea = (m.quohea<>0)
SELECT tratmp
APPEND BLANK
REPLACE trahea  WITH m.serial
REPLACE tratype WITH 'N'
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
SELECT (alias)
GATHER MEMVAR
SELECT trarow
SEEK trahea.serial
DO WHILE trahea = trahea.serial AND !EOF()
   SELECT srntab
   DELETE FOR serial = RECNO('trarow')
   SELECT trarow
   SKIP
ENDDO

SET DELETED OFF
SELECT tratmp
GO TOP
DO WHILE !EOF()
   DO CASE
      CASE tratype = 'A'
           SCATTER MEMVAR
           SELECT trarow
           APPEND BLANK
           GATHER MEMVAR
           REPLACE trahea WITH m.serial
           SELECT sertmp1
           SEEK RECNO('tratmp')
           DO WHILE serial = RECNO('tratmp') AND !EOF()
              SELECT srntab
              APPEND BLANK
              REPLACE serial  WITH RECNO('trarow')
              REPLACE ser_nbr WITH sertmp1.ser_nbr
              SELECT sertmp1
              SKIP
           ENDDO
           SELECT tratmp
      CASE tratype = 'U'
           IF trarecno <> 0
              SCATTER MEMVAR
              SELECT trarow
              GO RECORD m.trarecno
              GATHER MEMVAR
              REPLACE trahea WITH m.serial
              SELECT sertmp1
              SEEK RECNO('tratmp')
              DO WHILE serial = RECNO('tratmp') AND !EOF()
                 SELECT srntab
                 APPEND BLANK
                 REPLACE serial  WITH RECNO('trarow')
                 REPLACE ser_nbr WITH sertmp1.ser_nbr
                 SELECT sertmp1
                 SKIP
              ENDDO
              SELECT tratmp
           ENDIF
      CASE tratype = 'D'
           IF trarecno <> 0
              SCATTER MEMVAR
              SELECT trarow
              GO RECORD m.trarecno
              DELETE
           ENDIF
   ENDCASE
   SELECT tratmp
   REPLACE tratype WITH IIF(tratype = 'A' , 'U', tratype )
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
SELECT trarow
DELETE FOR trahea = m.serial
SELECT (alias)
RETURN .T.

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
	REPORT FORM traredt FOR RECNO() = wrec PREVIEW
ELSE
	REPORT FORM traredt FOR RECNO() = wrec TO PRINT
ENDIF	
return
SELECT (alias)
STORE RECNO() TO m.saverec
SELECT 0
USE prttab  
SET ORDER TO serial
ZAP
SELECT (alias)
GO TOP
DO WHILE !EOF()	
	SELECT prttab
	APPEND BLANK
	REPLACE serial WITH trahea.serial
	REPLACE name WITH clitab.name
	REPLACE number WITH ALLTRIM(DTOC(trahea.date))
	SELECT trahea
	SKIP
ENDDO	
IF m.saverec <> 0
	GO m.saverec
ENDIF	
SELECT prttab
DO printall.spr
