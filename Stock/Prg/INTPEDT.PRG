CLOS DATA
report ='intredt'
report1='intredt1'
DO intsedt.spr
CLOS DATA

PROCEDURE vblank
****************
=checkzapfile("sertmp1")
=checkzapfile("inttmp")
SELECT (alias)
m.wardes1 = ''
m.wardes2 = ''
m.reades  = ''
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
SELECT inttmp
APPEND BLANK
REPLACE inthea   WITH m.serial
REPLACE inttype  WITH 'N'
SHOW GETS
SELECT (alias)

PROCEDURE getdata
*****************
SELECT warhea
LOCATE FOR serial = inthea.warhea1
m.wardes1 = name
LOCATE FOR serial = inthea.warhea2
m.wardes2 = name
SELECT reatab
m.reades = name
SELECT (alias)
SCATTER MEMVAR
=checkzapfile("sertmp1")
=checkzapfile("inttmp")
SELECT introw
SEEK inthea.serial
DO WHILE inthea = inthea.serial AND !EOF()
   SCATTER MEMVAR
   SELECT inttmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE inthea   WITH m.serial
   REPLACE name     WITH protab.name
   REPLACE inttype  WITH 'U'
   REPLACE intrecno WITH RECNO('introw')
   SELECT sritab
   SEEK RECNO('introw')
   DO WHILE serial = RECNO('introw') AND !EOF()
      SELECT sertmp1
      APPEND BLANK
      REPLACE serial WITH RECNO('inttmp')
      REPLACE ser_nbr WITH sritab.ser_nbr
      SELECT sritab
      SKIP
   ENDDO
   SELECT introw
   SKIP
ENDDO
SELECT inttmp
APPEND BLANK
REPLACE inthea  WITH m.serial
REPLACE inttype WITH 'N'
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
SELECT introw
SEEK inthea.serial
DO WHILE inthea = inthea.serial AND !EOF()
   SELECT sritab
   DELETE FOR serial = RECNO('introw')
   SELECT introw
   SKIP
ENDDO

SET DELETED OFF
SELECT inttmp
GO TOP
DO WHILE !EOF()
   DO CASE
      CASE inttype = 'A'
           SCATTER MEMVAR
           SELECT introw
           APPEND BLANK
           GATHER MEMVAR
           REPLACE inthea WITH m.serial
           SELECT sertmp1
           SEEK RECNO('inttmp')
           DO WHILE serial = RECNO('inttmp') AND !EOF()
              SELECT sritab
              APPEND BLANK
              REPLACE serial  WITH RECNO('introw')
              REPLACE ser_nbr WITH sertmp1.ser_nbr
              SELECT sertmp1
              SKIP
           ENDDO
           SELECT inttmp
      CASE inttype = 'U'
           IF intrecno <> 0
              SCATTER MEMVAR
              SELECT introw
              GO RECORD m.intrecno
              GATHER MEMVAR
              REPLACE inthea WITH m.serial
              SELECT sertmp1
              SEEK RECNO('inttmp')
              DO WHILE serial = RECNO('inttmp') AND !EOF()
                 SELECT sritab
                 APPEND BLANK
                 REPLACE serial  WITH RECNO('introw')
                 REPLACE ser_nbr WITH sertmp1.ser_nbr
                 SELECT sertmp1
                 SKIP
              ENDDO
              SELECT inttmp
           ENDIF
      CASE inttype = 'D'
           IF intrecno <> 0
              SCATTER MEMVAR
              SELECT introw
              GO RECORD m.intrecno
              DELETE
           ENDIF
   ENDCASE
   SELECT inttmp
   REPLACE inttype WITH IIF(inttype = 'A' , 'U', inttype )
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
DO intsearch.spr
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
SELECT introw
DELETE FOR inthea = m.serial
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
	REPLACE serial WITH inthea.serial
	REPLACE name WITH clitab.name
	REPLACE number WITH ALLTRIM(DTOC(inthea.date))
	SELECT inthea
	SKIP
ENDDO	
IF m.saverec <> 0
	GO m.saverec
ENDIF	
SELECT prttab
DO printall.spr
