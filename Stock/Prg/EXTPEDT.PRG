CLOS DATA
report ='extredt'
DO extsedt.spr
CLOS DATA

PROCEDURE vblank
****************
SELECT (alias)
SCATTER MEMVAR BLANK
m.curdes = ''
SHOW GETS

PROCEDURE getdata
*****************
SELECT (alias)
SCATTER MEMVAR
m.curdes = curtab.number
DO CASE
CASE m.dest='A'
	m.dest = 'Arrival'
CASE m.dest='O'
	m.dest = 'Order'
CASE m.dest='B'
	m.dest = 'Both'	
ENDCASE		
SHOW GETS

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
DO search.spr
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
RETURN .T.

PROCEDURE prtrec
***************
SELECT prttab
DELETE FOR EMPTY(bidon)
SELECT (alias)
SET RELATION TO serial INTO Prttab 
DO print WITH report,IIF(m.criteria=1,'RECNO()=m.saverec',IIF(m.criteria=2,'FOUND("prttab")','.T.')),IIF(m.print=1,'PREVIEW','TO PRINT')
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
=checkzapfile("prttab")
SET ORDER TO serial
APPEND FROM (alias)
IF m.saverec <> 0
	GO m.saverec
ENDIF	
DO printall.spr
