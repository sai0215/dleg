CLOS DATA
report1 ='supredt'
DO supsedt.spr
CLOS DATA

PROCEDURE vblank
****************
m.curdes = ''
m.citdes = ''
m.cntdes = ''
SELECT (alias)
SCATTER MEMVAR BLANK
SHOW GETS

PROCEDURE getdata
*****************
m.curdes = curtab.number
m.citdes = cittab.name
m.cntdes = cnttab.name
SELECT (alias)
SCATTER MEMVAR
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

PROCEDURE prtdata
*****************
SELECT (alias)
STORE RECNO() TO m.saverec
SELECT prttab  
SET ORDER TO serial
= checkzapfile('prttab')
APPEND FROM (alias)
IF m.saverec <> 0
	GO m.saverec
ENDIF	
DO printall.spr

PROCEDURE prtrec
***************
SELECT prttab
DELETE FOR EMPTY(bidon)
SELECT (alias)
SET FILTER TO FOUND("prttab")
GO TOP
IF !EOF()
	DO newprint WITH report1,'FOUND("prttab")',IIF(m.print=1,'PREVIEW','TO PRINT')
ELSE
    WAIT WINDOW "EMPTY FILE" NOWAIT        
ENDIF    
