CLOS DATA
report1 ='fam1redt'
DO fam1sedt.spr
CLOS DATA

PROCEDURE vblank
****************
SELECT (alias)
SCATTER MEMVAR BLANK
SHOW GETS

PROCEDURE getdata
*****************
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

PROCEDURE prtdata
*****************
SELECT (alias)
STORE RECNO() TO m.saverec
= checkzapfile("prttab")
APPEND FROM (alias)
IF m.saverec <> 0
	GO m.saverec
ENDIF	
DO printall.spr
SELECT (alias)