CLOS DATA
report ='fldredt'
DO fldsedt.spr
CLOS DATA

PROCEDURE vblank
****************
SELECT (alias)
SCATTER MEMVAR BLANK

PROCEDURE getdata
*****************
SELECT (alias)
SCATTER MEMVAR

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

PROEDURE prtrec
***************
SELECT (alias)
DO print WITH report,IIF(m.criteria=1,'RECNO()=m.saverec',IIF(m.criteria=2,'NOT EMPTY(bidon)','.T.')),IIF(m.print=1,'PREVIEW','TO PRINT')
