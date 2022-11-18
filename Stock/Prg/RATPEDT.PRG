CLOS DATA
report1 ='ratredt'
DO ratsedt.spr
CLOS DATA

PROCEDURE vblank
****************
SELECT rattmp
ZAP
SELECT curtab
GO TOP
DO WHILE !EOF()
   SELECT rattmp
   APPEND BLANK
   REPLACE name WITH curtab.name
   REPLACE curtab WITH curtab.serial
   SELECT curtab
   SKIP
ENDDO   
SELECT (alias)
m.supdes   = ''
m.saldes   = ''
m.bandes   = ''
m.curdes   = ''
m.bandes   = ''
m.lcurdes  = ''
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
SELECT rattmp
SHOW GETS
SELECT (alias)

PROCEDURE getdata
*****************
SELECT (alias)
SCATTER MEMVAR
SELECT ratrow
SEEK rathea.serial
DO WHILE rathea = rathea.serial AND !EOF()
   SCATTER MEMVAR
   SELECT rattmp
   LOCATE FOR curtab = ratrow.curtab
   IF FOUND()
      GATHER MEMVAR
      REPLACE rathea   WITH m.serial
      REPLACE name     WITH curtab.name
      REPLACE rattype  WITH 'U'
      REPLACE ratrecno WITH RECNO('ratrow')
   ENDIF
   SELECT ratrow
   SKIP
ENDDO
SELECT rattmp
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
SELECT rattmp
GO TOP
DO WHILE !EOF()
   SCATTER MEMVAR
   SELECT ratrow
   LOCATE FOR curtab = rattmp.curtab AND rathea = m.serial
   IF !FOUND()
      APPEND BLANK
   ENDIF   
   GATHER MEMVAR
   REPLACE rathea WITH m.serial
   SELECT rattmp
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
DO ratsearch.spr
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
SELECT ratrow
DELETE FOR rathea = m.serial
SELECT (alias)
RETURN .T.

PROCEDURE prtdata
*****************
SELECT (alias)
STORE RECNO() TO m.saverec
SELECT prttab  
ZAP
APPEND FROM (alias)
IF m.saverec <> 0 
	GO m.saverec
ENDIF	
DO printall.spr

PROCEDURE caldisc
*****************
RETURN

PROCEDURE prtrec
***************
SELECT prttab
DELETE FOR EMPTY(bidon)
SELECT (alias)
SET FILTER TO FOUND("prttab")
SET ORDER TO serial
SELECT ratrow
SET RELATION TO ratrow.rathea INTO rathea
SET RELATION TO ratrow.curtab INTO curtab ADDITIVE
GO TOP
IF !EOF()
	DO newprint WITH report1,'.T.',IIF(m.print=1,'PREVIEW','TO PRINT')
ELSE
    WAIT WINDOW "EMPTY FILE" NOWAIT        
ENDIF    
SELECT ratrow
SET RELATION OFF INTO rathea
SET RELATION OFF INTO curtab
SET RELATION TO ratrow.curtab INTO curtab
SELECT rathea
SET RELATION TO rathea.serial INTO prttab

