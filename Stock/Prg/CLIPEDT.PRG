CLOS DATA
report1 ='cliredt'
DO clisedt.spr
CLOS DATA

PROCEDURE vblank
****************
m.citdes = ''
m.cntdes = ''
SELECT (alias)
SCATTER MEMVAR BLANK
SHOW GETS
=checkzapfile("contmp")
APPEND BLANK
REPLACE clitab   WITH m.serial
REPLACE contype  WITH 'N'
SHOW GETS
SELECT (alias)

PROCEDURE getdata
*****************
m.citdes = cittab.name
m.cntdes = cnttab.name
=checkzapfile("contmp")
SELECT contab
SEEK clitab.serial
DO WHILE clitab = clitab.serial AND !EOF()
   SCATTER MEMVAR
   SELECT contmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE clitab   WITH m.serial
   REPLACE name     WITH contab.name
   REPLACE des      WITH contab.des
   REPLACE date     WITH contab.date
   REPLACE contype  WITH 'U'
   REPLACE conrecno WITH RECNO('contab')
   SELECT contab
   SKIP
ENDDO
SELECT contmp
APPEND BLANK
REPLACE clitab  WITH m.serial
REPLACE contype WITH 'N'
SELECT (alias)
SCATTER MEMVAR
DO getclient WITH m.serial,2
SELECT (alias)
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
SET DELETED OFF
SELECT contmp
GO TOP
DO WHILE !EOF()
   DO CASE
      CASE contype = 'A'
           SCATTER MEMVAR
           SELECT contab
           APPEND BLANK
           GATHER MEMVAR
           REPLACE clitab WITH m.serial
      CASE contype = 'U'
           IF conrecno <> 0
              SCATTER MEMVAR
              SELECT contab
              GO RECORD m.conrecno
              GATHER MEMVAR
              REPLACE clitab WITH m.serial
           ENDIF
      CASE contype = 'D'
           IF conrecno <> 0
              SCATTER MEMVAR
              SELECT contab
              GO RECORD m.conrecno
              DELETE
           ENDIF
   ENDCASE
   SELECT contmp
   REPLACE contype WITH IIF(contype = 'A' , 'U', contype )
   SKIP        
ENDDO
SET DELETED ON


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
GO TOP
DO WHILE !EOF()
	m.tpbalance=0
	m.tptotal  =0 
	DO gettotcli
	SELECT clitab
	REPLACE balance WITH m.tpbalance
	REPLACE tot_pur WITH m.tptotal 
	SKIP
ENDDO	
=checkzapfile("prttab")
APPEND FROM (alias)
IF m.saverec <> 0 AND m.saverec <= RECCOUNT()
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

PROCEDURE releve
****************
tselect = SELECT()
m.clitab = clitab.serial
SELECT mourow.* FROM mourow,mouhea WHERE mouhea.serial=mourow.mouhea AND mouhea.suptab=m.clitab INTO CURSOR releve
SELECT releve
SET ORDER TO TAG serial IN mouhea
SET RELATION TO releve.mouhea INTO mouhea ADDITIVE
SET ORDER TO TAG serial IN protab
SET RELATION TO releve.protab INTO protab ADDITIVE
m.out   = 1
m.okbtn = 1
DO printrec.spr
IF m.okbtn=2
	RETURN
ENDIF	
IF m.out=1
	REPORT FORM relredt.frx PREVIEW
ELSE
	REPORT FORM relredt.frx TO PRINTER OFF
ENDIF	
SELECT (tselect)
