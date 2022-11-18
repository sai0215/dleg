CLOS DATA
report1  ='dclredt'
m.nbpro  = 0
m.totqty = 0
DO dclsedt.spr
CLOS DATA

PROCEDURE vblank
****************
SELECT (alias)
SCATTER MEMVAR BLANK
m.nbpro  = 0
m.totqty = 0
IF isadding
   trecno = RECNO()
   GO BOTTOM
   m.serial = serial+1
   IF trecno>0 AND trecno<=RECCOUNT()
      GO RECORD trecno
   ENDIF
ENDIF
IF m.isediting
	m.date = DATE()
ENDIF
= checkzapfile('dcltmp')
APPEND BLANK
REPLACE dclhea  WITH m.serial
REPLACE dcltype WITH 'N'
SHOW GETS
SELECT (alias)

PROCEDURE getdata
*****************
SELECT (alias)
SCATTER MEMVAR
= checkzapfile("dcltmp")
SELECT dclrow
SET ORDER TO dclhea
SEEK dclhea.serial
DO WHILE dclhea=dclhea.serial AND !EOF()
   SCATTER MEMVAR
   SELECT dcltmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE prodes   WITH protab.name
   REPLACE dcltype  WITH 'U'
   REPLACE dclrecno WITH RECNO('dclrow')
   m.inqty  = getinqty(m.protab)
   m.ouqty  = getouqty(m.protab)
   REPLACE curqty   WITH m.inqty-m.ouqty
   REPLACE adjqty   WITH protab.adjqty
   SELECT dclrow
   SKIP
ENDDO
SELECT dcltmp
COUNT TO m.nbpro
SUM qty TO m.totqty
APPEND BLANK
REPLACE dclhea  WITH m.serial
REPLACE dcltype WITH 'N'
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
tserial = m.serial
SELECT dcltmp
GO TOP
DO WHILE !EOF()
   DO CASE
      CASE dcltype='A'
           SCATTER MEMVAR
           SELECT dclrow
           APPEND BLANK
           GATHER MEMVAR
           REPLACE dclhea WITH tserial
      CASE dcltype='U'
           IF dclrecno<>0
              SCATTER MEMVAR
              SELECT dclrow
              GO RECORD m.dclrecno
              GATHER MEMVAR
           ENDIF
      CASE dcltype='D'
           IF dclrecno<>0
              SCATTER MEMVAR
              SELECT dclrow
              GO RECORD m.dclrecno
              DELETE
           ENDIF
   ENDCASE
   SELECT dcltmp
   REPLACE dcltype WITH IIF(dcltype='A','U',dcltype)
   SKIP
ENDDO
SET DELETED ON
SELECT (alias)

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
SELECT dclrow
DELETE FOR dclhea=m.serial
SELECT (alias)
RETURN .T.

PROCEDURE prtdata
*****************
m.qtyini  = 0
m.qtyexit = 0
m.tot1    = 0
SELECT dcltmp
GO TOP
DO WHILE !EOF()
	REPLACE selling WITH protab.selling
	REPLACE cost    WITH calcostpr(protab)
	REPLACE famtab1 WITH protab.famtab1
	SELECT dcltmp
	SKIP
ENDDO
m.out   = 1
m.okbtn = 2
DO printrec.spr
IF m.okbtn=2
	RETURN
ENDIF
SELECT dcltmp
SET ORDER TO famtab1
DELETE FOR dcltype='N'
IF m.out=1
	REPORT FORM dclredt PREVIEW
ELSE
	REPORT FORM dclredt TO PRINT OFF
ENDIF
SELECT dcltmp
SET ORDER TO
RECALL FOR dcltype='N'
SELECT (alias)

PROCEDURE search
*****************
SELECT (alias)
m.sfound = .F.
m.srecno = 0
IF issearch
   m.srecno = RECNO()
ENDIF
filesearch = "dclhea"
DO mousearc.spr
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

PROCEDURE import
****************
dfile = SYS(2000,'c:\stock\unitech\*.txt')
IF EMPTY(dfile)
	WAIT WINDOW 'NO INVENTORY FILES'
	RETURN
ENDIF
SELECT dcltmp
SET RELATION OFF INTO Protab
SELECT protab
SET ORDER TO barcode
SELECT reader
SET RELATION TO barcode INTO Protab ADDITIVE

SELECT dcltmp
SET ORDER TO des
= checkzapfile("reader")

tfile = 'c:\stock\unitech\'+dfile
FOR i=1 TO 20
	IF FILE(tfile)
		APPEND FROM &tfile TYPE DELIMITED WITH TAB FIELDS Reader.code
	ENDIF	
	tfile = 'c:\stock\unitech\'+SYS(2000,'c:\stock\unitech\*.txt',1)	
ENDFOR

SELECT reader
DELETE FOR EMPTY(code)
REPLACE ALL barcode WITH SUBSTR(code,1,13)
**REPLACE ALL barcode WITH IIF(LEN(ALLTRIM(code))<=13,PADL(ALLTRIM(code),13,'0'),RIGHT(ALLTRIM(code),13))
REPLACE ALL qty     WITH SUBSTR(code,15,3)
GO TOP
DO WHILE !EOF()
	m.barcode = ALLTRIM(barcode)
	m.qty     = VAL(qty)
	SELECT dcltmp
	IF !SEEK(m.barcode)
		APPEND BLANK
		REPLACE dclhea   WITH m.serial
		REPLACE des      WITH m.barcode
		REPLACE protab   WITH protab.serial
		REPLACE prodes   WITH protab.name
		REPLACE field2   WITH protab.field2
		REPLACE size     WITH fl4tab.name
		REPLACE qty      WITH m.qty
		REPLACE dcltype  WITH 'A'
		REPLACE dclrecno WITH m.serial
		m.inqty  = getinqty(protab.serial)
		m.ouqty  = getouqty(protab.serial)
		REPLACE curqty   WITH m.inqty-m.ouqty
		REPLACE adjqty   WITH protab.adjqty
	ELSE
		REPLACE qty      WITH qty+m.qty
	ENDIF
	SELECT reader
	SKIP
ENDDO
SHOW GET m.itmpop

tfile = 'c:\stock\unitech\'+SYS(2000,'c:\stock\unitech\*.txt')
FOR i=1 TO 20
	IF FILE(tfile)
		DELETE FILE &tfile
	ENDIF	
	tfile = 'c:\stock\unitech\'+SYS(2000,'c:\stock\unitech\*.txt',1)	
ENDFOR

SELECT reader
SET RELATION OFF INTO Protab
SELECT protab
SET ORDER TO serial
SELECT dcltmp
SET RELATION TO protab INTO Protab ADDITIVE
SELECT dclhea

PROCEDURE update
****************
IF msg_box('This will modify the current quantities in your stock. Proceed?',1,'Stock update')=2
	RETURN
ENDIF
SELECT protab
REPLACE ALL bidon WITH ''
SELECT dcltmp
GO TOP
DO WHILE !EOF()
	m.protab = protab
	m.qty    = qty
	m.curqty = curqty
	SELECT protab
	IF SEEK(m.protab) AND serial<>0
		REPLACE adjqty        WITH m.qty-m.curqty
		REPLACE dcltmp.adjqty WITH m.qty-m.curqty
		REPLACE dcltmp.curqty WITH m.curqty
		REPLACE bidon WITH 'X'
	ENDIF
	SELECT dcltmp
	SKIP
ENDDO
GO TOP
SELECT protab
SET FILTER TO EMPTY(bidon)
GO TOP
DO WHILE !EOF()
	m.inqty  = getinqty(protab.serial)
	m.ouqty  = getouqty(protab.serial)
	m.adjqty = protab.adjqty
	m.curqty   = m.inqty-m.ouqty+m.adjqty
	IF m.curqty<>0
		SELECT protab
		REPLACE adjqty WITH -(m.inqty-m.ouqty)
	ENDIF
	SELECT protab
	SKIP
ENDDO
SELECT dclhea
WAIT WINDOW 'Finished updating' NOWAIT

PROCEDURE getinqty
******************
PARAMETERS pserial

m.qtyini = 0
SELECT arrrow
SUM qty TO m.qtyini FOR protab=pserial
SELECT dcltmp
RETURN m.qtyini

PROCEDURE getouqty
******************
PARAMETERS pserial

**m.qtyexit = protab.adjqty
m.qtyexit = 0
m.protab  = protab.serial
SELECT mourow
SET ORDER TO protab
SEEK m.protab
DO WHILE !EOF() AND mourow.protab=m.protab
	IF EMPTY(mourow.proret) AND !mouhea.return
		m.qtyexit = m.qtyexit + (mourow.qty+mourow.free)
	ELSE
		m.qtyexit = m.qtyexit - mourow.qty
	ENDIF
	SELECT mourow
	SKIP
ENDDO
SELECT dcltmp
RETURN m.qtyexit

PROCEDURE calcostpr
*******************
PARAMETERS tserial

tselect = SELECT()
SELECT arrrow
SUM cal_cur(price,arrhea.date,arrhea.curtab,3) to tprice FOR protab=tserial
SUM qty to tqty FOR protab=tserial
tcost = IIF(tqty<>0,(tprice/tqty),tprice)
SELECT (tselect)
RETURN tcost