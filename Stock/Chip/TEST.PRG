SELECT clitab
GO TOP
DO WHILE !EOF()
	m.serial = serial
	m.recno  = RECNO()
	IF SEEK(m.serial) AND m.recno<>RECNO()
		WAIT WIND STR(serial)
	ENDIF
	GO RECORD m.recno
	SKIP
ENDDO
RETURN

CLOS DATA
m.qtyexit=0
m.out=1
*DO printrec.spr
SELECT 0
USE mourow
SELECT 0
USE arrrow
SELECT 0
USE protab
SET FILTER TO 
SELECT 0
USE famtab1
SET ORDER TO serial
SELECT 0
USE fl2tab
SET ORDER TO serial
SELECT 0
USE fl1tab
SET ORDER TO serial
SELECT 0
USE itmsold
ZAP
SELECT protab
SET RELATION TO Protab.famtab1 INTO Famtab1 ADDITIVE
SET RELATION TO Protab.fl1tab INTO Fl1tab ADDITIVE
SET RELATION TO Protab.fl2tab INTO Fl2tab ADDITIVE
SELECT protab
SET ORDER TO proray
GO TOP
DO WHILE !EOF()
   qtyini=tiniqty()
   qtyexit=texitqty()
   SELECT itmsold
   APPEND BLANK
   REPLACE item    WITH protab.number
   REPLACE des     WITH protab.name
   REPLACE family  WITH famtab1.name
   REPLACE itmsold WITH qtyexit
   REPLACE itminit WITH qtyini
   SELECT protab
   SKIP
ENDDO
CLOS DATA
RETURN

PROCEDURE tiniqty
*****************
m.qtyini = 0
m.protab = protab.serial
SELECT arrrow
SET ORDER TO protab
SEEK m.protab
DO WHILE !EOF() AND arrrow.protab = m.protab
	m.qtyini = m.qtyini + arrrow.qty
	SELECT arrrow
	SKIP
ENDDO
SELECT protab
RETURN m.qtyini

PROCEDURE texitqty
******************
m.qtyexit = 0
m.protab = protab.serial
SELECT mourow
SET ORDER TO protab
SEEK m.protab
DO WHILE !EOF() AND mourow.protab = m.protab
	IF EMPTY(mourow.proret)
		m.qtyexit = m.qtyexit + mourow.qty
	ELSE
		m.qtyexit = m.qtyexit - mourow.qty
	ENDIF			
	SELECT mourow
	SKIP
ENDDO
SELECT protab
RETURN m.qtyexit