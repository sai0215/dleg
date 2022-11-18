CLOS DATA
SELECT 0
USE arrhea
SET ORDER TO serial
SELECT 0
USE arrrow
SET RELATION TO Arrrow.arrhea INTO Arrhea ADDITIVE
SELECT 0
USE protab
SELECT 0
USE rathea
SELECT 0
USE ratrow
m.date = DATE()
m.ok   = 1
DO arrredt.spr
IF m.ok=2
	CLOS DATA
	RETURN
ENDIF
m.out   = 1
m.okbtn = 1
DO printrec.spr
IF m.okbtn=2
	CLOS DATA
	RETURN
ENDIF
SELECT arrrow
SET FILTER TO arrhea.date>=m.date
IF m.out=1
	REPORT FORM arredt1.frx PREVIEW
ELSE
	REPORT FORM arredt1.frx TO PRINT OFF
ENDIF	
CLOS DATA

PROCEDURE getname
*****************
PARAMETERS tpro

tselect = SELECT()
SELECT protab
SET ORDER TO serial
IF SEEK(tpro)
	tname = name
ELSE
	tname = ''
ENDIF
SELECT (tselect)
RETURN tname
