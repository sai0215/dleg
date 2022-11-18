PROCEDURE browpro
*****************
IF !USED('protab')
	SELECT 0
	USE protab SHARED
ELSE
	SELECT protab
ENDIF
SET FILTER TO serial>=4962
GO TOP
BROWSE FIELDS name:R,number:R,selling,curtab

IF !USED('ordrow')
	SELECT 0
	USE ordrow
ENDIF
SELECT protab
GO TOP
DO WHILE !EOF()
	IF curtab<>1
		m.protab = serial
		m.price  = selling
		SELECT ordrow
		REPLACE uprice WITH m.price     FOR protab=m.protab
		REPLACE price  WITH m.price*qty FOR protab=m.protab
	ENDIF
	SELECT protab
	SKIP
ENDDO


PROCEDURE priceit
*****************
WAIT WINDOW 'Are you sure? (Y/N)'
IF LASTKEY()<>89 AND LASTKEY()<>121
	RETURN
ENDIF
IF !USED('protab')
	SELECT 0
	USE protab SHARED
ELSE
	SELECT protab
ENDIF
SET FILTER TO serial>=4962
REPLACE selling1 WITH selling FOR curtab<>1 AND !EMPTY(selling)
REPLACE selling  WITH ROUND(((selling*2.5)/5.5)*1550,-3) FOR curtab<>1 AND !EMPTY(selling)
REPLACE curtab   WITH 1 FOR curtab<>1 AND !EMPTY(selling)
WAIT WIND "Conversion finished" NOWAIT