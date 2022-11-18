SELECT trans
GO TOP
DO WHILE !EOF()
	m.barcode = barcode
	m.qty     = VAL(qty)
	SELECT protab
	IF SEEK(m.barcode)
		REPLACE invqty WITH invqty+m.qty
	ELSE
		WAIT WINDOW m.barcode	
	ENDIF
	SELECT trans
	SKIP
ENDDO
return

SELECT protab
GO TOP
DO WHILE !EOF()
	m.qty = tiniqty1() - texitqty()
	SELECT protab
	REPLACE qty WITH m.qty
	SKIP
ENDDO
RETURN

SELECT protab
GO TOP
DO WHILE !EOF()
	m.barcode = ALLTRIM(barcode)
	SKIP
	IF ALLTRIM(barcode)=m.barcode
		BROW
	ENDIF
ENDDO
RETURN