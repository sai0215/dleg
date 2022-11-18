SELECT arrrow
GO TOP
DO WHILE !EOF()
	m.qty = qty
	SELECT protab
	REPLACE protab.qte2 WITH m.qty
	SELECT arrrow
	SKIP
ENDDO
wait wind 'Finish'