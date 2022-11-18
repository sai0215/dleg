SELECT 0
USE dbf
GO TOP
DO WHILE !EOF()
	tname = name
	SELECT 0
	USE &tname
	IF ALIAS()=ALLTRIM(tname)
		REINDEX
	ENDIF	
	SELECT dbf
	SKIP
ENDDO