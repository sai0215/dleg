CLOS DATA
IF !USED('pmotab')
	SELECT 0
	USE c:\stock\dbf1\pmotab
ENDIF
ZAP
APPEND FROM c:\stock\dbf\pmotab.dbf