IF !USED('c_s_lst')
	SELECT 0
	USE c_s_lst
ELSE
	SELECT c_s_lst
ENDIF

SET ORDER TO TAG name
GO TOP
SET FILTER TO categ = 'Customers'
REPORT FORM cusredt ENVIRONMENT NOCONSOLE TO PRINTER
SET FILTER TO 