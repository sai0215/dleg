CLEAR
CLEAR ALL
CLOSE DATA
CLEAR MACROS
SET PATH TO c:\stock\prg;c:\stock\bmp;c:\stock\scr;c:\stock\dbf;c:\stock\cdx
SET TALK    OFF
SET SAFETY  OFF
SET ESCAPE  OFF
SET CENTURY OFF
SET HELP    OFF
SET DELETED ON
SET CENTURY ON
SET DATE FRENCH
SET HOURS TO 24
SET STATUS BAR ON
SET CLOCK STATUS
SET PROC TO proc
ON ERROR DO err_prg WITH ERROR(),MESSAGE(),MESSAGE(1),PROGRAM(),LINENO()
PUBLIC dcnt
dcnt = 115
DO fixdata.spr