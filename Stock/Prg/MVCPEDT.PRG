CLOS DATA
DO mvcsedt.spr
CLOS DATA

PROCEDURE getcred
*****************
m.tpbalance=0
m.tptotal=0
SELECT clitab
GO TOP
DO WHILE !EOF()
	DO gettotcli 
	SELECT clitab
	REPLACE balance WITH m.tpbalance
	REPLACE tot_pur WITH m.tptotal
	SKIP
ENDDO	
