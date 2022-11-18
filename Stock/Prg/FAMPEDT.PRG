PARAMETERS m.choice

IF m.choice <> 3
	DO famsedt1.spr
ELSE
	DO delfam	
ENDIF	
SHOW GET m.fampop

PROCEDURE newfam
****************
IF EMPTY(m.number)
	RETURN
ENDIF	
SELECT famtab
m.serial = getmax('famtab','serial') + 1
APPEND BLANK
GATHER MEMVAR

PROCEDURE editfam
*****************
SELECT famtab
IF EMPTY(m.number) OR EOF()
	RETURN
ENDIF	
GATHER MEMVAR

PROCEDURE delfam
****************
IF EMPTY(m.fampop)
	RETURN
ENDIF	
SELECT famtab
tname = ALLTRIM(UPPER(name))
IF fox_alert(IIF(m.publang,'Delete '+tname+'?','Supprimer'+tname+'?'))
	DELETE
ELSE
	RETURN
ENDIF