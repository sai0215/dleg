PARAMETERS m.choice

IF m.choice <> 3
	DO itmsedt1.spr
ELSE
	DO delitm	
ENDIF	
SHOW GET m.itmpop

PROCEDURE newitm
****************
IF EMPTY(m.itemno)
	RETURN
ENDIF	
SELECT itmtab
m.serial = tserial + 1
m.ctgtab = ctgtab.serial
APPEND BLANK
GATHER MEMVAR MEMO

PROCEDURE edititm
*****************
SELECT itmtab
IF EMPTY(m.itemno) OR EOF()
	RETURN
ENDIF	
m.ctgtab = ctgtab.serial
GATHER MEMVAR MEMO

PROCEDURE delitm
****************
IF EMPTY(m.itmpop)
	RETURN
ENDIF	
SELECT itmtab
titem = ALLTRIM(UPPER(item))
IF fox_alert(IIF(m.publang,'Delete '+titem+'?','Supprimer'+titem+'?'))
	DELETE
ELSE
	RETURN
ENDIF