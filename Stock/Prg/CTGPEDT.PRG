PARAMETERS m.choice

IF m.choice <> 3
	DO ctgsedt1.spr
ELSE
	DO delctg	
ENDIF	
SHOW GET m.ctgpop

PROCEDURE newctg
****************
IF EMPTY(m.code)
	RETURN
ENDIF	
SELECT ctgtab
m.serial = getmax('ctgtab','serial') + 1
APPEND BLANK
GATHER MEMVAR

PROCEDURE editctg
*****************
SELECT ctgtab
IF EMPTY(m.code) OR EOF()
	RETURN
ENDIF	
GATHER MEMVAR

PROCEDURE delctg
****************
IF EMPTY(m.ctgpop)
	RETURN
ENDIF	
SELECT ctgtab
tcategory = ALLTRIM(UPPER(category))
IF fox_alert(IIF(m.publang,'Delete '+tcategory+'?','Supprimer'+tcategory+'?'))
	DELETE
ELSE
	RETURN
ENDIF