SELECT invent
GO TOP
DO WHILE !EOF()
	m.bar = ALLTRIM(barcode)
	m.inv = endq
	SELECT protab
	IF SEEK(m.bar)
		REPLACE inv97 WITH inv97+m.inv	
	ENDIF
	SELECT invent	
	SKIP
ENDDO
return

SELECT protab
GO TOP
DO WHILE !EOF()
	m.ret = getret(serial)
	SELECT protab
	REPLACE ret98 WITH m.ret
	SKIP
ENDDO
return

PROCEDURE getret
****************
PARAMETERS ttprotab

m.qtyret = 0
SELECT mouhea
SET ORDER TO serial
SELECT mourow
SET RELATION TO mourow.mouhea INTO mouhea ADDITIVE
SET ORDER TO protab
SEEK ttprotab
DO WHILE protab=ttprotab AND !EOF()
	IF mouhea.supcli<>1
		IF !EMPTY(mourow.proret) 
			m.qtyret = m.qtyret+qty
		ENDIF    
	ENDIF   
    SKIP
ENDDO
SET RELATION OFF INTO mouhea
RETURN m.qtyret


RETURN


SELECT invent
GO TOP
DO WHILE !EOF()
	m.bar = ALLTRIM(barcode)
	m.inv = endq
	SELECT protab
	IF SEEK(m.bar)
		REPLACE inv97 WITH m.inv	
	ENDIF
	SELECT invent	
	SKIP
ENDDO

RETURN

SELECT protab
GO TOP
DO WHILE !EOF()
	m.stk = getqty(serial)
	SELECT protab
	REPLACE stk97 WITH m.stk
	SKIP
ENDDO


PROCEDURE getqty
****************
PARAMETERS ttprotab

m.qtyent = 0
m.qtyext = 0
SELECT arrrow
SET ORDER TO protab
SEEK ttprotab
DO WHILE protab=ttprotab AND !EOF()
   m.qtyent = m.qtyent+qty
   SKIP
ENDDO
SELECT mouhea
SET ORDER TO serial
SELECT mourow
SET RELATION TO mourow.mouhea INTO mouhea ADDITIVE
SET ORDER TO protab
SEEK ttprotab
DO WHILE protab=ttprotab AND !EOF()
	IF mouhea.supcli=1
   		m.qtyent = m.qtyent+qty
	ELSE
*   	m.qtyext = m.qtyext+IIF(mouhea.return,-qty,qty)   
		IF !EMPTY(mourow.proret) 
			m.qtyext = m.qtyext-qty
		ELSE
		    m.qtyext = m.qtyext+(qty+free)
		ENDIF    
	ENDIF   
    SKIP
ENDDO
SET RELATION OFF INTO mouhea
RETURN (m.qtyent-m.qtyext)

RETURN


SELECT mourow
GO TOP
DO WHILE !EOF()
	m.bar = ALLTRIM(des)
	m.qty = qty
	SELECT invent
	SET FILT TO wh='MG'	
	IF SEEK(m.bar)
		IF !EMPTY(code)
			REPLACE qsold WITH qsold+m.qty			
		ENDIF	
	ELSE
		SET FILT TO wh='FW'
		IF SEEK(m.bar)
			REPLACE qsold WITH qsold+m.qty		
		ELSE
			SET FILT TO wh='SA'
			IF SEEK(m.bar)
				REPLACE qsold WITH qsold+m.qty		
*			ELSE
*				APPEND BLANK
*				REPLACE wh      WITH 'MG'
*				REPLACE barcode WITH m.bar
*				REPLACE qsold   WITH qsold+m.qty
			ENDIF
		ENDIF	
	ENDIF
	SELECT mourow
	SKIP
ENDDO
? s


RETURN
SELECT invent
GO TOP
DO WHILE !EOF()
	m.bar = barcode
	SELECT protab
	IF SEEK(m.bar)
		m.name    = name
		m.selling = selling
		SELECT invent		
		REPLACE desc    WITH m.name
		REPLACE selling WITH m.selling	
	ENDIF
	SELECT invent
	SKIP
ENDDO
