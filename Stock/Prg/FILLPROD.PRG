RETURN
SET DEFA TO ..\dbf
CLOS ALL

SELECT 0
USE protab
ZAP
SELECT 0
USE protab1
ZAP
SELECT 0
USE famtab
ZAP
SELECT 0
USE famtab1
ZAP
SELECT 0
USE fl1tab
ZAP
SELECT 0
USE fl2tab
ZAP
SELECT 0
USE fl3tab
ZAP
SELECT 0
USE fl4tab
ZAP

SELECT 0
USE prenatal

DO WHILE NOT EOF()
	SCATTER MEMVAR
	SELECT protab
	APPEND BLANK
	REPLACE serial WITH RECNO()
	REPLACE name WITH m.descriptio
	REPLACE field1 WITH m.modele
	REPLACE field2 WITH m.theme__
	REPLACE number WITH STR(m.ref_nb)
	REPLACE barcode WITH STR(m.ref_nb)
	REPLACE qty WITH m.qty
	REPLACE sugsel WITH m.prix_de_ve
	
	SELECT famtab
	LOCATE FOR name = m.saiso__
	IF NOT FOUND()
		APPEND BLANK
		REPLACE serial WITH RECNO()
		REPLACE number WITH STR(RECNO())
		REPLACE name WITH m.saiso__
	ENDIF
	SELECT protab
	REPLACE famtab WITH famtab.serial

	SELECT famtab1
	LOCATE FOR name = m.ray__ AND famtab = famtab.serial
	IF NOT FOUND()
		APPEND BLANK
		REPLACE serial WITH RECNO()
		REPLACE famtab WITH famtab.serial
		REPLACE number WITH m.ray__
		REPLACE name WITH m.ray__
	ENDIF
	SELECT protab
	REPLACE famtab1 WITH famtab1.serial

	SELECT fl1tab
	LOCATE FOR name = m.fam__
	IF NOT FOUND()
		APPEND BLANK
		REPLACE serial WITH RECNO()
		REPLACE number WITH m.fam__
		REPLACE name WITH m.fam__
	ENDIF
	SELECT protab
	REPLACE fl1tab WITH fl1tab.serial

	SELECT fl2tab
	LOCATE FOR name = m.sf__ AND fl1tab = fl1tab.serial 
	IF NOT FOUND()
		APPEND BLANK
		REPLACE serial WITH RECNO()
		REPLACE fl1tab WITH fl1tab.serial 
		REPLACE number WITH m.sf__
		REPLACE name WITH m.sf__
	ENDIF
	SELECT protab
	REPLACE fl2tab WITH fl2tab.serial

	SELECT fl3tab
	LOCATE FOR name = m.color__
	IF NOT FOUND()
		APPEND BLANK
		REPLACE serial WITH RECNO()
		REPLACE number WITH STR(RECNO())
		REPLACE name WITH m.color__
	ENDIF
	SELECT protab
	REPLACE fl3tab WITH fl3tab.serial

	SELECT fl4tab
	LOCATE FOR name = m.taille
	IF NOT FOUND()
		APPEND BLANK
		REPLACE serial WITH RECNO()
		REPLACE number WITH STR(RECNO())
		REPLACE name WITH m.taille
	ENDIF
	SELECT protab
	REPLACE fl4tab WITH fl4tab.serial

	SELECT protab1
	APPEND BLANK
	REPLACE protab WITH RECNO()
	REPLACE supplier WITH 'PRENATAL'
	REPLACE price WITH m.pvp_
	REPLACE currency WITH 'FF'
		
	SELECT prenatal
	SKIP
ENDDO		