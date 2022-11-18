m.stat       = ''
itemsold     = 0
tttprice     = 0
tttitem      = 0
tttpriceitem = 0
tttperc      = 0
xxxprice     = 0
xxxitem      = 0
xxxpriceitem = 0
xxxperc      = 0
zzzitem      = 0
DO stasedt.spr

PROCEDURE statday
*****************
= checkzapfile("statmp1")
APPEND BLANK
REPLACE pday WITH "Monday"
APPEND BLANK
REPLACE pday WITH "Tuesday"
APPEND BLANK
REPLACE pday WITH "Wednesday"
APPEND BLANK
REPLACE pday WITH "Thursday"
APPEND BLANK
REPLACE pday WITH "Friday"
APPEND BLANK
REPLACE pday WITH "Saturday"
APPEND BLANK
REPLACE pday WITH "Sunday"
SELECT mouhea
SET ORDER TO date
LOCATE FOR date >= m.date1
DO WHILE !EOF() AND date <= m.date2
	SELECT statmp1
	LOCATE FOR UPPER(ALLTRIM(pday)) = UPPER(ALLTRIM(CDOW(mouhea.date)))
	IF FOUND()
		REPLACE nbfact    WITH nbfact+1
**		REPLACE prixtotal WITH prixtotal+cal_cur(mouhea.totprice,mouhea.date,mouhea.curtab,2)
		REPLACE prixtotal WITH prixtotal+cal_cur(mouhea.dirsales-IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,2)
		SELECT mourow
		SET ORDER TO mouhea
		SEEK mouhea.serial
		m.nbprod = 0
		DO WHILE mourow.mouhea = mouhea.serial
			IF EMPTY(proret)
				m.nbprod = m.nbprod + qty
			ELSE
				m.nbprod = m.nbprod - qty
			ENDIF	
			SKIP
		ENDDO	
		SELECT statmp1
		REPLACE nbprod WITH nbprod + m.nbprod
	ENDIF
	SELECT mouhea
	SKIP
ENDDO	
SELECT statmp1
SUM nbfact TO m.totfact
GO TOP
DO WHILE !EOF()
	REPLACE prix_fact WITH IIF(!EMPTY(nbfact),prixtotal/nbfact,prixtotal)
	REPLACE prod_fact WITH IIF(!EMPTY(nbfact),nbprod/nbfact,nbprod)
	REPLACE perc      WITH IIF(!EMPTY(m.totfact),100*nbfact/m.totfact,0)
	SKIP
ENDDO	
			
PROCEDURE stattime
******************
= checkzapfile("statmp1")
SELECT mouhea
SET ORDER TO time
GO TOP
m.begint = time
GO BOTTOM
m.endt = time
FOR i=VAL(SUBSTR(m.begint,1,2)) TO VAL(SUBSTR(m.endt,1,2))
	SELECT statmp1
	APPEND BLANK
	REPLACE ptime WITH PADL(ALLTRIM(STR(I)),2,"0")
ENDFOR
SELECT mouhea
SET ORDER TO date
LOCATE FOR date >= m.date1
DO WHILE !EOF() AND date <= m.date2
	SELECT statmp1
	LOCATE FOR VAL(ptime) = VAL(SUBSTR(mouhea.time,1,2))	
	IF FOUND()
		REPLACE nbfact    WITH nbfact+1
**		REPLACE prixtotal WITH prixtotal+cal_cur(mouhea.totprice,mouhea.date,mouhea.curtab,2)
		REPLACE prixtotal WITH prixtotal+cal_cur(mouhea.dirsales-IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,2)
		SELECT mourow
		SET ORDER TO mouhea
		SEEK mouhea.serial
		m.nbprod = 0
		DO WHILE mourow.mouhea = mouhea.serial
			IF EMPTY(proret)
				m.nbprod = m.nbprod + qty
			ELSE
				m.nbprod = m.nbprod - qty
			ENDIF	
			SKIP
		ENDDO	
		SELECT statmp1
		REPLACE nbprod WITH nbprod + m.nbprod
	ENDIF
	SELECT mouhea
	SKIP
ENDDO	
SELECT statmp1
SUM nbfact TO m.totfact
GO TOP
DO WHILE !EOF()
	REPLACE prix_fact WITH IIF(!EMPTY(nbfact),prixtotal/nbfact,prixtotal)
	REPLACE prod_fact WITH IIF(!EMPTY(nbfact),nbprod/nbfact,nbprod)
	REPLACE perc      WITH IIF(!EMPTY(m.totfact),100*nbfact/m.totfact,0)
	SKIP
ENDDO	

PROCEDURE statfam
*****************
= checkzapfile("statmp2")
SET ORDER TO fl1tab
SELECT protab
SET RELATION OFF INTO famtab1
SET RELATION OFF INTO fl1tab
SET RELATION TO protab.famtab1 INTO famtab1 ADDITIVE
SET RELATION TO protab.fl1tab  INTO fl1tab ADDITIVE
SET ORDER TO bidon
SEEK "X"
DO WHILE !EOF()
	SELECT statmp2
	SEEK fl1tab.name
	IF !FOUND() AND !EMPTY(fl1tab.name)
		APPEND BLANK
		REPLACE fl1tab  WITH fl1tab.name
		REPLACE number2 WITH fl1tab.number
	ENDIF
	REPLACE famtab1 WITH famtab1.name
	REPLACE number1 WITH famtab1.number
	m.nbitem = 0
	m.prixt  = 0
	SELECT mourow
	SET ORDER TO protab
	SEEK protab.serial
	DO WHILE !EOF() AND mourow.protab = protab.serial
		SELECT mouhea
		SET ORDER TO serial
		SEEK mourow.mouhea
		IF date >= m.date1 AND date <= m.date2
			IF EMPTY(mourow.proret)
				m.nbitem = m.nbitem + mourow.qty
				m.prixt  = m.prixt + cal_cur(mourow.price,mouhea.date,mouhea.curtab,2)
			ELSE
				m.nbitem = m.nbitem - mourow.qty
				m.prixt  = m.prixt - cal_cur(mourow.price,mouhea.date,mouhea.curtab,2)
			ENDIF	
		ENDIF
		SELECT mourow
		SKIP
	ENDDO	
	SELECT statmp2
	REPLACE	totitem  WITH totitem + m.nbitem 
	REPLACE totprice WITH totprice + m.prixt
	SELECT protab
	SKIP
ENDDO
SELECT statmp2
SET ORDER TO famtab1
SUM totitem  TO itemsold 
SUM totprice TO tttprice
tttpriceitem = tttprice/itemsold
GO TOP

PROCEDURE xxxitem
*****************
zzzitem = 0
SELECT statmp2
xrecc     = RECNO()
m.famtab1 = famtab1
SUM totitem  TO xxxitem  FOR famtab1 = m.famtab1
SUM totprice TO xxxprice FOR famtab1 = m.famtab1
xxxpriceitem = IIF(!EMPTY(xxxitem),xxxprice/xxxitem,0)
xxxperc = (xxxitem*100)/itemsold
IF xrecc > 0 AND xrecc <= RECCOUNT()
	GO xrecc
	IF !EMPTY(xxxitem) 
		zzzitem = (totprice*100)/xxxprice
	ENDIF	
ENDIF
RETURN zzzitem
