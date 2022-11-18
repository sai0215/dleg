WAIT WINDOW "Insert the disk containing the file 'Newarr.dbf' into drive A:\. Press any key to continue..."
cfile = ALLTRIM(m.pubdir)+'dbf\newarr.dbf'
drivea = .T.
COPY FILE a:\newarr.dbf TO &cfile
IF !drivea
	RETURN
ENDIF
IF !USED('protab')
	SELECT 0
	USE protab
ENDIF
IF !USED('famtab')
	SELECT 0
	USE famtab
ENDIF
IF !USED('famtab')
	SELECT 0
	USE famtab
ENDIF
IF !USED('famtab1')
	SELECT 0
	USE famtab1
ENDIF
IF !USED('fl4tab')	
	SELECT 0
	USE fl4tab
ENDIF
IF !USED('rathea')
	SELECT 0
	USE rathea
ENDIF
IF !USED('ratrow')
	SELECT 0
	USE ratrow
ENDIF
IF !USED('genpar')
	SELECT 0
	USE genpar
ENDIF
IF !USED('newarr')	
	SELECT 0
	USE newarr.dbf
ENDIF
m.newcnt = 0
m.serial = getser()
SELECT famtab
GO BOTTOM
m.famtab = serial
SELECT newarr
m.reccnt = ALLTRIM(STR(RECCOUNT()))
GO TOP
DO WHILE !EOF()
	m.number  = ALLTRIM(barcode)
	m.barcode = m.number
	WAIT WINDOW 'Processing: '+ALLTRIM(product) NOWAIT
	SELECT protab
	SET ORDER TO barcode
	IF !SEEK(m.barcode)
		SELECT newarr
		m.serial   = m.serial+1
		m.name     = product
**		fifty      = price+(price*0.5)
**		fifty      = price+(price*0.4)
**		eight      = fifty+(fifty*0.8)
**		m.selling  = ROUND(((((price*0.1)+eight)/6)+1)*3.65,0)
**		m.selling  = ROUND(cal_cur(price*genpar.rate,DATE(),3,1),0)
**		m.selling  = ROUND(price*genpar.rate,0)
		m.selling  = uprice
		m.selling1 = m.selling
		m.number   = ALLTRIM(barcode)
		m.field2   = model
		m.famtab   = m.famtab
		m.famtab1  = genpar.famtab1
		m.des      = ALLTRIM(m.name)+'/'+ALLTRIM(m.field2)+'/'+m.number
		m.spacloc  = RAT(' ',ALLTRIM(size))
		tsize      = ALLTRIM(SUBSTR(size,m.spacloc))
		m.fl4tab   = getsize(tsize)
		m.curtab   = 1
		m.adjqty   = 0
		SELECT protab
		APPEND BLANK
		GATHER MEMVAR MEMO
		m.newcnt = m.newcnt+1
	ELSE
		m.famtab1  = genpar.famtab1
		REPLACE famtab1 WITH m.famtab1
	ENDIF
	SELECT newarr
	SKIP
ENDDO
CLOSE DATA

DO neword
WAIT WINDOW 'Stock is updated: '+m.reccnt+' Items, '+ALLTRIM(STR(m.newcnt))+' New' NOWAIT

PROCEDURE getsize
*****************
PARAMETERS tsize

IF EMPTY(tsize)
	RETURN 0
ENDIF	
tselect = SELECT()
SELECT fl4tab
LOCATE FOR name=tsize
IF FOUND()
	tser = serial
ELSE
	SET ORDER TO 0
	GO BOTTOM
	tser = serial+1
	APPEND BLANK
	REPLACE serial WITH tser
	REPLACE number WITH ALLTRIM(STR(tser))
	REPLACE name   WITH ALLTRIM(tsize)
ENDIF
SELECT (tselect)
RETURN tser

PROCEDURE getser
****************
tselect = SELECT()
torder  = ORDER()
SELECT protab
SET ORDER TO serial
GO BOTTOM
tserial = serial
SET ORDER TO (torder)
SELECT (tselect)
RETURN tserial