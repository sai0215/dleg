PARAMETERS ttype

CLOS DATA
mmmdate = {//}
m.totpriced =0
m.totpayday =0
IF ttype=1
   DO mouredt1.spr
ELSE   
	IF ttype=2   
		DO mouredt2.spr
	ELSE
		DO mouredt3.spr
	ENDIF
ENDIF   
CLOS DATA


PROCEDURE sum1totsal
********************
SELECT mouhea
GO TOP
m.sumtotsal = 0
DO WHILE !EOF()
    IF mouhea.date >= m.date1 AND mouhea.date <= m.date2 AND mouhea.supcli = 2
       m.sumtotsal = m.sumtotsal+cal_cur(mouhea.dirsales-IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,1)
    ENDIF   
    SELECT mouhea
    SKIP
ENDDO    
SELECT mourow
RETURN m.sumtotsal

PROCEDURE sum1totret
********************
SELECT mouhea
GO TOP
m.sumtotsal = 0
DO WHILE !EOF()
    IF mouhea.date >= m.date1 AND mouhea.date <= m.date2 AND mouhea.supcli = 2
       m.sumtotsal = m.sumtotsal+cal_cur(IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirret-mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,1)
    ENDIF   
    SELECT mouhea
    SKIP
ENDDO    
SELECT mourow
RETURN m.sumtotsal

PROCEDURE sum2totret
********************
SELECT mouhea
GO TOP
m.sumtotsal = 0
DO WHILE !EOF()
    IF mouhea.date >= m.date1 AND mouhea.date <= m.date2 AND mouhea.supcli = 2
       m.sumtotsal = m.sumtotsal+cal_cur(IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirret-mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,2)
    ENDIF   
    SELECT mouhea
    SKIP
ENDDO    
SELECT mourow
RETURN m.sumtotsal

PROCEDURE sum2totsal
*******************
SELECT mouhea
GO TOP
m.sumtotsal = 0
DO WHILE !EOF()
    IF mouhea.date >= m.date1 AND mouhea.date <= m.date2 AND mouhea.supcli = 2
       m.sumtotsal = m.sumtotsal+cal_cur(mouhea.dirsales-IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,2)
    ENDIF   
    SELECT mouhea
    SKIP
ENDDO    
SELECT mourow
RETURN m.sumtotsal

PROCEDURE mysumtot
******************
SELECT mouhea
GO TOP
m.mysumtot = 0
DO WHILE !EOF()
    IF mouhea.date >= m.date1 AND mouhea.date <= m.date2 AND mouhea.supcli = 2 AND !FOUND("pmotab")
       m.mysumtot = m.mysumtot+cal_cur(mouhea.dirsales,mouhea.date,mouhea.curtab,2)
    ENDIF   
    SELECT mouhea
    SKIP
ENDDO    
SELECT mourow
RETURN m.mysumtot

PROCEDURE sum2totpay
********************
SELECT mouhea
GO TOP
m.sumtotpay = 0
DO WHILE !EOF()
    IF mouhea.date >= m.date1 AND mouhea.date <= m.date2 AND mouhea.supcli = 2
       m.sumtotpay = m.sumtotpay+mouhea.totpayed2
    ENDIF   
    SELECT mouhea
    SKIP
ENDDO    
SELECT mourow
RETURN m.sumtotpay

PROCEDURE sum1totpay
********************
SELECT mouhea
GO TOP
m.sumtotpay = 0
DO WHILE !EOF()
    IF mouhea.date >= m.date1 AND mouhea.date <= m.date2 AND mouhea.supcli = 2
       m.sumtotpay = m.sumtotpay+mouhea.totpayed1
    ENDIF   
    SELECT mouhea
    SKIP
ENDDO    
SELECT mourow
RETURN m.sumtotpay

PROCEDURE sum1totcre
********************
m.sumtotcre = 0
m.sumtotcre = sum1totsal() - sum1totpay()
RETURN m.sumtotcre 

PROCEDURE sum2totcre
********************
m.sumtotcre = 0
m.sumtotcre = sum2totsal() - sum2totpay()
RETURN m.sumtotcre

PROCEDURE totpayd
*****************
m.totpayday = 0
SELECT mouhea
m.serhhh = serial
SELECT mourow
SKIP
IF mourow.mouhea = m.serhhh
	SKIP -1
	RETURN ''
ENDIF	
SKIP -1	
SELECT mouhea
SKIP
IF mmmdate = mouhea.date
	SKIP -1
	mmmdate = mouhea.date
	SELECT mourow
	RETURN ''
ELSE
	SKIP -1
	mmmdate = mouhea.date
	xrecno = RECNO()
	xorder = ORDER()
	SET ORDER TO date
	SEEK mmmdate
	DO WHILE !EOF() AND mouhea.date=mmmdate
		m.totpayday = m.totpayday+totpayed2
		SKIP
	ENDDO
	SET ORDER TO &xorder
	IF xrecno > 0 AND xrecno <= RECCOUNT()
		GO xrecno
	ENDIF	
	SELECT mourow
	RETURN m.totpayday
ENDIF		

PROCEDURE totpriced
*******************
m.totpriced = 0
SELECT mouhea
m.serhhh = serial
SELECT mourow
IF !EOF()
	SKIP
ENDIF	
IF mourow.mouhea=m.serhhh
	IF !BOF()
		SKIP -1
	ENDIF	
	RETURN ''
ENDIF
IF !BOF()
	SKIP -1	
ENDIF	
SELECT mouhea
IF !EOF()
	SKIP
ENDIF	
IF mmmdate=mouhea.date
	IF !BOF()
		SKIP -1
	ENDIF
	mmmdate = mouhea.date
	SELECT mourow
	RETURN ''
ELSE
	IF !BOF()
		SKIP -1
	ENDIF
	mmmdate = mouhea.date
	xrecno  = RECNO()
	xorder  = ORDER()
	SET ORDER TO date
	SEEK mmmdate
	DO WHILE !EOF() AND mouhea.date=mmmdate
		m.totpriced = m.totpriced+cal_cur(mouhea.dirsales-IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,2)
		SELECT mouhea
		SKIP
	ENDDO
	SET ORDER TO &xorder
	IF xrecno>0 AND xrecno<=RECCOUNT()
		GO xrecno
	ENDIF	
	SELECT mourow
	RETURN m.totpriced
ENDIF

PROCEDURE totretd
*****************
m.totpriced = 0
SELECT mouhea
m.serhhh = serial
SELECT mourow
SKIP
IF mourow.mouhea = m.serhhh
	SKIP -1
	RETURN ''
ENDIF	
SKIP -1	
SELECT mouhea
SKIP
IF mmmdate = mouhea.date
	SKIP -1
	mmmdate = mouhea.date
	SELECT mourow
	RETURN ''
ELSE
	SKIP -1
	mmmdate = mouhea.date
	xrecno = RECNO()
	xorder = ORDER()
	SET ORDER TO date
	SEEK mmmdate
	DO WHILE !EOF() AND mouhea.date=mmmdate
		m.totpriced = m.totpriced+cal_cur(IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirret-mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,2)
		SELECT mouhea
		SKIP
	ENDDO
	SET ORDER TO &xorder
	IF xrecno > 0 AND xrecno <= RECCOUNT()
		GO xrecno
	ENDIF	
	SELECT mourow
	RETURN m.totpriced
ENDIF



PROCEDURE pricedate
*******************
SELECT mouhea
m.serhhh = serial
SELECT mourow
SKIP
IF mourow.mouhea = m.serhhh
	SKIP -1
	RETURN ''
ENDIF	
SKIP -1	
SELECT mouhea
SKIP
IF mmmdate = mouhea.date
	SKIP -1
	mmmdate = mouhea.date
	SELECT mourow
	RETURN ''
ELSE
	SKIP -1
	mmmdate = mouhea.date
	SELECT mourow
	RETURN "TOTAL PURCHASED FOR "+ALLTRIM(DTOC(mouhea.date)) 
ENDIF		

PROCEDURE retdate
*****************
SELECT mouhea
m.serhhh = serial
SELECT mourow
SKIP
IF mourow.mouhea = m.serhhh
	SKIP -1
	RETURN ''
ENDIF	
SKIP -1	
SELECT mouhea
SKIP
IF mmmdate = mouhea.date
	SKIP -1
	mmmdate = mouhea.date
	SELECT mourow
	RETURN ''
ELSE
	SKIP -1
	mmmdate = mouhea.date
	SELECT mourow
	RETURN "TOTAL RETURN FOR "+ALLTRIM(DTOC(mouhea.date)) 
ENDIF		


PROCEDURE payeddate
*******************
SELECT mouhea
m.serhhh = serial
SELECT mourow
SKIP
IF mourow.mouhea = m.serhhh
	SKIP -1
	RETURN ''
ENDIF	
SKIP -1	
SELECT mouhea
SKIP
IF mmmdate = mouhea.date
	SKIP -1
	mmmdate = mouhea.date
	SELECT mourow
	RETURN ''
ELSE
	SKIP -1
	mmmdate = mouhea.date
	SELECT mourow
	RETURN "TOTAL PAID FOR "+ALLTRIM(DTOC(mouhea.date)) 
ENDIF		


PROCEDURE totcurd
*****************
SELECT mouhea
m.serhhh = serial
SELECT mourow
SKIP
IF mourow.mouhea = m.serhhh
	SKIP -1
	RETURN ''
ENDIF	
SKIP -1	
SELECT mouhea
SKIP
IF mmmdate = mouhea.date
	SKIP -1
	mmmdate = mouhea.date
	SELECT mourow
	RETURN ''
ELSE
	SKIP -1
	mmmdate = mouhea.date
	SELECT mourow
	RETURN 'US$'
ENDIF		

PROCEDURE caltotprice
*********************
SELECT mourow
wait wind proret
RETURN ""
