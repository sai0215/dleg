CLOS DATA
DO mouredt4.spr
CLOS DATA

PROCEDURE totpriced
*******************
m.totpriced = 0
m.date = dattmp.date
SELECT mouhea
SUM cal_cur(mouhea.dirsales-IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,2) TO m.totpriced FOR mouhea.date=m.date AND mouhea.dirsales<>0 AND mouhea.supcli=2
SELECT dattmp
RETURN m.totpriced

PROCEDURE totpayd
*****************
m.totpayday = 0
m.date = dattmp.date
SELECT mouhea
SUM totpayed2 TO m.totpayday FOR mouhea.date=m.date AND mouhea.supcli=2
SELECT dattmp
RETURN m.totpayday

PROCEDURE sum1totsal
*******************
SELECT mouhea
m.sumtotsal = 0
SUM cal_cur(mouhea.dirsales-IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,1) TO m.sumtotsal FOR mouhea.date>=m.date1 AND mouhea.date<=m.date2 AND mouhea.supcli=2 AND mouhea.dirsales<>0
RETURN m.sumtotsal

PROCEDURE sum2totsal
*******************
SELECT mouhea
m.sumtotsal = 0
SUM cal_cur(mouhea.dirsales-IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,2) TO m.sumtotsal FOR mouhea.date>=m.date1 AND mouhea.date<=m.date2 AND mouhea.supcli=2 AND mouhea.dirsales<>0
RETURN m.sumtotsal

PROCEDURE sum1totpay
********************
SELECT mouhea
m.sumtotpay = 0
SUM mouhea.totpayed1 TO m.sumtotpay FOR mouhea.date>=m.date1 AND mouhea.date<=m.date2 AND mouhea.supcli=2
RETURN m.sumtotpay

PROCEDURE sum2totpay
********************
SELECT mouhea
m.sumtotpay = 0
SUM mouhea.totpayed2 TO m.sumtotpay FOR mouhea.date>=m.date1 AND mouhea.date<=m.date2 AND mouhea.supcli=2
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

PROCEDURE sum1totret
********************
SELECT mouhea
m.sumtotsal = 0
SUM cal_cur(IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirret-mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,1) TO m.sumtotsal FOR mouhea.date>=m.date1 AND mouhea.date<=m.date2 AND mouhea.supcli=2
RETURN m.sumtotsal

PROCEDURE sum2totret
********************
SELECT mouhea
m.sumtotsal = 0
SUM cal_cur(IIF(mouhea.dirret>mouhea.dirsales,mouhea.dirret-mouhea.dirsales,mouhea.dirret),mouhea.date,mouhea.curtab,2) TO m.sumtotsal FOR mouhea.date>=m.date1 AND mouhea.date<=m.date2 AND mouhea.supcli=2
RETURN m.sumtotsal

PROCEDURE totcurd
*****************
RETURN 'US$'