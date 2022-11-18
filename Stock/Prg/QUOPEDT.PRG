CLOS DATA
report ='quoredt'
DO quosedt.spr
CLOS DATA

PROCEDURE vblank
****************
SELECT quotmp
ZAP
APPEND BLANK
SELECT grptmp
ZAP
SELECT (alias)
m.clides = ''
m.curdes = ''
m.saldes = ''
m.ldiscount = 0
m.totprice = 0
m.curgrp = 0
SCATTER MEMO MEMVAR BLANK
IF isadding
   torder = ORDER()
   trecno = RECNO()
   SET ORDER TO 0
   GO BOTTOM
   m.serial = serial+1
   SET ORDER TO &torder
   IF trecno > 0 AND trecno <= RECCOUNT()
      GO RECORD trecno
   ENDIF   
   IF isadding
      m.date = DATE()
   ENDIF   
ENDIF
m.changetot = ''
SELECT (alias)
SHOW GETS

PROCEDURE getdata
*****************
SELECT saltab
LOCATE FOR serial = quohea.saltab
m.saldes  = name
SELECT curtab
LOCATE FOR serial = quohea.curtab
m.curdes  = number
m.lcurdes = number
SELECT clitab
LOCATE FOR serial = quohea.clitab
m.clides = name
SELECT (alias)
SCATTER MEMVAR MEMO
m.ldiscount = m.discount
m.ltotprice = m.totprice
DO getallrow
DO getgrp
IF STR(m.totprice,10,2) = STR(m.ltotprice,10,2)
   m.changetot = ''
ELSE
   m.changetot = 'X'
ENDIF
m.totprice  = m.ltotprice
SELECT (alias)
SHOW GETS

PROCEDURE savedata
******************
SELECT (alias)
IF isadding
   torder = ORDER()
   SET ORDER TO 0
   GO BOTTOM
   m.serial = serial+1
   SET ORDER TO &torder
   APPEND BLANK
ENDIF
SELECT quogrp
SEEK m.serial
DO WHILE quohea = m.serial AND !EOF()
   DELETE
   SKIP
ENDDO
SELECT quorow
SEEK STR(m.serial)
DO WHILE quohea = m.serial AND !EOF()
   DELETE
   SKIP
ENDDO
SELECT quogrp
m.ttquoord = ORDER()
SET ORDER TO 0
SELECT grptmp
GO TOP
DO WHILE !EOF()
   SCATTER MEMVAR
   SELECT quogrp
   APPEND BLANK
   GATHER MEMVAR 
   IF quogrp = 0
      SKIP -1
      m.quogrp = m.quogrp + 1
      GO BOTTOM
      REPLACE quogrp WITH m.quogrp
   ENDIF
   SELECT grptmp
   SKIP
ENDDO
SELECT quogrp
SET ORDER TO &ttquoord
SELECT quotmp
DELETE FOR EMPTY(quohea)
SELECT quorow
APPEND FROM quotmp
SELECT (alias)
GATHER MEMVAR MEMO
isadding = .F.

PROCEDURE search
*****************
SELECT (alias)
m.sfound = .F.
m.srecno = 0
IF issearch
   m.srecno = RECNO()
ENDIF
DO trasearch.spr
IF m.srecno<>0 AND !m.sfound
   GO TOP
   IF NOT EOF()
      GO RECORD m.srecno
   ENDIF   
ENDIF
IF m.sfound
   m.issearch = .T.
   DO getdata
ENDIF

PROCEDURE okdel
***************
RETURN .T.

PROCEDURE prtrec
***************
SELECT (alias)
DO print WITH report,IIF(m.criteria=1,'RECNO()=m.saverec',IIF(m.criteria=2,'NOT EMPTY(bidon)','.T.')),IIF(m.print=1,'PREVIEW','TO PRINT')

PROCEDURE caldisc
*****************
RETURN
SELECT ordtmp
m.ttordrec = RECNO()
GO TOP
SUM price  TO m.totprice
m.totprice = (m.totprice*m.discount)/100
SELECT ordtmp
IF m.ttordrec > 0 AND m.ttordrec <=RECCOUNT()
   GO RECORD m.ttordrec
ENDIF   

PROCEDURE getgrp
****************
SELECT grptmp
ZAP
SELECT quogrp
SEEK quohea.serial
DO WHILE quohea = quohea.serial AND !EOF()
   SCATTER MEMVAR
   SELECT grptmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE name WITH typtab.name
   DO gettotal
   SELECT quogrp
   SKIP
ENDDO
SELECT grptmp
GO TOP
REPLACE bidon WITH '>'
m.curgrp = quogrp
DO getrow

PROCEDURE getallrow
*******************
SELECT quotmp
ZAP
SELECT quorow
SEEK STR(m.serial)
DO WHILE quohea=m.serial AND !EOF()
   SCATTER MEMVAR
   SELECT quotmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE name WITH protab.name
   SELECT quorow
   SKIP
ENDDO
SELECT quotmp
APPEND BLANK

PROCEDURE getrow
*****************
SELECT quotmp
SET FILTER TO EMPTY(quohea) OR quogrp = m.curgrp

PROCEDURE gettotal
******************
SELECT quotmp
SET FILTER TO
SUM fprice TO m.ttotal FOR quogrp = grptmp.quogrp
SET FILTER TO EMPTY(quohea) OR quogrp = m.curgrp
SELECT grptmp
REPLACE total WITH m.ttotal
DO getalltot

PROCEDURE getalltot
*******************
m.changetot = ''
SELECT grptmp
m.ttgrprec = RECNO()
GO TOP
SUM total  TO m.totprice
m.totprice = (m.totprice*(100-m.discount))/100
IF m.ttgrprec > 0 AND m.ttgrprec <=RECCOUNT()
   GO RECORD m.ttgrprec
ENDIF   
