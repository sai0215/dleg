CLOS DATA
report ='gftredt'
DO gftsedt.spr
CLOS DATA

PROCEDURE vblank
****************
=checkzapfile("gfttmp")
SELECT (alias)
m.supdes   = ''
m.saldes   = ''
m.bandes   = ''
m.curdes1  = ''
m.curdes2  = ''
m.bandes   = ''
m.lcurdes  = ''
m.totprice = 0
m.bnktotal = 0
m.curtotal = 0
SCATTER MEMVAR BLANK
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
SELECT gfttmp
APPEND BLANK
REPLACE gfthea   WITH m.serial
REPLACE gfttype  WITH 'N'
SHOW GETS
SELECT (alias)

PROCEDURE getdata
*****************
m.clides = clitab.name
m.saldes = saltab.name
SELECT (alias)
SCATTER MEMVAR
=checkzapfile("gfttmp")
SELECT gftrow
SEEK gfthea.serial
DO WHILE gfthea = gfthea.serial AND !EOF()
   SCATTER MEMVAR
   SELECT gfttmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE gfthea   WITH m.serial
   REPLACE name     WITH protab.name
   REPLACE des      WITH protab.number
   REPLACE gfttype  WITH 'U'
   REPLACE gftrecno WITH RECNO('gftrow')
   SELECT mourow
   m.qtysel=0
   SET ORDER TO gfthea
   SEEK gfthea.serial
   DO WHILE !EOF() AND mourow.gfthea = gfthea.serial 
   		IF mourow.protab = gfttmp.protab 
			SELECT gfttmp
			m.qtysel = m.qtysel + mourow.qty
			REPLACE qtysel WITH m.qtysel
			SELECT mourow
		ENDIF	
		SKIP
	ENDDO	
   SELECT gftrow
   m.qtysel=0
   SKIP
ENDDO
DO tottttp
SELECT gfttmp
APPEND BLANK
REPLACE gfthea  WITH m.serial
REPLACE gfttype WITH 'N'
SELECT gfttmp
SHOW GETS
SELECT (alias)

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
SET DELETED OFF
SELECT gfttmp
GO TOP
DO WHILE !EOF()
   DO CASE
      CASE gfttype = 'A'
           SCATTER MEMVAR
           SELECT gftrow
           APPEND BLANK
           GATHER MEMVAR
           REPLACE gfthea WITH m.serial
      CASE gfttype = 'U'
           IF gftrecno <> 0
              SCATTER MEMVAR
              SELECT gftrow
              GO RECORD m.gftrecno
              GATHER MEMVAR
              REPLACE gfthea WITH m.serial
           ENDIF
      CASE gfttype = 'D'
           IF gftrecno <> 0
              SCATTER MEMVAR
              SELECT gftrow
              GO RECORD m.gftrecno
              DELETE
           ENDIF
   ENDCASE
   SELECT gfttmp
   REPLACE gfttype WITH IIF(gfttype = 'A' , 'U', gfttype )
   SKIP        
ENDDO
SET DELETED ON
SELECT (alias)
GATHER MEMVAR
isadding = .F.

PROCEDURE search
*****************
SELECT (alias)
m.sfound = .F.
m.srecno = 0
IF issearch
   m.srecno = RECNO()
ENDIF

DIMENSION varary[3,3]
varary[1,1] = 'Reference'
varary[2,1] = 'Customer'
varary[3,1] = 'Exp. Date'

varary[1,2] = 'ref_no'
varary[2,2] = 'clitab'
varary[3,2] = 'date'

varary[1,3] = ''
varary[2,3] = 'name'
varary[3,3] = ''

SELECT (alias)
m.sfound = .F.
m.srecno = 0
IF issearch
   m.srecno = RECNO()
ENDIF
DO search3 WITH "gfthea", "!closed", 3, "Pending lists ONLY!"

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
SELECT gftrow
DELETE FOR gfthea = m.serial
SELECT (alias)
RETURN .T.

PROCEDURE prtrec
****************
SELECT (alias)
DO print WITH report,IIF(m.criteria=1,'RECNO()=m.saverec',IIF(m.criteria=2,'NOT EMPTY(bidon)','.T.')),IIF(m.print=1,'PREVIEW','TO PRINT')


PROCEDURE updprice
******************
SELECT gfttmp
GO TOP
DO WHILE !EOF()
   IF gfttype <> 'N'
      SELECT protab
      SEEK gfttmp.protab
      m.suprice = 0
      DO getuprice
      SELECT gfttmp
      REPLACE uprice WITH m.suprice
      REPLACE price  WITH uprice*qty
   ENDIF
   SKIP
ENDDO
SELECT (alias)

PROCEDURE getuprice
*******************
IF m.curtab = protab.curtab
   m.suprice = protab.selling
   SELECT protab
   RETURN
ENDIF
IF m.curtab = 1
   m.suprice = protab.selling1
   SELECT protab
   RETURN
ENDIF
SELECT rathea
LOCATE FOR date > m.date
IF EOF()
   GO BOTTOM
ELSE   
   SKIP -1
ENDIF   
m.rathea = serial
SELECT ratrow
SEEK STR(m.rathea)+STR(protab.curtab)
m.suprice = protab.selling*drate2
SEEK STR(m.rathea)+STR(m.curtab)
m.suprice = m.suprice*rate2
SELECT protab

PROCEDURE tottttp
*****************
m.xalias = ALIAS()
m.ttmourec  = RECNO()
m.totprice = 0
m.ttmourec = RECNO()
GO TOP
DO WHILE !EOF()
	m.totprice=m.totprice+price 
	SKIP
ENDDO		
SHOW GETS	
SELECT (xalias)
IF m.ttmourec > 0 AND m.ttmourec <=RECCOUNT()
   GO RECORD m.ttmourec
ENDIF   


PROCEDURE addcli
****************
SELECT clitab
torder = ORDER()
IF m.clitab = 0
   SET ORDER TO 0
   GO BOTTOM
   m.tserial = serial+1 
   SET ORDER TO &torder
   APPEND BLANK
   GATHER MEMVAR
   REPLACE serial WITH m.tserial
ELSE
   SET ORDER TO serial
   SEEK m.clitab
   IF FOUND()
      GATHER MEMVAR
      REPLACE serial WITH m.tserial
      m.clides = name
   ENDIF
   SET ORDER TO &torder
ENDIF   
m.tttsave = .T.
