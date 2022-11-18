CLOS DATA
report ='proredt'
varserial = 0
m.delqty  = 0
DO prosedt.spr
CLOS DATA

PROCEDURE vblank
****************
m.famdes = ''
m.famdes1 = ''
m.curdes = ''
m.fl1des = ''
m.fl2des = ''
m.fl3des = ''
m.fl4des = ''
m.curpro = 0
m.qtyini = 0
m.qtyent = 0
m.qtyext = 0
m.delqty = 0
SELECT protmp
= checkzapfile('protmp')
APPEND BLANK
REPLACE protab   WITH m.serial
REPLACE protype  WITH 'N'
SHOW GETS
SELECT (alias)
SCATTER MEMO MEMVAR BLANK
SHOW GETS

PROCEDURE getdata
*****************
*SELECT famtab
*LOCATE FOR serial = protab.famtab
m.famdes = famtab.name
varserial = famtab.serial
SELECT curtab
LOCATE FOR serial = protab.curtab
m.curdes = number
m.curpro = serial
*SELECT famtab1
*LOCATE FOR famtab=varserial
m.famdes1 = famtab1.name
m.fl1des = fl1tab.name
m.fl2des = fl2tab.name
m.fl3des = fl3tab.name
m.fl4des = fl4tab.name
SELECT protmp
= checkzapfile('protmp')
SELECT protab1
SEEK protab.serial
DO WHILE protab = protab.serial AND !EOF()
   SCATTER MEMVAR
   SELECT protmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE protab   WITH m.serial
   REPLACE suptab   WITH m.suptab
   REPLACE supname  WITH suptab.name
   REPLACE curtab   WITH m.curtab
   REPLACE curname  WITH curtab.number
   REPLACE protype  WITH 'U'
   REPLACE prorecno WITH RECNO('protab1')
   SELECT protab1
   SKIP
ENDDO
SELECT protmp
APPEND BLANK
REPLACE protab  WITH m.serial
REPLACE protype WITH 'N'
SELECT (alias)
SCATTER MEMVAR MEMO
SELECT rathea
GO BOTTOM
SELECT ratrow
SEEK STR(rathea.serial)+STR(1)
m.sellingll = INT((m.selling*rate2+500)/1000)*1000
SELECT (alias)
DO getqty
DO calcqty
SHOW GETS
SELECT protmp
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
   m.des = IIF(EMPTY(m.des),m.name+' '+m.field1+' '+m.field2,m.des)
   APPEND BLANK
ENDIF
SET DELETED OFF
SELECT protmp
GO TOP
DO WHILE !EOF()
   DO CASE
      CASE protype = 'A'
           SCATTER MEMVAR
           SELECT protab1
           APPEND BLANK
           GATHER MEMVAR
           REPLACE protab WITH m.serial
      CASE protype = 'U'
           IF prorecno <> 0
              SCATTER MEMVAR
              SELECT protab1
              GO RECORD m.prorecno
              GATHER MEMVAR
              REPLACE protab WITH m.serial
           ENDIF
      CASE protype = 'D'
           IF prorecno <> 0
              SCATTER MEMVAR
              SELECT protab1
              GO RECORD m.prorecno
              DELETE
           ENDIF
   ENDCASE
   SELECT protmp
   REPLACE protype WITH IIF(protype = 'A' , 'U', protype )
   SKIP        
ENDDO
SET DELETED ON
SELECT (alias)
GATHER MEMVAR MEMO
REPLACE curtab WITH m.curpro
isadding = .F.
SELECT (alias)
SHOW GETS

PROCEDURE search
*****************
DIMENSION varary[9,3]
varary[1,1] = 'Family'
varary[2,1] = 'Sub-family'
varary[3,1] = 'Theme'
varary[4,1] = 'Model'
varary[5,1] = 'Color'
varary[6,1] = 'Size'
varary[7,1] = 'Line'
varary[8,1] = 'barcode'
varary[9,1] = 'Label'

varary[1,2] = 'FL1TAB'
varary[2,2] = 'FL2TAB'
varary[3,2] = 'FIELD1'
varary[4,2] = 'FIELD2'
varary[5,2] = 'FL3TAB'
varary[6,2] = 'FL4TAB'
varary[7,2] = 'FAMTAB1'
varary[8,2] = 'NUMBER'
varary[9,2] = 'NAME'

varary[1,3] = 'name'
varary[2,3] = 'name'
varary[3,3] = ''
varary[4,3] = ''
varary[5,3] = 'name'
varary[6,3] = 'name'
varary[7,3] = 'name'
varary[8,3] = ''
varary[9,3] = ''

SELECT (alias)
m.sfound = .F.
m.srecno = 0
IF issearch
   m.srecno = RECNO()
ENDIF
DO spepedt WITH "protab", "itmsel=.T.",9, "Selected Items ONLY!","SUBSTR(number,1,13)+' '+SUBSTR(name,1,15)+' '+SUBSTR(fl1tab.name,1,10)+' '+SUBSTR(fl3tab.name,1,10)+' '+SUBSTR(fl4tab.name,1,10)",65
m.sfound = .T.
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
SELECT protab
SET FILTER TO

PROCEDURE okdel
***************
RETURN .T.

PROCEDURE warinit
*****************
SELECT warrow
SET ORDER TO warhea
SET FILTER TO protab = m.serial
SELECT warhea
GO TOP
DO WHILE !EOF()
   SELECT warrow
   SEEK warhea.serial
   IF !FOUND()
      APPEND BLANK
      REPLACE protab WITH m.serial
      REPLACE name   WITH warhea.name
      REPLACE warhea WITH warhea.serial
   ELSE   
      REPLACE name   WITH warhea.name
   ENDIF
   SELECT warhea
   SKIP
ENDDO
SELECT warrow
DEFINE WINDOW warbrow FROM 11,19 TO 30,52.5 TITLE 'Initial Stock' FONT 'Ms Sans Serif',8
MOVE WINDOW warbrow CENTER
ACTIVATE WINDOW warbrow
SET FILTER TO protab = m.serial
GO TOP
BROW FIELD name:R:H='Warehouse':35,qtyini:H='Init. qty' NOAPPEND NODELETE
SET FILTER TO
RELEASE WINDOW warbrow
DO calcqty
SELECT protmp
SET ORDER TO 0
SHOW GETS
SELECT protmp
SET ORDER TO protab
SELECT (alias)

PROCEDURE calcqty
*****************
SELECT warrow
SET ORDER TO protab
SEEK m.serial
m.qtyini = 0
DO WHILE protab=m.serial AND !EOF()
   m.qtyini = m.qtyini + qtyini
   SKIP
ENDDO
SELECT (alias)


PROCEDURE prtrec
***************
SELECT prttab
DELETE FOR EMPTY(bidon)
SELECT (alias)
SET RELATION TO serial INTO Prttab 
DO print WITH report,IIF(m.criteria=1,'RECNO()=m.saverec',IIF(m.criteria=2,'FOUND("prttab")','.T.')),IIF(m.print=1,'PREVIEW','TO PRINT')
SELECT (alias)
IF m.saverec <> 0
	GO m.saverec
ENDIF	

PROCEDURE prtdata
*****************
WAIT WINDOW 'Feature not available' NOWAIT
RETURN
SELECT (alias)
STORE RECNO() TO m.saverec
IF !USED('prttab')
	SELECT 0
	USE prttab
ELSE
	SELECT prttab
ENDIF	
SET ORDER TO serial
= checkzapfile('prttab')
APPEND FROM (alias)
IF m.saverec <> 0
	GO m.saverec
ENDIF	
DO printall.spr

PROCEDURE getqty
****************
*m.qtyent = 0
m.qtyent = protab.adjqty
m.qtyext = 0
SELECT arrrow
SEEK m.serial
DO WHILE m.serial = protab AND !EOF()
   m.qtyent = m.qtyent+qty
   SKIP
ENDDO
SELECT mourow
SEEK m.serial
DO WHILE m.serial = protab AND !EOF()
   IF mouhea.supcli = 1
      m.qtyent = m.qtyent+qty
   ELSE
      m.qtyext = m.qtyext+IIF(!EMPTY(mourow.proret) OR mouhea.return,-qty,qty+free)
   ENDIF   
   SKIP
ENDDO
m.delqty = m.qtyent-protab.adjqty