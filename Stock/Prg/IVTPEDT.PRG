CLOS DATA
report1  ='ivtredt'
m.nbpro  = 0
m.totqty = 0
DO ivtsedt.spr
CLOS DATA

PROCEDURE vblank
****************
SELECT (alias)
SCATTER MEMVAR BLANK
m.nbpro  = 0
m.totqty = 0
IF m.isediting
	m.date = DATE()
ENDIF
= checkzapfile('ivttmp')
APPEND BLANK
REPLACE ivthea  WITH m.serial
REPLACE ivttype WITH 'N'
SHOW GETS
SELECT (alias)

PROCEDURE getdata
*****************
SELECT (alias)
SCATTER MEMVAR
= checkzapfile("ivttmp")
SELECT ivtrow
SET ORDER TO ivthea
SEEK ivthea.serial
DO WHILE ivthea=ivthea.serial AND !EOF()
   SCATTER MEMVAR
   SELECT ivttmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE prodes   WITH protab.name
   REPLACE ivttype  WITH 'U'
   REPLACE ivtrecno WITH RECNO('ivtrow')
   SELECT ivtrow
   SKIP
ENDDO
SELECT ivttmp
COUNT TO m.nbpro
SUM qty TO m.totqty
APPEND BLANK
REPLACE ivthea  WITH m.serial
REPLACE ivttype WITH 'N'
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
GATHER MEMVAR
isadding = .F.
SET DELETED OFF
tserial = m.serial
SELECT ivttmp
GO TOP
DO WHILE !EOF()
   DO CASE
      CASE ivttype='A'
           SCATTER MEMVAR
           SELECT ivtrow
           APPEND BLANK
           GATHER MEMVAR
           REPLACE ivthea WITH tserial
      CASE ivttype='U'
           IF ivtrecno<>0
              SCATTER MEMVAR
              SELECT ivtrow
              GO RECORD m.ivtrecno
              GATHER MEMVAR
           ENDIF
      CASE ivttype='D'
           IF ivtrecno<>0
              SCATTER MEMVAR
              SELECT ivtrow
              GO RECORD m.ivtrecno
              DELETE
           ENDIF
   ENDCASE
   SELECT ivttmp
   REPLACE ivttype WITH IIF(ivttype='A','U',ivttype)
   SKIP
ENDDO
SET DELETED ON
SELECT (alias)

PROCEDURE search
*****************
SELECT (alias)
m.sfound = .F.
m.srecno = 0
IF issearch
   m.srecno = RECNO()
ENDIF
DO search.spr
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
SELECT ivtrow
DELETE FOR ivthea=m.serial
SELECT (alias)
RETURN .T.

PROCEDURE prtdata
*****************
m.qtyini  = 0
m.qtyexit = 0
m.tot1    = 0
SELECT ivttmp
GO TOP
DO WHILE !EOF()
	REPLACE prodes  WITH protab.name
	REPLACE selling WITH protab.selling
	REPLACE totsell WITH selling*qty
	REPLACE size    WITH fl4tab.name
	REPLACE cost    WITH getcostpr(protab)
	REPLACE totcost WITH cost*qty
	REPLACE famtab1 WITH protab.famtab1
	SELECT ivttmp
	SKIP
ENDDO
m.out   = 1
m.okbtn = 1
SUM totsell TO totsellm FOR famtab1=1
SUM totcost TO totcostm FOR famtab1=1
SUM totsell TO totselld FOR famtab1=2
SUM totcost TO totcostd FOR famtab1=2
DO printrec.spr
IF m.okbtn=2
	CLOSE DATA
	RETURN
ENDIF
SELECT ivttmp
SET ORDER TO famtab1
DELETE FOR ivttype='N'
IF m.out=1
	REPORT FORM ivtredt PREVIEW
ELSE
	REPORT FORM ivtredt TO PRINT OFF
ENDIF
SELECT ivttmp
SET ORDER TO
RECALL FOR ivttype='N'
SELECT (alias)

PROCEDURE search
*****************
SELECT (alias)
m.sfound = .F.
m.srecno = 0
IF issearch
   m.srecno = RECNO()
ENDIF
filesearch = "ivthea"
DO mousearch.spr
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

PROCEDURE prtrec
***************
SELECT prttab
DELETE FOR EMPTY(bidon)
SELECT (alias)
SET FILTER TO FOUND("prttab")
GO TOP
IF !EOF()
	DO newprint WITH report1,'FOUND("prttab")',IIF(m.print=1,'PREVIEW','TO PRINT')
ELSE
    WAIT WINDOW "EMPTY FILE" NOWAIT        
ENDIF