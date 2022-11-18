PARAMETERS pmousupcli

CLOS DATA
report ='mouredt'
IF pmousupcli=2
   m.pubprog1 = .T.
ENDIF   
m.card = .F.
DO mousedt.spr
m.pubprog1 = .F.
CLOS DATA

PROCEDURE vblank
****************
SHOW GET m.prvw_btn DISABLE
= checkzapfile("sertmp1")
= checkzapfile("moutmp")
m.ttsaltab = m.saltab
m.ttsaldes = m.saldes
m.ttwarhea = m.warhea
m.ttwardes = m.wardes
SELECT (alias)
m.prtcredit = .F.
m.mycurtab  = 1
m.curtab    = 1
m.wardes    = ''
m.rmodes    = ''
m.supdes    = ''
m.retdes    = '' 
m.saldes    = ''
m.bandes    = ''
m.curdes    = ''
m.bandes    = ''
m.lcurdes   = ''
m.totprice  = 0
m.totpayed  = 0
m.tpbalance = 0
m.tptotal   = 0
m.totpayed1 = 0
m.totpayed2 = 0
m.totret    = 0
m.card      = .F.
SCATTER MEMVAR BLANK
m.supcli    = pmousupcli
m.supclides = IIF(m.supcli=1,'Supplier','Customer')
IF isadding
   torder = ORDER()
   trecno = RECNO()
   SET FILTER TO
   SET ORDER TO 0
   GO BOTTOM
   m.serial = serial+1
   SET ORDER TO &torder
   SET FILTER TO supcli=pmousupcli
   IF trecno>0 AND trecno<=RECCOUNT()
      GO RECORD trecno
   ENDIF
   IF isadding
      m.saltab = m.ttsaltab
      m.saldes = m.ttsaldes
      m.warhea = m.ttwarhea
      m.wardes = m.ttwardes
      m.date   = DATE()
      SELECT curtab
      m.curtab = 1
      m.curdes = "DHS"
   ENDIF
ENDIF
SELECT moutmp
APPEND BLANK
REPLACE mouhea   WITH m.serial
REPLACE moutype  WITH 'N'
SHOW GETS
SELECT (alias)

PROCEDURE getdata
*****************
SELECT (alias)
IF !EOF()
   GO RECORD RECNO()
ENDIF   
SHOW GET m.prvw_btn ENABLE
m.wardes    = warhea.name
m.rmodes    = rmotab.name
m.bandes    = bantab.name
m.curdes    = curtab.number
m.lcurdes   = curtab.number
m.supdes    = suptab.name
m.saldes    = saltab.name
m.prtcredit = .F.
SELECT (alias)
SCATTER MEMVAR
m.mycurtab = m.curtab
IF !EMPTY(mouhea.supret)
	SELECT clitab
	torder = ORDER()
	SET ORDER TO serial
	SEEK supret
	m.supret = serial
	m.retdes = name
	SET ORDER TO &torder
	SELECT (alias)
ELSE
	m.retdes = ''
	m.supret = suptab	
ENDIF	
m.supcli    = pmousupcli
m.supclides = IIF(m.supcli=1,'Supplier','Customer')
m.supdes    = IIF(m.supcli=1,suptab.name,clitab.name)
m.card      = IIF(m.supcli=1,.F.,clitab.card)
= checkzapfile("sertmp1")
= checkzapfile("moutmp")
SELECT mourow
SEEK mouhea.serial
DO WHILE mouhea = mouhea.serial AND !EOF()
   SCATTER MEMVAR
   SELECT moutmp
   APPEND BLANK
   GATHER MEMVAR
   REPLACE mouhea   WITH m.serial
   REPLACE name     WITH protab.des
   REPLACE des      WITH protab.number
   REPLACE fl4des   WITH fl4tab.name
   m.sprotab = protab
   REPLACE stkqty   WITH getqty(protab)
   REPLACE moutype  WITH 'U'
   REPLACE mourecno WITH RECNO('mourow')
   SELECT srmtab
   SEEK RECNO('mourow')
   DO WHILE serial=RECNO('mourow') AND !EOF()
      SELECT sertmp1
      APPEND BLANK
      REPLACE serial  WITH RECNO('moutmp')
      REPLACE ser_nbr WITH srmtab.ser_nbr
      SELECT srmtab
      SKIP
   ENDDO
   SELECT mourow
   SKIP
ENDDO
SELECT moutmp
APPEND BLANK
REPLACE mouhea  WITH m.serial
REPLACE moutype WITH 'N'
m.curtab = m.mycurtab
DO caldisc
DO getbl
isadding = .F.
SELECT moutmp
SHOW GETS
SELECT (alias)

PROCEDURE getbl
***************
SELECT mouhea
mrecno = RECNO()
morder = ORDER()
SELECT mourow
wrecno = RECNO()
worder = ORDER()
SELECT pmotab
precno = RECNO()
porder = ORDER()
DO getclient WITH m.suptab,2
SELECT mouhea
IF mrecno>0 AND mrecno<=RECCOUNT()
    GO RECORD mrecno
ENDIF
SET ORDER TO &morder
SELECT mourow
IF wrecno>0 AND wrecno<=RECCOUNT()
    GO RECORD wrecno
ENDIF
SET ORDER TO &worder
SELECT pmotab
IF precno>0 AND precno<=RECCOUNT()
    GO RECORD precno
ENDIF
SET ORDER TO &porder
SHOW GETS
SELECT (alias)

PROCEDURE savedata
******************
SHOW GET m.prvw_btn ENABLE
SELECT (alias)
m.supcli = pmousupcli
IF isadding
   APPEND BLANK
ENDIF
SELECT mourow
SEEK mouhea.serial
DO WHILE mouhea=mouhea.serial AND !EOF()
   SELECT srmtab
   DELETE FOR serial=RECNO('mourow')
   SELECT mourow
   SKIP
ENDDO
m.mycurtab = m.curtab
m.mydate   = m.date
SET DELETED OFF
SELECT moutmp
COUNT TO recnn
GO TOP
IF recnn=1 AND moutype='N'
	SELECT mourow
	APPEND BLANK
    GATHER MEMVAR
    REPLACE mouhea WITH m.serial
ELSE	
DO WHILE !EOF()
   DO CASE
      CASE moutype='A' 
           SCATTER MEMVAR
           SELECT mourow
           APPEND BLANK
           GATHER MEMVAR
           REPLACE mouhea 		   WITH m.serial
           REPLACE qtyfin 		   WITH qty
           REPLACE time            WITH TIME()
           REPLACE moutmp.mourecno WITH RECNO('mourow')
           SELECT sertmp1
           SEEK RECNO('moutmp')
           DO WHILE serial=RECNO('moutmp') AND !EOF()
              SELECT srmtab
              APPEND BLANK
              REPLACE serial  WITH RECNO('mourow')
              REPLACE ser_nbr WITH sertmp1.ser_nbr
              SELECT sertmp1
              SKIP
           ENDDO
           SELECT moutmp
      CASE moutype='R'
      	 SCATTER MEMVAR
         IF mourecno<>0
            SELECT mourow
            GO RECORD m.mourecno
            DELETE
         ENDIF
           SELECT mourow
           APPEND BLANK
           GATHER MEMVAR
           REPLACE mouhea WITH m.serial
           REPLACE qtyfin WITH qty
           REPLACE time   WITH TIME()
           SELECT sertmp1
           SEEK RECNO('moutmp')
           DO WHILE serial=RECNO('moutmp') AND !EOF()
              SELECT srmtab
              APPEND BLANK
              REPLACE serial  WITH RECNO('mourow')
              REPLACE ser_nbr WITH sertmp1.ser_nbr
              SELECT sertmp1
              SKIP
           ENDDO
           SELECT moutmp          
      CASE moutype = 'U'
           IF mourecno <> 0
              SCATTER MEMVAR
              SELECT mourow
              GO RECORD m.mourecno
              GATHER MEMVAR
              REPLACE mouhea WITH m.serial
              REPLACE time   WITH TIME()
              REPLACE qtyfin WITH qty
              SELECT sertmp1
              SEEK RECNO('moutmp')
              DO WHILE serial=RECNO('moutmp') AND !EOF()
                 SELECT srmtab
                 APPEND BLANK
                 REPLACE serial  WITH RECNO('mourow')
                 REPLACE ser_nbr WITH sertmp1.ser_nbr
                 SELECT sertmp1
                 SKIP
              ENDDO
              SELECT moutmp
           ENDIF
      CASE moutype='D'
           IF mourecno <> 0
              SCATTER MEMVAR
              SELECT mourow
              GO RECORD m.mourecno
              DELETE
           ENDIF
   ENDCASE
   SELECT moutmp
   REPLACE moutype WITH IIF(moutype='A','U',moutype)
   SKIP
ENDDO
ENDIF
SET DELETED ON
SELECT (alias)
m.supcli = pmousupcli
m.curtab = m.mycurtab
m.date   = m.mydate
GATHER MEMVAR
IF !m.return
	REPLACE supret WITH 0
ENDIF
REPLACE time WITH TIME()
DO getbl
isadding = .F.

PROCEDURE search
*****************
SELECT (alias)
m.sfound = .F.
m.srecno = 0
IF issearch
   m.srecno = RECNO()
ENDIF
IF pmousupcli=1
   filesearch = "suptab"
ELSE
   filesearch = "clitab"
ENDIF   
DO mousearch.spr
SELECT (alias)
IF LASTKEY()<>27
	SET FILTER TO supcli=pmousupcli
	IF m.srecno<>0 AND !m.sfound 
		GO TOP
		IF NOT EOF() 
			GO RECORD m.srecno
		ENDIF   
	ENDIF
	IF m.sfound
		m.issearch = .T.
		DO getdata
	ELSE
		SELECT moutmp
		SHOW GETS
		SELECT (alias)
	ENDIF
ENDIF

PROCEDURE okdel
***************
SELECT pmotab
DELETE FOR mouhea=m.serial
SELECT mourow
DELETE FOR mouhea=m.serial
SELECT (alias)
RETURN .T.

PROCEDURE prtrec
****************
SELECT (alias)
DO print WITH report,IIF(m.criteria=1,'RECNO()=m.saverec',IIF(m.criteria=2,'NOT EMPTY(bidon)','.T.')),IIF(m.print=1,'PREVIEW','TO PRINT')

PROCEDURE caldisc
*****************
SELECT moutmp
m.purtotprice = 0
m.totprice    = 0
m.rettotprice = 0
m.ttmourec    = RECNO()
GO TOP
DO WHILE !EOF()
	IF EMPTY(proret)
		m.purtotprice = m.purtotprice+price 
	ELSE
		m.discc       = disc
		m.rettotprice = m.rettotprice+price
	ENDIF
	SKIP
ENDDO
m.totprice =  m.purtotprice-m.rettotprice
m.totprice = (m.totprice*(100-m.discount))/100
m.totprice = INT(m.totprice+0.5*IIF(m.totprice>0,1,-1))
**m.totprice = ROUND(m.totprice,0)
IF m.curtab=1
	m.totprice = ROUND(m.totprice,0)
ENDIF	
DO calpayed
SELECT moutmp
IF m.ttmourec>0 AND m.ttmourec<=RECCOUNT()
   GO RECORD m.ttmourec
ENDIF

RETURN
IF !EMPTY(m.rettotprice)
	IF m.rettotprice<m.purtotprice
		m.xxrr  = ((m.rettotprice)*(100-m.discc))/100
		m.ssprr = ((m.purtotprice-m.rettotprice)*(100-m.discount))/100
		WAIT WIND STR(m.xxrr)
		WAIT WIND STR(m.ssprr)
		m.totprice = m.xxrr+m.ssprr
	ELSE
		m.xxrr     = ((m.purtotprice)*(100-m.discc))/100
		m.totprice = m.xxrr-m.rettotprice
	ENDIF	
ELSE
	m.totprice = (m.totprice*(100-m.discount))/100
ENDIF

PROCEDURE calssdisc
*******************
SELECT moutmp
m.ssprice  = 0
m.ttmourec = RECNO()
GO TOP
DO WHILE !EOF()
	IF EMPTY(proret) 
		m.ssprice = m.ssprice+price
	ELSE
		m.ssprice = m.ssprice-price
	ENDIF
	SKIP
ENDDO			
SELECT moutmp
GO TOP
DO WHILE !EOF()
	IF !EMPTY(free) AND EMPTY(proret)
		m.ssprice = m.ssprice+uprice*free
	ENDIF
	SKIP
ENDDO		
SELECT moutmp
IF m.ttmourec>0 AND m.ttmourec<=RECCOUNT()
   GO RECORD m.ttmourec
ENDIF   
RETURN m.ssprice

PROCEDURE calbalance
********************
return
DO gettotcli 
m.tpbalance  = m.tpbalance + m.totprice - mouhea.totprice
m.tpbalance  = m.tpbalance - (m.totpayed-mouhea.totpayed)
*m.tpbalance = cal_cur(m.tpbalance,date,curtab,m.curtab)

PROCEDURE addcli
****************
SELECT clitab
torder = ORDER()
IF m.suptab=0
   SET ORDER TO serial
   GO BOTTOM
   m.tserial = serial+1 
   m.number  = ROUND(VAL(number),0)+1
   m.number  = PADL(m.number,5,'0')
   SET ORDER TO &torder
   APPEND BLANK
   GATHER MEMVAR
   REPLACE serial WITH m.tserial
ELSE
   SET ORDER TO serial
   SEEK m.suptab
   IF FOUND()
*   	GATHER MEMVAR
		REPLACE name    WITH m.name
		REPLACE address WITH m.address
		REPLACE tel1    WITH m.tel1
		REPLACE tel2    WITH m.tel2		
*   	REPLACE serial  WITH m.tserial
		m.supdes = name
   ENDIF
   SET ORDER TO &torder
ENDIF   
m.tttsave = .T.

PROCEDURE addban
****************
SELECT bantab
torder = ORDER()
SET ORDER TO 0
GO BOTTOM
m.tserial = serial+1
SET ORDER TO &torder
APPEND BLANK
**GATHER MEMVAR
REPLACE serial  WITH m.tserial
REPLACE name    WITH m.name
REPLACE number  WITH m.number
REPLACE acc_no  WITH m.acc_no
REPLACE balance WITH m.balance
REPLACE curtab  WITH m.curtab
m.tttsave = .T.

PROCEDURE updprice
******************
SELECT moutmp
GO TOP
DO WHILE !EOF()
   IF moutype<>'N'
      SELECT protab
      SEEK moutmp.protab
      m.suprice = 0
      DO getuprice
      SELECT moutmp
      REPLACE uprice WITH m.suprice
      m.price = uprice*qty
      m.price = (m.price*(100-disc))/100
      REPLACE price WITH m.price
      REPLACE value WITH uprice-m.price
   ENDIF
   SKIP
ENDDO
DO caldisc
DO getbl
SELECT (alias)

PROCEDURE prtdata
*****************
m.out = 2
SELECT (alias)
wrec    = RECNO()
**m.okbtn = 1
**DO printrec.spr
**IF m.okbtn=2
**	RETURN
**ENDIF	
SELECT moutmp
CALCULATE MAX(Moutmp.disc) to m.sdiscount
IF m.out=1
	REPORT FORM dsaredt1 FOR moutype<>'N' PREVIEW
ELSE
	REPORT FORM dsaredt1 FOR moutype<>'N' NoConsole NoEject TO PRInt
***  EJECT
ENDIF	
SELECT (alias)
RETURN	

PROCEDURE prtdata1
******************
m.out = 1
SELECT (alias)
wrec = RECNO()
IF pmousupcli = 1 
	m.return = .F.
ELSE
	m.return = .T.
ENDIF	
DO printrec.spr
IF m.out=1
	REPORT FORM dsaredt FOR RECNO()=wrec PREVIEW 
	REPORT FORM dsaredt FOR RECNO()=wrec PREVIEW
ELSE
	REPORT FORM dsaredt1 FOR RECNO()=wrec TO PRINT
	REPORT FORM dsaredt FOR RECNO()=wrec NOEJECT TO PRINT
ENDIF
RETURN	

PROCEDURE locname
*****************
SELECT protab
SEEK moutmp.protab
SELECT moutmp
RETURN protab.number

PROCEDURE getold
****************
= checkzapfile("moutmp1")
m.sercli = clitab.serial
SELECT mouhea
torder = ORDER()
trecno = RECNO()
SET ORDER TO 0
GO BOTTOM
m.newhea = serial
SET ORDER TO suptab
SEEK m.sercli
DO WHILE FOUND() AND mouhea.suptab=m.sercli
	m.serhea = mouhea.serial
	SELECT mourow
	morder = ORDER()
	mrecno = RECNO()
	SEEK m.serhea
	DO WHILE FOUND() AND m.serhea=mourow.mouhea
		m.name   = protab.name
		m.protab = protab.serial
		m.des    = des
		m.qty    = qty
		m.date   = mouhea.date
		m.mouhea = m.serhea
		SELECT moutmp1
		APPEND BLANK
		GATHER MEMVAR
		SELECT mourow
		SKIP
	ENDDO
	SELECT mouhea
	SKIP
ENDDO
SELECT mouhea
SET ORDER TO &torder
IF trecno>0 AND trecno<=RECCOUNT()
	GO RECORD trecno
ENDIF
SELECT mourow
SET ORDER TO &morder
IF mrecno>0 AND mrecno<=RECCOUNT()
	GO RECORD mrecno
ENDIF
	 
PROCEDURE calsum1
*****************
SELECT moutmp
m.calsum1 = 0
GO TOP
DO WHILE !EOF()
	IF EMPTY(proret)
		IF moutype<>'N'
    		SELECT protab
	    	SEEK moutmp.protab
    		SELECT moutmp
	    	m.calsum1 = m.calsum1+(protab.selling*qty)*((100-disc)/100)
		ENDIF
	ELSE
    	m.calsum1 = m.calsum1-cal_cur(moutmp.price,mouhea.date,mouhea.curtab,1)
	ENDIF
	SKIP
ENDDO
m.calsum1 = INT(m.calsum1*((100-m.discount)/100))
SELECT moutmp
RETURN IIF(m.curtab<>1,m.calsum1,0)

PROCEDURE totppp
****************
SELECT moutmp
m.totppp = 0
GO TOP
DO WHILE !EOF()
	IF EMPTY(proret)
		m.totppp = m.totppp+price 
	ELSE
		m.totppp = m.totppp-price
	ENDIF	
	SKIP
ENDDO
m.totppp = (m.totppp*(100-m.discount))/100
**RETURN INT(m.totppp+.5)
RETURN ROUND(m.totppp,0)

PROCEDURE getqty
****************
PARAMETERS ttprotab

m.qtyent = 0
m.qtyext = 0
tselect  = SELECT()
SELECT protab
SET ORDER TO serial
SEEK ttprotab
m.adjqty = protab.adjqty
SELECT arrrow
trecno = RECNO()
torder = ORDER()
SET ORDER TO protab
SEEK ttprotab
DO WHILE protab=ttprotab AND !EOF()
   m.qtyent = m.qtyent+qty
   SKIP
ENDDO
IF trecno>0 AND trecno<=RECCOUNT()
	GO RECORD trecno
ENDIF
SET ORDER TO &torder
SELECT mouhea
torder1 = ORDER()
trecno1 = RECNO() 
SET ORDER TO serial
SELECT mourow
trecno2 = RECNO()
torder2 = ORDER()
SET RELATION TO mourow.mouhea INTO mouhea ADDITIVE
SET ORDER TO protab
SEEK ttprotab
DO WHILE protab=ttprotab AND !EOF()
   IF mouhea.supcli=1
      m.qtyent = m.qtyent+qty
   ELSE
*     m.qtyext = m.qtyext+IIF(mouhea.return,-qty,qty)   
	  IF !EMPTY(mourow.proret) 
	      m.qtyext = m.qtyext-qty
	  ELSE
	      m.qtyext = m.qtyext+(qty+free)
	  ENDIF    
   ENDIF   
   SKIP
ENDDO
SET RELATION OFF INTO mouhea
IF trecno2>0 AND trecno2<=RECCOUNT()
	GO RECORD trecno2
ENDIF
SET ORDER TO &torder2
SELECT mouhea
IF trecno1>0 AND trecno1<=RECCOUNT()
	GO RECORD trecno1
ENDIF
SET ORDER TO &torder1
SELECT (tselect)
RETURN (m.qtyent-m.qtyext)+m.adjqty

PROCEDURE somme
***************
tselect = SELECT()
SELECT moutmp
SUM qty  TO n FOR EMPTY(proret)
SUM free TO m
SELECT (tselect)
RETURN n+m

PROCEDURE viewdata
******************
m.out = 1
SELECT (alias)
wrec    = RECNO()
**m.okbtn = 1
**DO printrec.spr
**IF m.okbtn=2
**	RETURN
**ENDIF	
SELECT moutmp
CALCULATE MAX(Moutmp.disc) to m.sdiscount
IF m.out=1
	REPORT FORM dsaredt1 FOR moutype<>'N' PREVIEW
ELSE
	REPORT FORM dsaredt1 FOR moutype<>'N' NoConsole NoEject TO PRInt
***  EJECT
ENDIF	
SELECT (alias)
RETURN	

PROCEDURE prtdata
*****************
SELECT (alias)
m.sldqty = somme()
SELECT pmotmp
GO BOTTOM
IF !BOF()
	SKIP -1
ENDIF
DO CASE	
CASE pmotmp.type=1
	paytype  = 'Cash'
	m.ampaid = ampaid
	m.amret  = amret
CASE pmotmp.type=2
	paytype = 'Check'
	m.ampaid = amount
	m.amret  = 0	
CASE pmotmp.type=3
	paytype = 'American Express'
	m.ampaid = amount
	m.amret  = 0
CASE pmotmp.type=4
	paytype = 'Master Card'
	m.ampaid = amount
	m.amret  = 0
CASE pmotmp.type=5
	paytype = 'Visa'
	m.ampaid = amount
	m.amret  = 0		
ENDCASE
m.output = "P"
m.output = m.output+"          Du Pareil Au Meme"+CHR(13)
m.output = m.output+"          Du Pareil Au Meme"+CHR(13)
m.output = m.output+"          Du Pareil Au Meme"+CHR(10)
m.output = m.output+"         Wafi Center - Dubai"+CHR(13)
m.output = m.output+"         Wafi Center - Dubai"+CHR(13)
m.output = m.output+"         Wafi Center - Dubai"+CHR(10)
m.output = m.output+"          Tel: 04-351 51 15"+CHR(13)
m.output = m.output+"          Tel: 04-324 51 15"+CHR(10)
m.output = m.output+CHR(10)
m.output = m.output+"          Invoice no: M"+PADL(mouhea.serial,5,"0")+CHR(13)
m.output = m.output+"          Invoice no: M"+PADL(mouhea.serial,5,"0")+CHR(10)
m.output = m.output+CHR(10)+CHR(10)
DO sendcom WITH m.output
m.output = "P"
SELECT moutmp
COUNT TO crec
GO TOP
pcnt = 0
DO WHILE !EOF()
	IF !EMPTY(name)
		pcnt = pcnt+1
		m.output = m.output+"  "+SUBSTR(name,1,25)+" "+PADL(ALLTRIM(field2),10,' ')+CHR(10)
		IF disc=0
			m.output = m.output+IIF(!EMPTY(proret),'  Return     ',SPACE(13))+STR(IIF(qty=0,free,qty),2)+" x "+STR(ROUND(uprice,0),6,2)+SPACE(6)+STR(ROUND(IIF(!EMPTY(proret),-1,1)*price,0),8,2)+CHR(10)
		ELSE
			m.output = m.output+'  Disc: '+STR(disc,2)+'%  '+STR(IIF(qty=0,free,qty),2)+" x "+STR(ROUND(uprice,0),6,2)+SPACE(6)+STR(ROUND(price,0),8,2)+CHR(10)
		ENDIF
		IF pcnt=10
			DO wait WITH 12
			DO sendcom WITH m.output
			m.output = "P"
			pcnt = 0
		ENDIF
	ENDIF	
	SKIP
ENDDO
IF pcnt>0
	DO wait WITH 12
	DO sendcom WITH m.output
ENDIF	
m.output = "P"
m.output = m.output+CHR(10)
IF discount<>0
	m.output = m.output+SPACE(13)+"Discount"+SPACE(14)+STR(discount,2)+"%"+CHR(10)
ENDIF
m.output = m.output+SPACE(13)+"T O T A L"+SPACE(7)+STR(m.totprice,9,2)+CHR(13)
m.output = m.output+SPACE(13)+"T O T A L"+SPACE(7)+STR(m.totprice,9,2)+CHR(13)
m.output = m.output+SPACE(13)+"T O T A L"+SPACE(7)+STR(m.totprice,9,2)+CHR(10)
m.output = m.output+"  Paid By    "+paytype+SPACE(16-LEN(paytype))+STR(m.ampaid,9,2)+CHR(10)
m.output = m.output+SPACE(13)+"Change"+SPACE(10)+STR(m.amret,9,2)+CHR(10)
m.output = m.output+CHR(10)
m.output = m.output+"  Total Number of Items sold = "+ALLTRIM(STR(m.sldqty))+CHR(10)
m.output = m.output+"  "+DTOC(DATE())+" "+SUBSTR(TIME(),1,5)+CHR(10)
m.output = m.output+"    ** Keep Receipt for Exchange **"+CHR(10)
m.output = m.output+"       Thank you for Shopping at"+CHR(13)
m.output = m.output+"       Thank you for Shopping at"+CHR(10)
m.output = m.output+"           Du Pareil Au Meme"+CHR(13)
m.output = m.output+"           Du Pareil Au Meme"+CHR(10)
m.output = m.output+CHR(10)+CHR(10)+CHR(10)+CHR(10)+CHR(10)+CHR(10)+CHR(10)+CHR(10)+CHR(10)
m.output = m.output+CHR(27)+"i"
DO wait WITH crec*10
DO sendcom WITH m.output
SELECT (alias)
RETURN