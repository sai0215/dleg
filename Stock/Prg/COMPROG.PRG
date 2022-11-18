PARAMETERS p1,p2,p3,p4,p5,p6

comchannel = p1
**WAIT WINDOW p4 NOWAIT
windout = WOUTPUT()
windsel = ALIAS()
windord = ORDER()
windrec = RECNO()
IF p2='INITIATE'
   v = '1'
   = DDEPOKE(p1,'comitem',v,'CF_TEXT')
ENDIF   
**LOCATE FOR VAL(protab.number) = VAL(SUBSTR(p4,1,LEN(ALLTRIM(STR(VAL(p4))))-1))
IF LEN(ALLTRIM(p4))>13
**	p4 = SUBSTR(p4,2,LEN(ALLTRIM(p4)))
	p4 = RIGHT(ALLTRIM(p4),13)
ELSE
	p4 = PADL(ALLTRIM(p4),13,'0')
ENDIF
IF p4='mycomtopic'
	WAIT WINDOW 'Com Routines loaded' NOWAIT
ELSE
	WAIT WINDOW p4 NOWAIT
ENDIF	
IF m.pubprog1 AND m.isediting
   SELECT protab
   LOCATE FOR protab.number=p4
   IF FOUND()
      SELECT moutmp
      LOCATE FOR protab.serial=protab AND EMPTY(moutmp.proret)
      m.comttfound = .T.
      IF !FOUND()
         LOCATE FOR moutype='N' 
         m.comttfound = .F.
      ENDIF   
      m.suprice = 0
      DO getuprice
      SELECT moutmp
      IF FOUND() 
**         REPLACE name     WITH protab.des
         REPLACE name     WITH protab.name
         REPLACE protab   WITH protab.serial
         REPLACE des      WITH protab.number
         REPLACE field2   WITH protab.field2
         REPLACE qty      WITH 1+qty
         REPLACE uprice   WITH m.suprice
         REPLACE price    WITH m.suprice*qty
         REPLACE value    WITH 0
         REPLACE disc     WITH 0
         REPLACE free     WITH 0
         m.ttqty = getqty(protab.serial)
         REPLACE stkqty   WITH m.ttqty-(qty+free)
         IF !m.comttfound
            REPLACE moutype WITH 'A'
            APPEND BLANK
            REPLACE mouhea WITH m.serial
            REPLACE moutype WITH 'N'
         ENDIF   
         DO caldisc
         SELECT moutmp
         GO BOTTOM
         m.propop =' '+SUBSTR(moutmp.des,1,10)+' ' + ;
SUBSTR(moutmp.name,1,15) +' '+ STR(moutmp.free,5)+' '+ STR(moutmp.qty,5) ;
+' '+ STR(moutmp.uprice,10)+' '+STR(moutmp.disc,3)+' '+ ;
STR(moutmp.price,10)
         SELECT moutmp
         _CUROBJ = OBJNUM(m.propop)
         SHOW GETS
      ENDIF   
   ENDIF
ENDIF
IF !EMPTY(windout)
    ACTIVATE WINDOW &windout SAME
ENDIF
IF NOT EMPTY(windsel)
	SELECT (windsel)
ENDIF
IF NOT EMPTY(windord)
	SET ORDER TO (windord)
ENDIF
IF windrec<>0 AND windrec<=RECCOUNT()
	GO RECORD windrec
ENDIF		
RETURN
