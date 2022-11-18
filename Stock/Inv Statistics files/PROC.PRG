Procedure btn_val
*****************
PARAMETER m.tbtnname
DO CASE
   CASE m.tbtnname='TOP'
        m.issearch = .T.
		GO TOP
		WAIT WINDOW TOPFILE NOWAIT
		DO getdata
   CASE m.tbtnname='PREV'
        m.issearch = .T.
	 	IF !BOF()
		   SKIP -1
  	 	ENDIF
	 	IF BOF()
		   WAIT WINDOW TOPFILE NOWAIT
		   GO TOP
		ENDIF
		DO getdata
   CASE m.tbtnname='NEXT'
        m.issearch = .T.
 	    IF !EOF()
		   SKIP 1
		ENDIF
		IF EOF()
		   WAIT WINDOW ENDFILE NOWAIT
		   GO BOTTOM
		ENDIF
		DO getdata
   CASE m.tbtnname='END'
        m.issearch = .T.
		GO BOTTOM
		WAIT WINDOW ENDFILE NOWAIT
		DO getdata
   CASE m.tbtnname='LOCATE'
        DO search
   CASE m.tbtnname='ADD'
 		isediting= .T.
		issearch = .T.
		isadding = .T.
		DO vblank
		_curobj=1
   CASE m.tbtnname='EDIT'
		IF EOF() OR BOF()
		   WAIT WINDOW ENDFILE NOWAIT
		   RETURN
		ENDIF
 		isediting= .T.
		_curobj=1
   CASE m.tbtnname='SAVE'
		isediting=.F.
		DO savedata
   CASE m.tbtnname='CANCEL'
		isediting=.F.
   		DO getdata
   CASE m.tbtnname='DELETE' 
		IF EOF() OR BOF()
			WAIT WINDOW ENDFILE NOWAIT
			RETURN
		ENDIF
		IF fox_alert(DELREC) AND okdel()
			DELETE
			IF !EOF()
				SKIP 1
			ENDIF
			IF EOF()
				WAIT WINDOW ENDFILE NOWAIT
				GO BOTTOM
			ENDIF
		ENDIF
		DO getdata
   CASE m.tbtnname='PRINT'
		IF EOF() OR BOF()
		   WAIT WINDOW ENDFILE NOWAIT
		   RETURN
		ENDIF
		DO prtdata
		RETURN
   CASE m.tbtnname='VIEW'
		IF EOF() OR BOF()
		   WAIT WINDOW ENDFILE NOWAIT
		   RETURN
		ENDIF
		DO viewdata
		RETURN		
   CASE m.tbtnname='EXIT'
		CLEAR READ
		RETURN
ENDCASE
DO refresh
RETURN

PROCEDURE refresh
*****************
DO CASE
   CASE m.isediting
        SHOW GET find_drop DISABLE
  		SHOW GET top_btn DISABLE
		SHOW GET prev_btn DISABLE
		SHOW GET loc_btn DISABLE
		SHOW GET next_btn DISABLE
		SHOW GET end_btn DISABLE
		SHOW GET add_btn DISABLE
		SHOW GET edit_btn,1 DISABLE
		SHOW GET del_btn,1 DISABLE
		SHOW GET prnt_btn DISABLE
		SHOW GET exit_btn DISABLE
		SHOW GET save_btn ENABLE
		SHOW GET can_btn ENABLE
		RETURN
   CASE !isediting
	    SHOW GET find_drop ENABLE
		SHOW GET top_btn ENABLE
		SHOW GET prev_btn ENABLE
		SHOW GET loc_btn ENABLE
		SHOW GET next_btn ENABLE
		SHOW GET end_btn ENABLE
		SHOW GET add_btn ENABLE
		IF issearch
	 	   SHOW GET edit_btn,1 ENABLE
		   SHOW GET del_btn,1 ENABLE
  		   SHOW GET prnt_btn ENABLE
		ENDIF   
		SHOW GET exit_btn ENABLE
		SHOW GET save_btn DISABLE
		SHOW GET can_btn DISABLE
ENDCASE
RETURN

PROCEDURE fox_alert
*******************
PARAMETER tstring
m.tmessage = tstring
m.drec = .F.
m.ok = 2
DO message.spr
RETURN m.drec

PROCEDURE pro_pop
*****************
IF LASTKEY()=13
   DEACTIVATE POPUP 
ENDIF

PROCEDURE pro_popx
******************
IF LASTKEY()=13
   REPLACE bidon WITH IIF(EMPTY(bidon),'>','')
ENDIF

PROCEDURE print
***************
PARAMETERS tname,tcondition,ttype

REPORT FORMAT &tname FOR &tcondition &ttype

PROCEDURE getminutes
********************
PARAMETERS tmin,tsign,tlen

ttmin = ALLTRIM(tmin)
ttmin = IIF(tsign=1,VAL(SUBSTR(ttmin,2,tlen))*60+VAL(SUBSTR(ttmin,tlen+3,tlen+5)),VAL(SUBSTR(ttmin,1,tlen))*60+VAL(SUBSTR(ttmin,tlen+2,tlen+5)) )
ttmin = IIF(tsign=1 AND SUBSTR(tmin,1,1)='-',-ttmin,ttmin)
RETURN INT(ttmin)

PROCEDURE setminutes
********************
PARAMETERS tmin,tsign,tlen

ttsign = IIF(tmin<0,'-','+')
tmin   = ABS(tmin)
tthour = INT(tmin/60)
ttmin  = tmin%60
ttsign = IIF(tsign=1,ttsign,'')
RETURN ttsign+PADL(ALLTRIM(STR(tthour)),tlen,'0')+':'+PADL(ALLTRIM(STR(ttmin)),2,'0')

PROCEDURE spop
**************
PARAMETERS tfile,tfield,tx,ty,tsizex,tsizey,tvariable,tsearch,tfilter

IF TYPE('tsearch')='L'
   tseek = tfield
ELSE
   tseek = tsearch
ENDIF

SELECT (tfile)
DEFINE POPUP ttpop FROM tx,ty TO tx+tsizex,ty+tsizey PROMPT FIELD &tfield SCROLL
ON SELECTION POPUP ttpop DO pro_pop
LOCATE FOR ALLTRIM(&tseek)==ALLTRIM(&tvariable)
IF FOUND()
   STORE &tseek TO &tvariable
   SHOW GET &tvariable
   RETURN .T.
ENDIF   
*IF MDOWN()
*   RETURN .T.
*ENDIF
IF EMPTY(tfilter)
	SET FILTER TO ALLTRIM(&tseek)=ALLTRIM(&tvariable)
ELSE
	SET FILTER TO &tfilter AND ALLTRIM(&tseek)=ALLTRIM(&tvariable)
ENDIF	
GO TOP
* WAIT WINDOW IIF(MDOWN(),'MDOWN','NOT MDOWN')
IF EOF()
	WAIT WINDOW IIF(m.publang,'No matching record',"Enregistrement non trouve") NOWAIT
	IF EMPTY(tfilter)
		SET FILTER TO
	ELSE
		SET FILTER TO &tfilter
	ENDIF	
	GO TOP
	IF EOF()
    	RETURN .F.
	ENDIF   
ENDIF
ACTIVATE POPUP ttpop
IF EMPTY(tfilter)
	SET FILTER TO
ELSE
	SET FILTER TO &tfilter
ENDIF	
RELEASE POPUP ttpop
STORE &tseek TO &tvariable
SHOW GET &tvariable
RETURN .T.

PROCEDURE drwbox
****************
PARAMETERS tminx,tminy,tmaxx,tmaxy,tx,tcolor,tlcolor,tccolor,tecolor,ttype,tpattern

tred    = VAL(SUBSTR(tcolor,1,3))
tgreen  = VAL(SUBSTR(tcolor,4,3))
tblue   = VAL(SUBSTR(tcolor,7,3))
tlred   = VAL(SUBSTR(tlcolor,1,3))
tlgreen = VAL(SUBSTR(tlcolor,4,3))
tlblue  = VAL(SUBSTR(tlcolor,7,3))
tcred   = VAL(SUBSTR(tccolor,1,3))
tcgreen = VAL(SUBSTR(tccolor,4,3))
tcblue  = VAL(SUBSTR(tccolor,7,3))
tered   = VAL(SUBSTR(tecolor,1,3))
tegreen = VAL(SUBSTR(tecolor,4,3))
teblue  = VAL(SUBSTR(tecolor,7,3))
tlastx  = MROW()
tmrow   = tlastx
DO WHILE MDOWN()
   = INKEY(.06,'HM')
   trow = MROW( )
   tcol = MCOL( )
   IF BETWEEN(trow,IIF(ttype=2,tx,tminx),IIF(ttype=1,tx,tmaxx)) AND BETWEEN(tcol,tminy,tmaxy)
      IF ttype=1 AND trow<=tx
         tmrow = IIF(trow <tmrow,trow-.1,tmrow)
         @ trow,tminy TO tx,tmaxy PATTERN 1 COLOR RGB(tred,tgreen,tblue,tred,tgreen,tblue)
         @ tmrow,tminy TO trow,tmaxy PATTERN tpattern COLOR RGB(tlred,tlgreen,tlblue,tlred,tlgreen,tlblue)
         tlastx = MROW()
      ENDIF
      IF ttype=2 AND trow>=tx
         tmrow = IIF(trow>tmrow,trow+.1,tmrow)
         @ tx,tminy TO trow,tmaxy PATTERN 1 COLOR RGB(tred,tgreen,tblue,tred,tgreen,tblue)
         @ trow,tminy TO tmrow,tmaxy PATTERN tpattern COLOR RGB(tlred,tlgreen,tlblue,tlred,tlgreen,tlblue)
         tlastx = MROW()
      ENDIF
      IF (BETWEEN(trow,tx-0.24,tx) AND ttype=1) OR (BETWEEN(trow,tx,tx+0.24) AND ttype=2)
         @ tx+IIF(ttype=1,-.2,0),tminy TO tx+IIF(ttype=2,0.2,0),tmaxy PATTERN 1 COLOR RGB(tered,tegreen,teblue,tered,tegreen,teblue)
      ELSE   
         @ trow,tminy TO trow,tmaxy PATTERN 1 COLOR RGB(tcred,tcgreen,tcblue,tcred,tcgreen,tcblue)
      ENDIF   
   ENDIF
ENDDO
RETURN tlastx

PROCEDURE getcur
****************
PARAMETERS tpos,ttime

t_tx = m.tpos+((m.pubbox/24)*ttime)/60
RETURN t_tx

PROCEDURE fillbox
*****************
PARAMETERS ty,tfrom,tto,tlen,tcolor,tpattern

IF NOT EMPTY(tcolor)
   IF tfrom<=tto
      @ tfrom,ty CLEAR TO tto,ty+tlen
   ENDIF   
ENDIF
m.trcol = VAL(SUBSTR(tcolor,1,3))
m.tgcol = VAL(SUBSTR(tcolor,4,3))
m.tbcol = VAL(SUBSTR(tcolor,7,3))
IF tfrom<=tto
   @ tfrom,ty TO tto,ty+tlen PATTERN tpattern COLOR RGB(trcol,tgcol,tbcol,trcol,tgcol,tbcol)
ENDIF

PROCEDURE nottime
*****************
PARAMETERS tnottime

IF SUBSTR(tnottime,1,2)<'00' OR SUBSTR(tnottime,1,2)>'23'
   WAIT WINDOW IIF(m.publang,'Invalid time','Temps invalide') NOWAIT
   RETURN .T.
ENDIF
IF SUBSTR(tnottime,4,2) < '00' OR SUBSTR(tnottime,4,2) > '59' 
   WAIT WINDOW IIF(m.publang,'Invalid time','Temps invalide') NOWAIT
   RETURN .T.
ENDIF
RETURN .F.

PROCEDURE drwsch
****************
PARAMETERS txbegin,tybegin

m.tystep = m.pubbox/24
m.txlen  = 0.5
m.tylen  = m.pubbox
tm = 0
@ txbegin+0.1,tybegin+0.1 CLEAR TO txbegin+tylen-0.1,tybegin+3-0.1
@ txbegin,tybegin TO txbegin+tylen,tybegin+3 PEN 1, 8
ttybegin = tybegin+3
FOR tI=txbegin TO txbegin+tylen STEP tystep
    @ tI,ttybegin,tI,ttybegin+txlen BOX
    ttime = PADL(ALLTRIM(STR(tm)),2,'0')+':00'
    @ tI-0.5,ttybegin+txlen+0.5 SAY ttime FONT "MS Serif",6
    tm = tm+1
ENDFOR

PROCEDURE getzone
*****************
PARAMETERS tdate,temptab,ttype

IF ttype=0
ELSE
ENDIF
SELECT schupd
SEEK STR(temptab)+DTOC(tdate)
IF FOUND()
   SCATTER MEMVAR
   RETURN 1
ENDIF
SELECT schtra
SEEK temptab
IF !FOUND()
   SCATTER MEMVAR BLANK
   RETURN 0
ENDIF
DO WHILE pdate<=tdate AND pemptab=temptab AND !EOF()
*  SKIP (tdate-startdate)%(protation*7)
   Scatter MemVar
 * SKIP protation*7-(tdate-startdate)%(protation*7)
	SKIP
ENDDO
If m.pbegin>m.pend And m.ptype!=2 And m.ptype!=3
*	wait wind "Night Shift " Nowait
	m.plebegin 	 = SpecialCase(m.pbegin,m.plebegin)
	m.peebegin 	 = SpecialCase(m.pbegin,m.peebegin)
*	m.plbbegin 	 = SpecialCase(m.pbegin,m.plbbegin)
	m.peenbegin1 = SpecialCase(m.pbegin,m.peenbegin1)
	m.peenbegin2 = SpecialCase(m.pbegin,m.peenbegin2)
	m.plprbegin  = SpecialCase(m.pbegin,m.plprbegin)
	m.plexbegin1 = SpecialCase(m.pbegin,m.plexbegin1)
	m.plexbegin2 = SpecialCase(m.pbegin,m.plexbegin2)
	m.pubnight = .T.
EndIf
RETURN 1

Procedure GetPara
*===============*
Parameters tdate,tnbr

Dimension fac(30,2)

For nm = 1 TO 30 
	fac(nm,1) = 0
	fac(nm,2) = 0
EndFor 

ifac = 1
m.calwday     = 1
m.calnormal   = 0
m.calrequest  = 0
m.calrentry   = 0
m.calentry    = 0
m.calrexit    = 0
m.calexit     = 0
m.calrpres    = 0
m.calpres     = 0
m.calrabs     = 0
m.calabs      = 0
m.calalldelta = 0
m.talldelta   = 0
m.calpenadd   = 0
m.calallpen   = 0
m.calfactor   = 0
m.tcalfactor   = 0
m.calfacdelta = 0
m.calnormal   = m.ptotal
m.calrequest  = m.ptotal

If m.ptotal = 0
   m.calwday = 0
EndIf

If m.pbegin > m.pend And m.ptype != 2 And m.ptype != 3
	m.pend = m.pend + 1439 
EndIf

*^^^^^^^^^^^^^^ Variable Time ^^^^^^^^^^^^^^*
If m.ptype = 2
	Wait Wind "Variable Time " NoWait
	m.varsubst = 0
	For vi = 1 To 30
		If Empty(pubentry(vi)) And Empty(pubexit(vi))	 
			Exit 
		EndIf
		m.varsubst = m.varsubst + (pubexit(vi) - pubentry(vi))	
	EndFor
	Wait Wind "Substr     " +Str(m.varsubst)
	Wait Wind "Normal     " + Str(m.calnormal)
	Do case 
		Case m.varsubst > m.calnormal
					*** Late Exit ***
			m.varsubst = m.varsubst - m.calnormal
			m.calrexit = adjust(m.varsubst,4)
         	m.calexit  = getcalcul(pubentry(inc),pubexit(inc),4,pubobsent(inc),'IN',m.calexit)   &&&m.calentry+m.calexit
   			Wait Wind "calrexit     "+Str(m.calrexit)+"         1"
   			m.calfacdelta = m.calfacdelta+IIF(m.plexadd And !m.pformulas,m.varsubst,0)
   			Wait Wind "calfacdelta     "+Str(calfacdelta)+"         1"
   			m.calexit     = adjust(m.varsubst,4)
   			Wait Wind "calexit     "+Str(m.calexit)+"         1"
   			m.calfacdelta = m.calfacdelta+IIF(m.plexadd AND m.pformulas,m.calexit,0)
			If m.calfacdelta > 0 
				m.calalldelta = Iif(m.calfacdelta-m.plexbegin2 >0, m.plexfact2*(m.calfacdelta-m.plexbegin2),0) + Iif(m.calfacdelta-m.plexbegin1 >0, m.plexfact1*(m.calfacdelta-m.plexbegin1),0)								
			   	m.calalldelta = Adjust(m.calalldelta,4)
			EndIf   	
		Case m.varsubst < m.calnormal
					*** Early Exit ***
			m.varsubst    = m.varsubst - m.calnormal
			m.calrexit    = adjust(m.varsubst,3)
   			m.calexit     = getcalcul(m.varsubst,m.pend,3,pubobsext(1),'OUT')
   			m.calpenadd   = m.calpenadd + IIF(m.peeadd AND !m.pformulas,ABS(m.calexit),0)
   			m.calexit  	  = adjust(m.calexit,3)
   			m.calpenadd   = m.calpenadd + IIF(m.peeadd AND m.pformulas,ABS(m.calexit),0)
		   	m.calalldelta = Adjust(m.calpenadd,3)
		OtherWise
					*** On Time ***
			m.calrexit	  = 0
			m.calexit	  = 0		
			m.calalldelta = 0 
	EndCase	
	Return
EndIf	


*^^^^^^^^^^^^^^ Fixed Time ^^^^^^^^^^^^^^*

m.hol='HOL'+STR(YEAR(m.tdate),4)

Select (hol)
Seek Str(Month(tdate))+Str(Day(tdate))+Str(divtab)

If Found()
   m.calrequest = Iif(allday=1,0,m.ptotal+to-from)
EndIf

If !Empty(m.pubobstab)
   Select obstab
   Seek m.pubobstab
	
   m.calrequest = IIF(zero AND m.ptotal>0,0,m.calrequest)
EndIf

* Anormal presence 8
IF m.calrequest = 0 AND tnbr > 0
   FOR inc=1 TO tnbr
       m.calrabs = m.calrabs+pubexit(inc)-pubentry(inc)
       m.calabs = m.calabs+getcalcul(pubentry(inc),pubexit(inc),8,pubobsent(inc),'IN')
   ENDFOR
   m.calrabs = adjust(m.calrabs,8)
   m.calabs = adjust(m.calabs,8)
   m.calalldelta = m.calabs
   RETURN 
ENDIF


* Absence 7
IF tnbr = 0
   m.calrabs = m.calrequest
   m.calabs  = IIF(!m.pabs,0,m.calrabs)
   IF !EMPTY(m.pubobstab)
      SELECT obstab
      SEEK m.pubobstab
      IF !ab
         m.calabs  = 0
      ENDIF
   ENDIF
   m.calrabs = adjust(m.calrabs,7)
   m.calabs = adjust(m.calabs,7)
   m.calalldelta = m.calabs
   RETURN
ENDIF

* Early entry 1
IF pubentry(1) < m.peenbegin1
*	wait wind "early entry"
	inc = 1
   DO WHILE inc <=tnbr AND pubentry(inc) < m.pbegin
      IF inc > 1
         m.calentry = getcalcul(pubexit(inc-1),pubentry(inc),1,pubobsext(inc-1),'OUT',m.calentry)   &&&m.calentry+getcalcul(pubexit(inc-1),pubentry(inc),1,pubobsext(inc-1),'OUT',m.calentry)
      ENDIF
      IF pubexit(inc) < m.pbegin
         m.calrentry  = m.calrentry+pubexit(inc)-pubentry(inc)
         m.calentry = m.calentry+getcalcul(pubentry(inc),pubexit(inc),1,pubobsent(inc),'IN',m.calentry)
      ELSE
         m.calrentry  = m.calrentry+m.pbegin-pubentry(inc)
         m.calentry = m.calentry+getcalcul(pubentry(inc),m.pbegin,1,pubobsent(inc),'IN',m.calentry)
         EXIT
      ENDIF
      inc=inc+1
   ENDDO
   IF inc>1 AND inc<=tnbr AND BETWEEN(m.pbegin,pubexit(inc-1),pubentry(inc))
      m.calentry = m.calentry+getcalcul(pubexit(inc-1),m.pbegin,1,pubobsext(inc-1),'OUT',m.calentry)
   ENDIF
   m.calrentry   = adjust(m.calrentry,1)
   m.calfacdelta = m.calfacdelta+IIF(m.peenadd AND !m.pformulas,m.calentry,0)
   m.calentry    = adjust(m.calentry,1)
   m.calfacdelta = m.calfacdelta+IIF(m.peenadd AND m.pformulas,m.calentry,0)
ENDIF

* Late entry 2
IF pubentry(1) > m.plebegin
*	wait wind " late Entry "
*	wait wind Str(pubentry(1)) + Str(m.plebegin)
   m.calrentry  = adjust(pubentry(1)-m.pbegin,2)
   m.calentry   = getcalcul(m.pbegin,pubentry(1),2,pubobsent(1),'OUT',m.calentry)
   m.calpenadd  = m.calpenadd + IIF(m.pleadd AND !m.pformulas,ABS(m.calentry),0)
   m.calentry   = adjust(m.calentry,2)
   m.calpenadd  = m.calpenadd + IIF(m.pleadd AND m.pformulas,ABS(m.calentry),0)
ENDIF

* Early exit 3
IF pubexit(tnbr) < m.peebegin 
*		wait wind "early exit " 
*		wait wind "early exit" + Str(pubexit(tnbr) )+Str(m.peebegin) 
	If pubexit(tnbr) = 0 And peemb
		m.calrexit   = adjust(m.pend-Iif(pubentry(Tnbr) < m.pbegin,m.pbegin,pubentry(tnbr)),3)
   		m.calexit    = getcalcul(Iif(pubentry(Tnbr) < m.pbegin,m.pbegin,pubentry(tnbr)),m.pend,3,pubobsext(1),'OUT')
   		m.calpenadd  = m.calpenadd + IIF(m.peeadd AND !m.pformulas,ABS(m.calexit),0)
   		m.calexit    = adjust(m.calexit,3)
   		m.calpenadd  = m.calpenadd + IIF(m.peeadd AND m.pformulas,ABS(m.calexit),0)
   	Else	
		m.calrexit   = adjust(m.pend-Iif(pubexit(tnbr) < m.pbegin,m.pbegin,pubexit(tnbr)),3)
   		m.calexit    = getcalcul(Iif(pubexit(tnbr) < m.pbegin,m.pbegin,pubexit(tnbr)),m.pend,3,pubobsext(1),'OUT')
   		m.calpenadd  = m.calpenadd + IIF(m.peeadd AND !m.pformulas,ABS(m.calexit),0)
   		m.calexit    = adjust(m.calexit,3)
   		m.calpenadd  = m.calpenadd + IIF(m.peeadd AND m.pformulas,ABS(m.calexit),0)
   	EndIf	
ENDIF

* Late exit 4
IF pubexit(tnbr) > m.plexbegin1
   inc = tnbr
*   wait wind "Late Exit "
 *  Wait wind  Str(pubexit(tnbr))+" " +Str(m.plexbegin1)
   DO WHILE inc >= 1 AND pubexit(inc) > m.pend
      IF inc < tnbr
         m.calexit = m.calexit+getcalcul(pubentry(inc+1),pubexit(inc),4,pubobsent(inc+1),'IN',m.calexit)  &&m.calentry+m.calexit
      ENDIF
      IF pubentry(inc) > m.pend
         m.calrexit = m.calrexit+pubexit(inc)-pubentry(inc)
         m.calexit  = m.calexit+getcalcul(pubentry(inc),pubexit(inc),4,pubobsent(inc),'IN',m.calexit)   &&&m.calentry+m.calexit
      ELSE
		If Empty(pubentry(inc+1))
         	m.calrexit = m.calexit+pubexit(inc)-m.pend    &&m.calrexit+pubexit(inc)-m.pend
         	m.calexit  = getcalcul(m.pend,pubexit(inc),4,pubobsext(inc),'IN',m.calexit)  &&&m.calexit+getcalcul(m.pend,pubexit(inc),4,pubobsent(inc),'OUT',m.calexit)   &&&m.calentry+m.calexit
         Else
         	m.calrexit = m.calexit+pubexit(inc)-m.pend    &&m.calrexit+pubexit(inc)-m.pend
         	m.calexit  = getcalcul(m.pend,pubexit(inc),4,pubobsent(inc),'OUT',m.calexit)  &&&m.calexit+getcalcul(m.pend,pubexit(inc),4,pubobsent(inc),'OUT',m.calexit)   &&&m.calentry+m.calexit
         EndIf	
         EXIT
      ENDIF
      inc=inc-1
   ENDDO
   IF inc=>1 AND inc<tnbr AND BETWEEN(m.pend,pubentry(inc+1),pubexit(inc))
      m.calexit  = m.calexit+getcalcul(m.pend,pubentry(inc+1),4,pubobsext(inc),'OUT',m.calexit)  &&&m.calentry+m.calexit
   ENDIF
   m.calrexit  = adjust(m.calrexit,4)
   m.calfacdelta = m.calfacdelta+IIF(m.plexadd AND !m.pformulas,m.calexit,0)
   m.calexit   = adjust(m.calexit,4)
*   wait wind "Cal Exit " +Str(m.calexit)
*   wait wind "Cal rExit " +Str(m.calrexit)
   m.calfacdelta = m.calfacdelta+IIF(m.plexadd AND m.pformulas,m.calexit,0)
ENDIF

* Longue Presence 
m.rtbreak = 0
m.tbreak  = 0
m.tpres   = 0

For inc = 1 To tnbr
*	Wait Wind "Longue Presence "	
*	Wait Wind "Begin "+Str(m.pbegin) +"   " + "End "+Str(m.pend)+"  "+"Nbre "+Str(tnbr)
    IF (pubentry(inc) > m.pbegin AND inc > 1) And (Pubentry(inc) < m.pend)
*       		wait  wind "Longue "+Str(m.pbegin)+"  "+Str(pubentry(inc)) +"m.pbegin < pubentry(inc-1)"+"  "+"Nbre "+Str(tnbr)
       IF m.pbegin > pubexit(inc-1)
       		wait  wind "Longue "+Str(m.pbegin)+"  "+Str(pubentry(inc)) +"m.pbegin > pubexit(inc-1)"+"  "+"Nbre "+Str(tnbr)
          m.rtbreak  = m.rtbreak+pubentry(inc)-pubexit(inc-1)  &&M.PBEGIN
          m.tbreak   = m.tbreak +getcalcul(pubexit(inc-1),pubentry(inc),5,pubobsext(inc-1),'OUT',m.tbreak)
          m.tpres    = m.tpres  +getcalcul(pubexit(inc-1),pubentry(inc),6,pubobsext(inc-1),'OUT',m.tbreak)
* 		  WAIT WIND "m.rtbreak" + Str(m.rtbreak)+" " +"m.tbreak" + Str(m.tbreak)+"    1"
       ELSE
 *      		wait  wind "Longue "+Str(m.pbegin)+"  "+Str(pubentry(inc)) +"m.pbegin < pubexit(inc-1)"+"  "+"Nbre "+Str(tnbr)
          m.rtbreak  = m.rtbreak +pubentry(inc)-pubexit(inc-1)
          m.tbreak   = m.tbreak  +getcalcul(pubexit(inc-1),pubentry(inc),5,pubobsext(inc-1),'OUT',m.tbreak)
          m.tpres    = m.tpres   +getcalcul(pubexit(inc-1),pubentry(inc),6,pubobsext(inc-1),'OUT',m.tbreak)   
*		   wait wind "m.rtbreak" + Str(m.rtbreak)+" " +"m.tbreak" + Str(m.tbreak)+"    2" 
       ENDIF   
    ENDIF   
ENDFOR

*Wait wind "Total   "+"m.rtbreak" + Str(m.rtbreak)+" " +"m.tbreak" + Str(m.tbreak)
m.tnormbreak = m.pbreakout-m.pbreakin
*wait wind "m.tnormbreak"+str(m.tnormbreak)
*wait wind "m.plbbegin"+str(m.plbbegin)

* Missing Hours 5
IF m.rtbreak-tnormbreak > m.plbbegin
   m.calrpres = adjust(m.rtbreak-tnormbreak,5)
*	wait wind "Missing Hours 1"+Str(m.calrpres)	
ENDIF

IF m.tbreak-tnormbreak > m.plbbegin
   	m.calpres = IIF(m.plBadd AND !pformulas,adjust(m.tbreak-tnormbreak,5),0)
*	wait wind "Missing Hours 2"+Str(m.calrpres)	
ENDIF


*Missing Hours For Early Exit
For inc = 1 To tnbr
	If (pubentry(inc) > m.pend And inc > 1) And (Pubentry(inc) < m.pend)
		m.rtbreak = 0
		m.tbreak  = 0
		m.tpres   = 0

*		Wait Wind " Missing Hours " 
       	If m.pend < pubentry(inc)
*			Wait Wind " Missing Hours "+"    "+"m.pend < pubentry(inc)   "
        	m.rtbreak  = m.rtbreak+pubexit(inc-1)-pubentry(inc)  &&M.PBEGIN
        	m.tbreak   = m.tbreak +getcalcul(pubexit(inc-1),pubentry(inc),6,pubobsext(inc-1),'OUT',m.tbreak)
          	m.tpres    = m.tpres  +getcalcul(pubexit(inc-1),pubentry(inc),6,pubobsext(inc-1),'OUT',m.tbreak)
*		  	wait wind " Missing Hours "+"  "+"m.rtbreak "+Str(m.rtbreak)+"   "+"m.tbreak   "+str(m.tbreak) +"  "+"Nbre     1"
       	Else
*			Wait Wind " Missing Hours "+"    "+"m.pend > pubentry(inc)   "
          	m.rtbreak  = m.rtbreak +pubexit(inc-1)-pubentry(inc)
          	m.tbreak   = m.tbreak  +getcalcul(pubentry(inc),pubexit(inc-1),5,pubobsext(inc-1),'OUT',m.tbreak)
          	m.tpres    = m.tpres   +getcalcul(pubentry(inc),pubexit(inc-1),6,pubobsext(inc-1),'OUT',m.tbreak)   
*				Wait wind " Missing Hours "+"  "+"m.rtbreak "+Str(m.rtbreak)+"   "+"m.tbreak   "+str(m.tbreak) +"  "+"Nbre     2"
       	EndIf   
       	calrpres = 	adjust(m.rtbreak,6)
       	calpres  = 	adjust(m.tbreak,6)
    EndIf   
EndFor

*Wait wind " Missing Hours "+"  "+"m.rtbreak "+Str(m.rtbreak)+"   "+"m.tbreak   "+str(m.tbreak) +"  "+"Nbre     2"
*Wait wind " Missing Hours "+"  "+"m.tpres "+Str(m.tpres)
* Long presence 6
IF tnormbreak-m.rtbreak > m.plprbegin And !Empty(m.plprbegin)
   m.calrpres = adjust(tnormbreak-m.rtbreak,6)
ENDIF

IF tnormbreak-m.tbreak > m.plprbegin And !Empty(m.plprbegin)
   m.calpres = adjust(IIF(tnormbreak-m.tbreak>0,tnormbreak-m.tbreak,m.tpres),6)
ENDIF

* Delta
IF m.calpenadd > m.calfacdelta
   	if !Empty(m.calrpres) And m.calpres > 0
   		m.calfacdelta = m.calfacdelta + m.calpres*m.tcalfactor
   	EndIf	
   	if !Empty(m.calrpres) And m.calpres < 0
		m.calpenadd = m.calpenadd + Abs(m.calpres)
   	EndIf	
   m.calalldelta = adjust(m.calpenadd-m.calfacdelta,3)
ENDIF

IF m.pfactor AND m.calpenadd < m.calfacdelta
   	if !Empty(m.calrpres) And m.calpres > 0
   		m.calfacdelta = m.calfacdelta + m.calpres
   	EndIf
   m.calalldelta = (m.calfacdelta-m.calpenadd)*m.calfactor
   If m.calalldelta  > (m.plexbegin1-m.pend)
   		m.calalldelta = (m.plexbegin1-m.pend)*m.plexfact1 + (m.calalldelta-(m.plexbegin1-m.pend))*m.plexfact2
   EndIf	
   m.calalldelta = adjust(m.calalldelta,4)
ENDIF

IF !m.pfactor AND m.calpenadd < m.calfacdelta
   	If !Empty(m.calrpres) And m.calpres < 0
		m.calpenadd = m.calpenadd + Abs(m.calpres)
   	EndIf	

	m.sumdelta = 0
	For nj = 1 To ifac
		m.sumdelta = m.sumdelta + Iif(!Empty(fac(nj,1)),fac(nj,1)*fac(nj,2),0)
	EndFor 
 		m.calalldelta = m.sumdelta-m.calpenadd 
   	
   	If !Empty(m.calrpres) And m.calpres > 0
   		m.calalldelta = m.calalldelta + m.calpres*m.tcalfactor
   	EndIf	
   	
   	m.calalldelta = adjust(m.calalldelta,4)
ENDIF




PROCEDURE getcalcul
*******************
PARAMETERS tbegin,tend,ttype,tobstab,tin,tlast

IF tbegin>tend
   RETURN 0
ENDIF
m.ttime = 0
SELECT obstab
IF !EMPTY(tobstab)
   SEEK tobstab
   DO CASE
      CASE ttype=1       && Early entry 1
        IF m.peenbegin1>m.tend AND een AND m.peenadd
*           m.calfactor  = getfactor(m.tlast,m.calfactor,m.tend-m.tbegin,IIF(deffactor,m.peenfact1,eenfactor))
			fac(ifac,1)  = IIF(deffactor,m.peenfact1,eenfactor) 				
			fac(ifac,2)  = Adjust(m.tend-m.tbegin,1)
			ifac         = ifac + 1	
        ENDIF
        IF BETWEEN(m.peenbegin1,m.tbegin,m.tend) AND m.peen AND m.peenadd
*           m.calfactor  = getfactor(m.tlast,m.calfactor,peenbegin1-m.tbegin,IIF(deffactor,m.peenfact1,eenfactor))
 			fac(ifac,1)  = IIF(deffactor,m.peenfact1,eenfactor) 				
			fac(ifac,2)  = Adjust(peenbegin1-m.tbegin,1)
			ifac         = ifac + 1	
*           m.tlast      = m.tlast+(peenbegin1-m.tbegin)
*           m.calfactor  = getfactor(m.tlast,m.calfactor,m.tend-peenbegin1,IIF(deffactor,m.peenfact2,eenfactor))
 			fac(ifac,1)  = IIF(deffactor,m.peenfact2,eenfactor) 				
			fac(ifac,2)  = Adjust(m.tend-peenbegin1,1)
			ifac         = ifac + 1	
        ENDIF
        IF m.tbegin>m.peenbegin1 AND m.peen AND m.peenadd
*           m.calfactor   = getfactor(m.tlast,m.calfactor,m.tend-m.tbegin,IIF(deffactor,m.peenfact2,eenfactor))
 			fac(ifac,1)  = IIF(deffactor,m.peenfact2,eenfactor)	
			fac(ifac,2)  = Adjust(m.tend-m.tbegin,1)
			ifac         = ifac + 1	
        ENDIF
        RETURN IIF(!een,0,tend-tbegin)
      CASE ttype=2      && Late entry 2
        RETURN IIF(!le,0,tend-tbegin)
      CASE ttype=3      && Early exit 3
        RETURN IIF(!ee,0,tend-tbegin)
      CASE ttype=4      && Late exit 4
*		wait wind "obs late exit "
        IF m.plexbegin2<m.tbegin AND lex
*           m.calfactor  = getfactor(m.tlast,m.calfactor,m.tend-m.tbegin,IIF(deffactor,m.plexfact1,lexfact2))
 			fac(ifac,1)  = IIF(deffactor,m.plexfact1,lexfact2)	
			fac(ifac,2)  = Adjust(m.tend-m.tbegin,4)
			ifac         = ifac + 1	
        ENDIF     
        IF BETWEEN(m.plexbegin1,m.tbegin,m.tend) AND m.plex
*           m.calfactor  = getfactor(m.tlast,m.calfactor,m.plexbegin2-m.tbegin,IIF(deffactor,m.plexfact1,lexfact1))
 			fac(ifac,1)  = IIF(deffactor,m.plexfact1,lexfact1)
			fac(ifac,2)  = Adjust(m.plexbegin2-m.tbegin,4)
			ifac         = ifac + 1	
*           m.tlast      = m.tlast+(peenbegin2-m.tbegin)
*           m.tlast      = m.tlast+(plexbegin2-m.tbegin)
*           m.calfactor  = getfactor(m.tlast,m.calfactor,m.tend-plexbegin2,IIF(deffactor,m.plexfact2,lexfact2))
 			fac(ifac,1)  = IIF(deffactor,m.plexfact2,lexfact2)
			fac(ifac,2)  = Adjust(m.tend-plexbegin2,4)
			ifac         = ifac + 1	
        ENDIF
        IF m.plexbegin2>m.tend AND m.plex
*           m.calfactor  = getfactor(m.tlast,m.calfactor,m.tend-m.tbegin,IIF(deffactor,m.plexfact2,lexfact1))
 			fac(ifac,1)  = IIF(deffactor,m.plexfact2,lexfact1)
			fac(ifac,2)  = Adjust(m.tend-m.tbegin,4)
			ifac         = ifac + 1	
        ENDIF
        RETURN IIF(!lex,0,tend-tbegin)
      CASE ttype=5
*       m.calallpen = m.calpenadd+IIF(de,tend-tbegin,0)
*       m.calpenadd = m.calpenadd+IIF(m.plbadd AND de,tend-tbegin,0)
		RETURN IIF(!de,0,tend-tbegin)
      CASE ttype=6
		DO CASE
		  CASE dpr AND m.plpradd AND .F.
*            m.calfactor   = getfactor(m.tlast,m.calfactor,m.tend-m.tbegin,IIF(m.deffactor,m.plprfact,dprfactor))			             
          ENDCASE
      RETURN IIF(!dpr,0,tend-tbegin)
      CASE ttype=7
      CASE ttype=8
          RETURN IIF(!ovh,0,tend-tbegin)
	  ENDCASE
ELSE
   IF (tin='IN' AND (ttype=1 OR ttype=4 OR ttype=6 OR ttype=8)) OR (tin='OUT' AND (ttype=2 OR ttype=3 OR ttype=5 Or ttype=6))
      DO CASE
      CASE ttype = 1       && Early entry 1
         IF m.peenbegin1>m.tend AND m.peen AND m.peenadd
*            m.calfactor   = getfactor(m.tlast,m.calfactor,m.tend-m.tbegin,m.peenfact1)
			fac(ifac,1)   = m.peenfact1
			fac(ifac,2)	  = Adjust(m.tend-m.tbegin,1)
			ifac          = ifac + 1	
         ENDIF   
         IF BETWEEN(m.peenbegin1,m.tbegin,m.tend) AND m.peen AND m.peenadd
*            m.calfactor   = getfactor(m.tlast,m.calfactor,peenbegin1-m.tbegin,m.peenfact1)
			fac(ifac,1)   = m.peenfact1
			fac(ifac,2)	  = Adjust(peenbegin1-m.tbegin,1)
			ifac          = ifac + 1	
*           m.tlast       = m.tlast+(peenbegin1-m.tbegin)
*            m.calfactor   = getfactor(m.tlast,m.calfactor,m.tend-peenbegin1,m.peenfact2)
			fac(ifac,1)   = m.peenfact2
			fac(ifac,2)	  = Adjust(m.tend-peenbegin1,1)
			ifac          = ifac + 1	
         ENDIF   
         IF m.tbegin>m.peenbegin1 AND m.peen AND m.peenadd
*            m.calfactor   = getfactor(m.tlast,m.calfactor,m.tend-m.tbegin,m.peenfact2)
			fac(ifac,1)   = m.peenfact2
			fac(ifac,2)	  = Adjust(m.tend-m.tbegin,1)
			ifac          = ifac + 1	
         ENDIF
         RETURN IIF(!m.peen,0,tend-tbegin)
      CASE ttype = 2       && Late entry 2
           RETURN IIF(!m.ple,0,tend-tbegin)
      CASE ttype = 3
           RETURN IIF(!m.pee,0,tend-tbegin)
      CASE ttype = 4       && Late exit 4
         IF m.plexbegin2<m.tbegin AND m.plex AND m.plexadd
*               m.calfactor   = getfactor(m.tlast,m.calfactor,m.tend-m.tbegin,m.plexfact2)
       			fac(ifac,1) = m.plexfact2
       			fac(ifac,2) = Adjust(IIF(!m.plex,0,tend-tbegin),4)
       			ifac = ifac + 1 
         ENDIF     
         IF BETWEEN(m.plexbegin2,m.tbegin,m.tend) AND m.plex AND m.plexadd
*              	m.calfactor   = getfactor(m.tlast,m.calfactor,plexbegin2-m.tbegin,m.plexfact1)
				fac(ifac,1) = m.plexfact1
				fac(ifac,2) = Adjust(plexbegin2-m.tbegin,4)
				ifac = ifac + 1
*          	  	m.tlast       = m.tlast+(peenbegin2-m.tbegin)
*               m.tlast       = m.tlast+(plexbegin2-m.tbegin)
*               m.calfactor   = getfactor(m.tlast,m.calfactor,m.tend-plexbegin2,m.plexfact2)
				fac(ifac,1) = m.plexfact2
				fac(ifac,2) = Adjust(m.tend-plexbegin2,4)
				ifac = ifac + 1
         ENDIF
         IF m.plexbegin2>m.tend AND m.plex AND m.plexadd
*               m.calfactor   = getfactor(m.tlast,m.calfactor,m.tend-m.tbegin,m.plexfact1)
				fac(ifac,1) = m.plexfact1
				fac(ifac,2) = Adjust(m.tend-m.tbegin,4)
				ifac = ifac + 1
        ENDIF
*        	wait wind "nicholas " + Str(tend-tbegin)
        
         RETURN IIF(!m.plex,0,tend-tbegin)
      CASE ttype = 5
*           m.calpenadd = m.calpenadd+IIF(m.plbadd AND m.plb,tend-tbegin,0)
           RETURN IIF(!m.plb,0,tend-tbegin)
      CASE ttype = 6
         DO CASE
         	CASE m.plpr AND m.plpradd  
*			wait wind " long pres for overtime "
              m.tcalfactor   = getfactor(m.tlast,m.calfactor,m.tend-m.tbegin,m.plprfact)
              Return Iif(!m.plpr,0,tbegin-tend)
         ENDCASE
         	RETURN IIF(!m.plpr,0,tend-tbegin)
      CASE ttype = 8
           RETURN IIF(!m.papr,0,tend-tbegin)
   ENDCASE
   ENDIF   
ENDIF
RETURN m.ttime

PROCEDURE adjust
****************
PARAMETERS ttime,ttype

IF ttime=0
   RETURN 0
ENDIF   
m.bidon = ttime
m.pmin  = 0
m.ppen  = 0
m.ptab  = 0
DO CASE
   CASE ttype = 1
        m.ppen = m.peenpen
        m.pmin = m.peenmin
        m.ptab = m.peentab
   CASE ttype = 2
        m.ppen=m.plepen
        m.pmin=m.plemin
        m.ptab=m.pletab
   CASE ttype = 3 OR ttype = 7
        m.ppen=m.peepen
        m.pmin=m.peemin
        m.ptab=m.peetab
   CASE ttype = 4 OR ttype = 8
        m.ppen=m.plexpen
        m.pmin=m.plexmin
        m.ptab=m.plextab
   CASE ttype = 5
        m.ppen=m.plbpen
        m.pmin=m.plbmin
        m.ptab=m.plbtab
   CASE ttype = 6 
        m.ppen=m.plprpen
        m.pmin=m.plprmin
        m.ptab=m.plprtab
ENDCASE

IF (m.ppen <> 0)  .AND. ttime > 0 .AND. m.ptab = 0
   IF (ttime%m.ppen) > m.pmin
       m.bidon=INT(ttime/m.ppen)*m.ppen+m.ppen
   ELSE
       m.bidon=INT(ttime/m.ppen)*m.ppen
   ENDIF
ENDIF 
IF m.ptab <> 0
   m.bidon = getpenal(ttime,m.ptab)
ENDIF
RETURN IIF(ttype=1 OR ttype=4 OR ttype=6 OR ttype=8,m.bidon,-m.bidon)

PROCEDURE getext
****************
PARAMETERS tdate

RETURN PADL(MONTH(tdate),2,'0')+STR(YEAR(tdate),4)

PROCEDURE getselect
*******************
PARAMETERS tfile,talias,torder

IF !USED(tfile)
   SELECT (talias)
   USE
   USE (tfile)
   IF !EMPTY(torder)
      SET ORDER TO &torder
   ENDIF   
ELSE
   SELECT (tfile)
ENDIF
RETURN

Procedure filtrans
*================*
Parameters ttrfile,tobsfile,temptab,tdate
Select (ttrfile)
Seek Str(temptab)+Str(Day(m.tdate))
tnbr = 0
m.calwork = 0
=Reftrans()
tnbr		= 0
m.ckbegin   = 0
m.ckend     = 0
m.cknormal  = 0
m.ckconsnight = .F.
m.chltbegin = 0
m.chltend   = 0
m.chltnormal= 0
m.chatbegin = 0
m.chatend   = 0
m.chatnormal= 0
m.chltconsnight = .F.

=ChckZone(EmpTab.serial,tdate-1,.T.)
m.chltbegin = m.ckbegin
m.chltend   = m.ckend
m.chltnormal= m.cknormal
m.chltconsnight = m.ckconsnight

=ChckZone(EmpTab.serial,tdate+1,.F.)
m.chatbegin = m.ckbegin
m.chatend   = m.ckend
m.chatnormal= m.cknormal

*Wait wind "Last    " + Str(m.chltbegin) + "  " + Str(m.chltend) 
*Wait wind "After   " + Str(m.chatbegin) + "  " + Str(m.chatend) 

If m.pbegin > m.pend And m.ptype != 2 And m.ptype != 3
	m.pubnight = .T.
	If m.chltnormal != 1 And m.chatnormal != 1
		*** Day Off For Yesterday And Next Day ***
		=SchTr(Iif(!m.pconsnight,tdate  ,tdate-1))	
		=SchTr(Iif(!m.pconsnight,tdate+1,tdate))	
		Return m.tnbr
	EndIf	
	Do Case
		Case !m.povernight And !pconsnight
			*** Over Early Entry , Night Shift , Valid. Entry ***
			If (m.chltnormal = 1) And (m.chltbegin > m.chltend) And (m.chatnormal = 1) And (m.chatbegin < m.chatend)
				*** Night Shift Before & Normal After ***
				=SchTr1(tdate  ,"entry",m.chltend,">")	
				=SchTr(tdate+1)	
				Return m.tnbr
			EndIf
			If (m.chltnormal = 1) And (m.chltbegin > m.chltend) And (m.chatnormal = 1) And (m.chatbegin > m.chatend)	Or ((m.chltnormal = 1) And (m.chltbegin > m.chltend) And (m.chatnormal != 1))
				*** Night Shift For Yesterday And Next Day ***
				=SchTr1(tdate  ,"entry",m.chltend,">")	
				=SchTr1(tdate+1,"entry",m.pend   ,"<=")	
			Else
			 	*** Normal Shift Yesterday And Night Shift Next Day 	***
			 	*** Normal Shift Yesterday And Next Day 			 	***
				*** Day Off Yesterday And Night Shift Next Day 		***
				*** Day Off Yesterday And Normal Shift Next Day		***
				*** Normal Before & Day Off After ***
*				Wait wind "Day Off Yesterday And Night Shift Next Day " +Str(m.pend)
				=SchTr(tdate)	
				=SchTr1(tdate+1,"entry",m.pend,"<=")	
			EndIf
		Case m.povernight And !pconsnight
			*** Over Late Exit , Night Shift , Valid. Exit ***
			If (m.chltnormal = 1) And (m.chltbegin < m.chltend) And (m.chatnormal = 1) And (m.chatbegin < m.chatend)
				wait wind " nicholas" 
				=SchTr1(tdate  ,"exit" ,m.pbegin   ,">")	
				=SchTr1(tdate+1 ,"entry" ,m.pend    ,"<=")	
				Return m.tnbr
			EndIf
			If (m.chltnormal = 1) And (m.chltbegin > m.chltend) And (m.chatnormal != 1) 
				wait wind " nicholas" 
				=SchTr1(tdate  ,"exit" ,m.pbegin   ,">")	
				=SchTr(tdate+1)
				Return m.tnbr
			EndIf
			If (m.chltnormal = 1) And (m.chltbegin < m.chltend) And (m.chatnormal = 1) And (m.chatbegin > m.chatend)
				wait wind " nicholas" 
				=SchTr(tdate)	
				=SchTr1(tdate+1 ,"entry" ,m.chatbegin ,"<=")	
				Return m.tnbr
			EndIf
			If (m.chltnormal != 1) And (m.chatnormal = 1) And (m.chatbegin < m.chatend) 
				 *** Day Off Yesterday And Normal Shift Next Day		***
				=SchTr(tdate)	
				=SchTr1(tdate+1,"entry",m.pend,"<=")	
			Else
				=SchTr1(tdate  ,"exit" ,m.pbegin   ,">")	
				=SchTr1(tdate+1,"entry",m.chatbegin,"<=")	
			EndIf
			If (m.chltnormal = 1) And (m.chltbegin > m.chltend) And (m.chatnormal = 1) 
					*** Normal Before & Day Off After ***
				=SchTr(tdate)	
				=SchTr(tdate+1)
			EndIf		
			If (m.chltnormal = 1) And (m.chltbegin < m.chltend) And (m.chatnormal != 1) And (m.chatbegin < m.chatend) 
					*** Night Shift Before & Normal After ***
				=SchTr1(tdate  ,"exit" ,m.pbegin   ,">")	
				=SchTr(tdate+1,"entry",m.pend ,"<=")
			EndIf		
		Case !m.povernight And pconsnight
			 *** Over Early Entry , Night Shift , Valid. Entry ***
			If (m.chltnormal = 1) And (m.chltbegin > m.chltend) And (m.chatnormal = 1) And (m.chatbegin > m.chatend)
					*** Night Shift Before & After ***
				=SchTr1(tdate-1,"entry",m.chltend,">")	
				=SchTr1(tdate  ,"entry" ,m.pend   ,"<=")
			EndIf		
			If (m.chltnormal = 1) And (m.chltbegin < m.chltend) And (m.chatnormal = 1) And (m.chatbegin > m.chatend)
					*** Normal Before & Night Shift After ***
				=SchTr1(tdate-1,"exit",m.pbegin,">=")	
				=SchTr1(tdate  ,"entry" ,m.pend   ,"<=")	
			EndIf
			If (m.chltnormal = 1) And (m.chltbegin < m.chltend) And (m.chatnormal = 1) And (m.chatbegin < m.chatend)
					*** Normal Before & After ***
				=SchTr1(tdate-1,"exit",m.pbegin,">=")	
				=SchTr(tdate)	
			EndIf
			If (m.chltnormal != 1) And (m.chatnormal = 1) And (m.chatbegin > m.chatend)
					*** Day Off Before & Night Shift After ***
				=SchTr1(tdate-1,"exit",m.pbegin,">=")	
				=SchTr1(tdate  ,"entry" ,m.pend   ,"<=")	
			EndIf
			If (m.chltnormal != 1) And (m.chatnormal = 1) And (m.chatbegin < m.chatend)
					*** Day Off Before & Normal After ***
				=SchTr1(tdate-1,"exit",m.pbegin,">=")	
				=SchTr(tdate)	
			EndIf
			If (m.chltnormal = 1) And (m.chltbegin > m.chltend) And (m.chatnormal != 1) 
					*** Night Shift Before & After ***
				=SchTr1(tdate-1,"entry",m.chltend,">")	
				=SchTr(tdate)
			EndIf		
			If (m.chltnormal = 1) And (m.chltbegin < m.chltend) And (m.chatnormal != 1) 
					*** Normal Before & Day Off After ***
				=SchTr1(tdate-1,"exit",m.pbegin,">=")	
				=SchTr(tdate)
			EndIf		
			If (m.chltnormal = 1) And (m.chltbegin > m.chltend) And (m.chatnormal = 1) And And (m.chatbegin < m.chatend)
					*** Night Shift Before & Normal After ***
				=SchTr1(tdate-1,"entry",m.chltend,">")	
				=SchTr(tdate)
			EndIf		
		Case m.povernight And pconsnight
			If (m.chltnormal = 1) And (m.chltbegin < m.chltend) And (m.chatnormal = 1) And (m.chatbegin < m.chatend)
					*** Normal Before &  After ***
				=SchTr1(tdate-1,"exit"  ,m.pbegin      ,">=")	
				=SchTr1(tdate  ,"entry" ,m.pend , "<=")
			EndIf		
			If (m.chltnormal = 1) And (m.chltbegin < m.chltend) And (m.chatnormal = 1) And (m.chatbegin > m.chatend)
					*** Normal Before & Night Shift After ***
				=SchTr1(tdate-1,"exit"  ,m.pbegin      ,">=")	
				=SchTr(tdate)
			EndIf
			If (m.chltnormal = 1) And (m.chltbegin < m.chltend) And (m.chatnormal != 1) 
					*** Normal Before & Day Off After ***
					wait wind "day off"
				=SchTr1(tdate-1,"exit"  ,m.pbegin      ,">=")	
				=SchTr(tdate)
			EndIf
			If (m.chltnormal = 1) And (m.chltbegin > m.chltend) And (m.chatnormal = 1) And (m.chatbegin > m.chatend)
					*** Night Shift Before & After ***
				=SchTr1(tdate-1,"exit"  ,m.pbegin      ,">=")	
				=SchTr1(tdate  ,"entry" ,m.chatbegin   ,"<=")
			EndIf		
			If (m.chltnormal = 1) And (m.chltbegin > m.chltend) And (m.chatnormal = 1) And (m.chatbegin < m.chatend)
					*** Night Shift Before & Normal After ***
				=SchTr1(tdate-1,"exit"  ,m.pbegin      ,">=")	
				=SchTr(tdate)
			EndIf		
			If (m.chltnormal = 1) And (m.chltbegin > m.chltend) And (m.chatnormal != 1) 
					*** Night Shift Before & Day Off After ***
				=SchTr1(tdate-1,"exit"  ,m.pbegin      ,">=")	
				=SchTr(tdate)
			EndIf		
			If (m.chltnormal != 1) And (m.chltbegin > m.chltend) And (m.chatnormal = 1) And (m.chatbegin < m.chatend)
					*** Day Off Before & Night Shift After ***
				=SchTr1(tdate-1,"exit"  ,m.pbegin      ,">=")	
				=SchTr1(tdate  ,"entry" ,m.chatbegin   ,"<=")
			EndIf		
			If (m.chltnormal != 1) And (m.chltbegin > m.chltend) And (m.chatnormal = 1) And (m.chatbegin < m.chatend)
					*** Day Off Before & Normal After ***
				=SchTr1(tdate-1,"exit"  ,m.pbegin      ,">=")	
				=SchTr(tdate)
			EndIf		
	EndCase	
Else
	wait wind "normal shift " nowait
	m.pubnight = .F.
	Do Case
		Case m.chltbegin > m.chltend And !m.chltconsnight And m.ptype !=2
			*** Night Shift Yesterday ***
			=SchTr1(tdate,"entry",chltend,">")	
			=GetCheckBad(tdate)
		Case m.chltbegin > m.chltend And m.ptype !=2 
			*** Night Shift Next Day ***
			=SchTr1(tdate,"exit",chatend,"<")	
			=GetCheckBad(tdate)
		OtherWise	
			=SchTr(tdate)	
			=GetCheckBad(tdate)
	EndCase
EndIf
Return tnbr

PROCEDURE getcl
***************
m.calwday     = m.wday
m.calnormal   = m.normal
m.calrequest  = m.request
m.calalldelta = m.alldelta
m.calrabs     = m.rabs
m.calabs      = m.abs
m.calrentry   = m.rentry
m.calentry    = m.entry
m.calrexit    = m.rexit
m.calexit     = m.exit
m.calrpres    = m.rpres
m.calpres     = m.pres
m.calwork     = m.work
RETURN

PROCEDURE setcl
***************
m.wday     = m.calwday
m.normal   = m.calnormal
m.request  = m.calrequest
m.alldelta = m.calalldelta
m.rabs     = m.calrabs
m.abs      = m.calabs
m.rentry   = m.calrentry
m.entry    = m.calentry
m.rexit    = m.calrexit
m.exit     = m.calexit
m.rpres    = m.calrpres
m.pres     = m.calpres
m.work     = m.calwork
RETURN


PROCEDURE getsch
****************
FOR tschi = 1 TO FCOUNT()
    m = FIELD(tschi)
    n = 'm.p'+FIELD(tschi)
    &n = &m
ENDFOR
RETURN


PROCEDURE calpara
*****************
PARAMETERS tcaldate,tcalemptab

=getzone(tcaldate,tcalemptab,0)
trfile  = 'TR'+getext(m.tcaldate)
dlfile  = 'DL'+getext(m.tcaldate)
obfile  = 'OB'+getext(m.tcaldate)
m.tcalnbr=filtrans('TR'+getext(tcaldate),'OB'+getext(tcaldate),tcalemptab,tcaldate)
SELECT (dlfile)
SEEK STR(tcalemptab)+STR(DAY(tcaldate))
m.pubobstab = obstab
DO getpara WITH m.tcaldate,m.tcalnbr
RETURN

PROCEDURE getdatabase
*********************
PARAMETERS tdattadate

pubclfile  = 'CL'+getext(tdattadate)
pubtrfile  = 'TR'+getext(tdattadate)
pubdlfile  = 'DL'+getext(tdattadate)
pubobfile  = 'OB'+getext(tdattadate)
pubholfile = 'HOL'+STR(YEAR(tdattadate),4)
=getselect(pubclfile,120,'empday')
=getselect(pubtrfile,121,'empdayord')
=getselect(pubdlfile,122,'empday')
=getselect(pubobfile,123,'serial')
=getselect(pubholfile,124,'mondaydiv')
RETURN

PROCEDURE checkdatabase
***********************
PARAMETERS tdate

IF !USED('CL'+getext(m.tdate))
   IF !FILE('CL'+getext(m.tdate)+'.DBF')
      RETURN .F.
   ENDIF   
ENDIF
RETURN .T.

PROCEDURE drawline
******************
PARAMETERS tbeginx , tbeginy , tlenx , tleny , tcolor

@ tbeginx,tbeginy CLEAR TO tbeginx+tlenx,tbeginy+tleny 
@ tbeginx,tbeginy FILL  TO tbeginx+tlenx,tbeginy+tleny  COLOR RGB(&tcolor,&tcolor)

PROCEDURE crifil
****************
SELECT bratab
SET ORDER TO serial
SELECT dpttab
SET ORDER TO serial
SET RELATION TO bratab INTO bratab ADDITIVE
SELECT divtab
SET ORDER TO serial
SET RELATION TO dpttab INTO dpttab ADDITIVE
SELECT pertab
SET ORDER TO emptab
SELECT emptab
SET ORDER TO serial
SET RELATION TO serial INTO pertab ADDITIVE
SELECT emptmp1
SET RELATION TO serial INTO emptab ADDITIVE
FOR I=3 TO 7
    SELECT (cridata(I))
    SET ORDER TO serial
    SELECT emptab
    t=cridata(I)
    SET RELATION TO &t INTO &t ADDITIVE
ENDFOR
tfil = IIF(pubandor=1,'.T.','.F.')
FOR I=1 TO 9
    IF pubcriteria(I)<>0
       tfil = tfil+IIF(pubandor=1,' AND ',' OR ')+'!EMPTY('+cridata(I)+'.bidon)'
    ENDIF
ENDFOR
IF tfil == '.F.'
   tfil = '.T.'
ENDIF
RETURN '('+tfil+')'


PROCEDURE getal
***************
PARAMETERS tdate,tcgntab

IF YEAR(tdate) = 0
   RETURN 0
ENDIF
IF YEAR(tdate) = YEAR(DATE())
   RETURN 0
ENDIF
SELECT cgntab
IF tcgntab <> 0
   GO RECORD tcgntab
ELSE
   RETURN 0   
ENDIF
IF alyear1 > YEAR(DATE())-YEAR(tdate)
   RETURN 0
ENDIF   
FOR AN=1 TO 6
    OAN='ALYEAR'+STR(AN,1)
    OLE='ALDAY'+STR(AN,1)
    NAN='ALYEAR'+STR(AN+1,1)
    IF &NAN = 0
       RETURN &OLE
    ENDIF   
    IF YEAR(DATE())-YEAR(tdate) >= &OAN .AND. YEAR(DATE())-YEAR(tdate)<=&NAN
       RETURN  &OLE
    ENDIF
ENDFOR    
RETURN ALDAY7

PROCEDURE getpenal
******************
PARAMETERS tvalue,tptab

IF tptab = 0
   RETURN tvalue
ENDIF
SELECT pentab
SEEK m.tptab
IF !FOUND()
   RETURN tvalue
ENDIF
IF tvalue <= time1
   RETURN factor1
ENDIF
T31 = .F.
O='factor1'
FOR ti=1 TO 6
    M = 'time'+STR(ti,1)
    N = 'time'+STR(ti+1,1)
   
    IF &N = 0
    	T31 = .T.
    ENDIF
   
    IF &m < tvalue AND tvalue <= &n
       O = 'factor'+STR(ti+1,1)
       RETURN &O
    ENDIF
ENDFOR
    IF T31 
       	RETURN tvalue
    ELSE
    	RETURN &O
    ENDIF
RETURN factor7

PROCEDURE getfactor
*******************
PARAMETERS tolddelta,toldfactor,tnewdelta,tnewfactor

IF (tolddelta+tnewdelta) = 0
   RETURN toldfactor
ENDIF
RETURN ((tolddelta*toldfactor)+(tnewdelta*tnewfactor))/(tolddelta+tnewdelta)

PROCEDURE openall
*****************
PARAMETERS tdate1,tdate2

m.ttopen = MONTH(tdate1)+YEAR(tdate1)*100
FOR topeni=MONTH(tdate1)+YEAR(tdate1)*100 TO MONTH(tdate2)+YEAR(tdate2)*100
    ttdate = CTOD( "01/"+STR(topeni%100,2)+'/'+STR(topeni/100,4,0) )
    clfile  = 'CL'+getext(ttdate)
    trfile  = 'TR'+getext(ttdate)
    dlfile  = 'DL'+getext(ttdate)
    obfile  = 'OB'+getext(ttdate)
    holfile = 'HOL'+STR(YEAR(ttdate),4)
    =getselect(clfile,100+(topeni-ttopen)*5,'empday')
    =getselect(trfile,101+(topeni-ttopen)*5,'empdayord')
    =getselect(dlfile,102+(topeni-ttopen)*5,'empday')
    =getselect(obfile,103+(topeni-ttopen)*5,'serial')
    =getselect(holfile,104+(topeni-ttopen)*5,'mondaydiv')
ENDFOR    
RETURN


PROCEDURE closeall
*****************
FOR tclosei=100 TO 225
    tclosej=str(i,3)
    select &tclosej
    use
ENDFOR

FUNCTION cypher
***************
PARAMETERS var

vlen = LEN(var)
npas = ""
FOR i = 1 TO vlen
	npas = npas + CHR(ASC(SUBSTR(var,i,1))-i)
ENDFOR
RETURN npas

FUNCTION uncypher
*****************
PARAMETERS ttvar

ttvar = ALLTRIM(ttvar)
vlen = LEN(ttvar)
npas = ""
FOR i = 1 TO vlen
	npas = npas + CHR(ASC(SUBSTR(ttvar,i,1))+i)
ENDFOR
RETURN npas

Procedure GetMark
*===============*
Parameters leent ,leexit ,lebreak ,absent 
If m.absent <> 0
	filmark(1) = "X"
	filmark(2) = "X"
EndIf	

If  leent < 0 Or leent > 0
	If pubentry(1) < m.pbegin And pubexit(1) < m.pbegin
		filmark(1)	= "X"
		filmark(2)	= "X"
	Else
		filmark(1)	= "X"
	EndIf	
EndIf

If leexit < 0 Or leexit > 0
	For i = 1 To 30 
		If GetEmpty(i)
			Exit
		EndIf	
		If pubexit(i) > m.pend And pubentry(i) > m.pend And leexit > 0
			If i <= 4
				filmark(2*i)   = "X"
				filmark(2*i-1) = "X"
				Exit 
			EndIf	
		Else
			If pubexit(i) <> m.pend  And Empty(pubexit(i+1))
				If i <= 4
					filmark(2*i)   = "X"
				EndIf	
			EndIf	
		EndIf	
	EndFor 
EndIf 	
If lebreak < 0 Or lebreak > 0
	For l = 2 To 30 
		If GetEmpty(l)
			Exit
		EndIf	
		If pubexit(l-1) < m.pbreakin Or (pubexit(l-1) > m.pbreakin And pubexit(l-1) < m.pend)
			filmark(2*(l-1))  = "X"
		EndIf	
		If pubentry(l) < m.pbreakout Or pubentry(l) > m.pbreakout
			filmark(2*l-1)  = "X"
		EndIf	
	EndFor 
EndIf

Procedure GetObsFile
*==================*
Select (pubclfile)
Seek Str(emptab.serial)+Str(DAY(m.rdate))
If !Found()
		Return .F.
EndIf	
Return .T.

Procedure GetGroup
*================*
Do Case
	Case grp1 = 1
		m.namegrp = ""
		m.grpname = ""
		Return 0
	Case grp2 = 1
		m.namegrp = Bratab.name
		m.grpname = Iif(publang,"Branch","Branche")
		Return BraTab.serial
	Case grp3 = 1
		m.namegrp = DptTab.name
		m.grpname = Iif(publang,"Department","Departement")
		Return DptTab.serial
	Case grp4 = 1
		m.namegrp = DivTab.name
		m.grpname = m.pubfield4
		Return DivTab.serial
EndCase		
		
Procedure GetEmpty
*================*
Parameters nfor	
If Empty(pubentry(nfor)) And Empty(pubexit(nfor))
	Return .T.
Else
	Return .F.
EndIf	
	
Procedure DeactPop
*================*
Deactivate Popup 	

Procedure GetZap
*==============*
Parameters zapfile
Select (zapfile)
If Alias() = m.zapfile
	Zap
	Return .T.
Else
	Wait Wind Iif(m.publang,"File Not Found " , "Fichier Inexistant ")+m.zapfile  Nowait	 
	Return .F.
EndIf	

Procedure variableTime
*====================*
m.varsubst = 0
For vi = 1 To 30
	m.time = 0
	Do Case
		Case Empty(pubentry(vi)) And !Empty(pubexit(vi))
			m.varsubst = m.varsubst + pubexit(vi)+1439 - Iif(m.time = 0,1439,m.time) 
		Case !Empty(pubentry(vi)) And Empty(pubexit(vi))
			=ReceiveVAr(tdate+1,"X",tdate,m.pend,vi)			
			m.varsubst = m.varsubst + m.time+1439 - pubentry(vi)
		OtherWise
			m.varsubst = m.varsubst + pubexit(vi) - pubentry(vi)
	EndCase		
EndFor		
Return m.varsubst 

Procedure ReceiveVAr
*==================*
Parameters recdat ,ppty ,recnbr
Select (ttrfile)
m.rectrfile  = 'TR'+getext(m.recdat)
=getselect(rectrfile,225,'empdayord')
Seek Str(EmpTab.serial)+Str(Dow(recdat))
Do While !Eof() And emptab = Emptab.serial And day = Dow(recdat)
	Do Case
		Case m.ppty = "E" And entry > m.chltend And Empty(exit) And !Empty(entry) 
			wait wind "not empty entry & empty exit "
			pubentry(m.recnbr) 	= entry
			pubnserial(m.recnbr)= serial
		Case m.ppty = "X" And Empty(entry) And !Empty(exit) And  exit < m.chatbegin
			wait wind "not empty entry & empty exit "
			pubentry(m.recnbr) 	= exit
			pubnserial(m.recnbr)= serial
	EndCase				
	Skip
EndDo
Return 

Procedure SpecialCase
*===================*
Parameters m.bvarspec ,m.evarspec  
Return Iif(m.bvarspec > m.evarspec ,m.evarspec+1439,m.evarspec)

Procedure Reftrans
*================*
For i = 1 To 30
    pubnserial(I) 	= 0
    pubserial(I) 	= 0
    pubentry(I)  	= 0
    pubexit(I)   	= 0
    pubobsent(I) 	= 0
    pubobsext(I) 	= 0
    pubhsent(I)  	= 0
    pubhsext(I)  	= 0
    pubdhsent(I) 	= 0
    pubdhsext(I) 	= 0
    pubcoment(I) 	= ''
    pubcomext(I) 	= ''
EndFor

Procedure RefSpecial
*==================*
Parameters refsep
Return Iif(refsep > 1439 ,(refsep - 1439),refsep)

Procedure ChckZone
*================*
Parameters chper ,chdat ,chklog
m.ckbegin 	= 0
m.ckend 	= 0
m.cknormal 	= 0
m.ckconsnight 	= .F.

Select schupd
m.chrec = Recno()
Seek Str(chper)+Dtoc(chdat)
If Found()
	m.ckbegin 	= pbegin
	m.ckend 	= pend
	m.cknormal 	= pnormal
	m.ckconsnight 	= pconsnight
	Go m.chrec
   	Return 1
EndIf

Select schtra
Seek chper

If !Found()
	m.ckbegin 	= 0
	m.ckend 	= 0
	m.cknormal 	= 0
	m.ckconsnight 	= .F.
	Return 0
EndIf

Do While pdate <= chdat And pemptab = chper And !Eof()
	m.ckbegin 	= pbegin
	m.ckend 	= pend
	m.cknormal 	= pnormal
	m.ckconsnight 	= pconsnight
*   	Skip protation*7-(chdat-startdate)%(protation*7)
	Skip
EndDo
Return 1

Procedure SavNight
*================*
Parameters savtime
Do Case
	Case !m.pconsnight
		m.savtime = Iif(m.pbegin > (m.savtime-1439) And (m.savtime-1439) >= 0  ,m.savtime,-1)
	Case m.pconsnight
		m.savtime = Iif(m.pbegin > (m.savtime-1439) And (m.savtime-1439) >= 0  ,m.savtime,-1)
EndCase
Return savtime

Procedure GetNight
*================*
Parameters gettime
If m.pconsnight
	If Obtmp.nserial != 0
		m.gettime = Iif(m.pbegin > m.gettime ,m.gettime,0)
	EndIf	
Else	
	m.gettime = Iif(m.pbegin < m.gettime ,m.gettime,0)
EndIf	
Return gettime

Procedure SchTr1
*==============*
Parameters schdat1 ,schbegin ,Schend ,schcont
Select (ttrfile)
m.varrec = Recno()
m.schobfile  = 'OB'+getext(m.schdat1)
m.schtrfile  = 'TR'+getext(m.schdat1)
=getselect(trfile,224,'serial')
=getselect(trfile,225,'empdayord')
Seek Str(EmpTab.serial)+Str(Dow(schdat1))
Do While !Eof() And emptab = EmpTab.serial And day = Dow(schdat1)
	If (&schbegin &schcont schend) 
		If m.tnbr != 0 And Empty(pubexit(m.tnbr)) And !Empty(pubentry(m.tnbr))
   			pubexit(tnbr)   = Iif(!Empty(exit)   ,ConfNight(exit),exit)
		Else	
			m.tnbr = m.tnbr + 1
			pubentry(tnbr)  = Iif(!Empty(entry)  ,ConfNight(entry),entry)
   			pubexit(tnbr)   = Iif(!Empty(exit)   ,ConfNight(exit),exit)
		EndIf

*		Pubserial(m.tnbr)	= Iif(m.pbegin > exit,pubserial(m.tnbr),serial)
*		Pubnserial(m.tnbr)	= Iif(m.pbegin > exit,serial,pubnserial(m.tnbr))

		Pubserial(m.tnbr)	= Iif(Dow(tdate) != day ,Iif(m.pconsnight,serial,0),Iif(m.pconsnight,0,serial))
		Pubnserial(m.tnbr)	= Iif(Dow(tdate) != day ,Iif(m.pconsnight,0,serial),Iif(m.pconsnight,serial,0))
  		
  		m.calwork = m.calwork + exit-entry

   		Select (schobfile)
   		Seek &schtrfile..serial 

   		pubobsent(tnbr) = obsent
   		pubobsext(tnbr) = obsext
   		pubhsent(tnbr)  = hsent
   		pubhsext(tnbr)  = hsext
   		pubdhsent(tnbr) = dhsent
   		pubdhsext(tnbr) = dhsext
   		pubcoment(tnbr) = coment
   		pubcomext(tnbr) = comext
	EndIf
	Select (schtrfile)
	Skip
EndDo
Return

Procedure SchTr
*=============*
Parameters schdat
*wait wind "SchTr" 
Select (ttrfile)
m.varrec = Recno()
m.schobfile  = 'OB'+getext(m.schdat)
m.schtrfile  = 'TR'+getext(m.schdat)
=getselect(trfile,224,'serial')
=getselect(trfile,225,'empdayord')
Seek Str(EmpTab.serial)+Str(Dow(schdat))
Do While !Eof() And emptab = EmpTab.serial And day = Dow(schdat)
	If m.tnbr != 0
		If Empty(pubexit(m.tnbr)) And !Empty(pubentry(m.tnbr))
			pubexit(m.tnbr) = Iif(!Empty(exit)   ,ConfNight(exit),exit)
		Else	
			m.tnbr = m.tnbr + 1
			pubentry(tnbr)  = Iif(!Empty(entry)  ,ConfNight(entry),entry)
   			pubexit(tnbr)   = Iif(!Empty(exit)   ,ConfNight(exit),exit)
		EndIf
	Else
		m.tnbr = m.tnbr + 1
		pubentry(tnbr)  = Iif(!Empty(entry)  ,ConfNight(entry),entry)
		pubexit(tnbr)   = Iif(!Empty(exit)   ,ConfNight(exit),exit)
	EndIf	
	
*	Pubserial(m.tnbr)	= Iif(Dow(tdate) != day ,0,serial)
*	Pubnserial(m.tnbr)	= Iif(Dow(tdate) != day ,serial,0)
	Pubserial(m.tnbr)	= Iif(Dow(tdate) != day ,Iif(m.pconsnight,serial,0),Iif(m.pconsnight,0,serial))
	Pubnserial(m.tnbr)	= Iif(Dow(tdate) != day ,Iif(m.pconsnight,0,serial),Iif(m.pconsnight,serial,0))

	m.calwork = m.calwork + (Iif(!Empty(entry)  ,SpecialCase(m.pbegin,entry) ,entry) - Iif(!Empty(exit)   ,SpecialCase(m.pbegin,exit)  ,exit))   &&&exit-entry
   	Select (schobfile)
   	Seek &schtrfile..serial       &&&pubserial(tnbr)
   	pubobsent(tnbr) = obsent
   	pubobsext(tnbr) = obsext
   	pubhsent(tnbr)  = hsent
   	pubhsext(tnbr)  = hsext
   	pubdhsent(tnbr) = dhsent
   	pubdhsext(tnbr) = dhsext
   	pubcoment(tnbr) = coment
   	pubcomext(tnbr) = comext
	Select (schtrfile)
	Skip
EndDo
Return
 
Procedure ConfNight
*=================*
Parameters conight
If m.pubnight 
	Do Case 
		Case !m.pconsnight
			If Dow(tdate) != day 
				Return SpecialCase(m.pbegin,conight)
			Endif
	Case m.pconsnight
		If Dow(tdate) = day 
			Return SpecialCase(m.pbegin,conight)
		Endif
EndCase
EndIf
Return m.conight

Procedure GetCheckBad
*===================*
Parameters baddat  
For gi = 1 To 30
	If Empty(pubentry(gi)) And Empty(pubexit(gi))
		Exit
	EndIf
	If Empty(pubentry) And !Empty(pubexit(gi))
		=ReceiveVar(baddat-1,"E",gi)
	EndIf	
	If !Empty(pubentry) And Empty(pubexit(gi))
		=ReceiveVar(baddat+1,"X",gi)
	EndIf	
EndFor

PROCEDURE getmax
****************
PARAMETERS tdbf,tfield

SELECT (tdbf)
torder = ORDER()
trecno = RECNO()
SET ORDER TO TAG (tfield)
GO BOTTOM
m.tvar = &tfield
SET ORDER TO TAG (torder)
GO RECORD trecno
RETURN m.tvar


PROCEDURE getclient
*******************
PARAMETERS tpclitab,tpcurtab

m.tptotal   = 0
m.tpbalance = 0
SELECT pcltab
tppclorder = ORDER()
SET ORDER TO clitab
SELECT pmotab
tppmoorder = ORDER()
SET ORDER TO mouhea
SELECT mouhea
tpmouorder = ORDER()
SET ORDER TO suptab

SELECT pcltab
SEEK tpclitab
DO WHILE clitab = tpclitab AND !EOF()
   m.tpbalance = m.tpbalance + amount
   SKIP
ENDDO
SELECT mouhea
SEEK tpclitab
DO WHILE suptab = tpclitab AND !EOF()
   IF supcli = 2
      m.tptotal = m.tptotal+cal_cur(mouhea.totprice,mouhea.date,mouhea.curtab,tpcurtab)
      SELECT pmotab
      SEEK mouhea.serial
      DO WHILE mouhea=mouhea.serial AND !EOF()
         m.tpbalance = m.tpbalance + cal_cur(amount,date,mouhea.curtab,tpcurtab)
         SELECT pmotab
         SKIP
      ENDDO
   ENDIF
   SELECT mouhea
   SKIP
ENDDO
m.tptotal   = INT(m.tptotal)
m.tpbalance = INT(m.tptotal)-INT(m.tpbalance)

PROCEDURE gettotcli
*******************
tpcurtab = 2
tpclitab = clitab.serial
m.tpbalance = 0
m.tptotal=0
SELECT pcltab
SEEK tpclitab
DO WHILE clitab = tpclitab AND !EOF()
   m.tpbalance = m.tpbalance + amount
   SKIP
ENDDO
**SELECT pmotab
**SET FILTER TO mouhea<=2377
SELECT mouhea
**SET FILTER TO serial<=2377
torder = ORDER()
trecno = RECNO()
SET ORDER TO suptab
SEEK tpclitab
DO WHILE suptab = tpclitab AND !EOF()
   IF supcli = 2
         m.tptotal = m.tptotal+cal_cur(mouhea.totprice,mouhea.date,mouhea.curtab,tpcurtab)
      SELECT pmotab
      SEEK mouhea.serial
      DO WHILE mouhea=mouhea.serial AND ! EOF()
         m.tpbalance = m.tpbalance + cal_cur(amount,date,mouhea.curtab,tpcurtab)
         SELECT pmotab
         SKIP
      ENDDO
   ENDIF
   SELECT mouhea
   SKIP
ENDDO
SELECT mouhea
SET ORDER TO torder
IF trecno > 0 AND trecno <= RECCOUNT()
	GO RECORD trecno
ENDIF
m.tptotal   = INT(m.tptotal)
m.tpbalance = INT(m.tptotal)-INT(m.tpbalance)

PROCEDURE cal_cur
*****************
PARAMETERS tpvalue,tpdate,tplcurtab,tpcurtab

IF tpcurtab=0 OR tplcurtab=0
	RETURN 0
ENDIF
IF tpcurtab=tplcurtab
   RETURN tpvalue
ENDIF
ttselect = SELECT()
SELECT rathea
LOCATE FOR date > tpdate
IF EOF()
   GO BOTTOM
ELSE   
   SKIP -1
ENDIF   
m.ttprathea = serial
SELECT ratrow
ttpporder = ORDER()
SET ORDER TO ratcur
SEEK STR(m.ttprathea)+STR(tplcurtab)
*m.ttpvalue = tpvalue*drate2
m.ttpvalue = tpvalue/rate2
SEEK STR(m.ttprathea)+STR(m.tpcurtab)
*m.ttpvalue = m.ttpvalue*rate2
m.ttpvalue = m.ttpvalue*rate2
SET ORDER TO (ttpporder)
SELECT (ttselect)
RETURN m.ttpvalue

PROCEDURE ordredt
*****************
m.date1=DATE()
m.date2=DATE()
m.nat=1
m.ok=1
DO selord.spr
alias='ordhea'
report1='ordredt1'
=getselect('ordhea',0,'0')
=getselect('arrhea',0,'ordhea')
=getselect('protab',0,'0')
=getselect('prttab',0,'0')
SET ORDER TO serial
=checkzapfile("prttab")
SELECT ordhea
GO TOP
DO WHILE !EOF()	
	xval=.T.
	IF date >= m.date1 AND date <= m.date2 
		IF nat = 2
			xval=.F.
			SELECT arrhea
			SET ORDER TO ordhea
			SEEK ordhea.serial
			IF !FOUND()
				xval=.T.
			ENDIF	
		ENDIF	
			IF xval 
				SET ORDER TO 
				SELECT prttab
				APPEND BLANK
				REPLACE serial WITH ordhea.serial
				REPLACE name WITH protab.name
				REPLACE number WITH ALLTRIM(DTOC(ordhea.date))
			ENDIF	
	ENDIF		
	SELECT ordhea
	SKIP
ENDDO	
SELECT prttab
GO TOP
IF !EOF()
	IF LASTKEY() <> 27  AND m.ok = 1
		DO printall.spr
	ENDIF	
ELSE
    WAIT WINDOW "EMPTY FILE" NOWAIT        
ENDIF	

PROCEDURE arrredt
*****************
m.date1=DATE()
m.date2=DATE()
m.nat=1
m.ok=1
DO selarr.spr
alias='arrhea'
report1='arrredt1'
=getselect('arrhea',0,'0')
=getselect('protab',0,'0')
=getselect('prttab',0,'0')
SET ORDER TO serial
=checkzapfile("prttab")
SELECT arrhea
GO TOP
DO WHILE !EOF()	
	IF date >= m.date1 AND date <= m.date2 
		SET ORDER TO 
		SELECT prttab
		APPEND BLANK
		REPLACE serial WITH arrhea.serial
		REPLACE name WITH protab.name
		REPLACE number WITH ALLTRIM(DTOC(arrhea.date))
   	ENDIF	
	SELECT arrhea
	SKIP
ENDDO	
SELECT prttab
GO TOP
IF !EOF()
	IF LASTKEY() <> 27  AND m.ok = 1
		DO printall.spr
	ENDIF	
ELSE
    WAIT WINDOW "EMPTY FILE" NOWAIT        
ENDIF	

PROCEDURE salredt
*****************
m.date1=DATE()
m.date2=DATE()
DO criteria.spr

PROCEDURE cliredt
******************
m.nat=1
m.ok=1
DO selcust.spr
alias='clitab'
report1='cliredt1'
=getselect('clitab',0,'0')
=getselect('prttab',0,'serial')
=checkzapfile("prttab")
DO CASE
	CASE m.nat = 1
		SELECT prttab
		APPEND FROM clitab
	CASE m.nat = 2
		SELECT clitab
		GO TOP
		DO WHILE !EOF()
			IF SIGN(balance) = 1
				SCATTER MEMVAR
				SELECT prttab
				APPEND BLANK
				GATHER MEMVAR
			ENDIF
			SELECT clitab
			SKIP
		ENDDO
	CASE m.nat = 3
		SELECT clitab
		GO TOP
		DO WHILE !EOF()
			IF SIGN(balance) = -1
				SCATTER MEMVAR
				SELECT prttab
				APPEND BLANK
				GATHER MEMVAR
			ENDIF
			SELECT clitab
			SKIP
		ENDDO
	CASE m.nat = 4
		SELECT clitab
		GO TOP
		DO WHILE !EOF()
			IF clitab.tot_set > clitab.tot_inv
				SCATTER MEMVAR
				SELECT prttab
				APPEND BLANK
				GATHER MEMVAR
			ENDIF
			SELECT clitab
			SKIP
		ENDDO
ENDCASE		
SELECT prttab
GO TOP
IF !EOF()
	IF LASTKEY() <> 27  AND m.ok = 1
		DO printall.spr
	ENDIF	
ELSE
    WAIT WINDOW "EMPTY FILE" NOWAIT        
ENDIF	
			
PROCEDURE warredt
*****************
m.nat=1
m.ok=1
DO selwea.spr
alias='warhea'
report1='warredt'
=getselect('warrow',0,'warhea')
=getselect('protab',0,'serial')
=getselect('warhea',0,'0')
SET RELATION TO Warhea.serial INTO Warrow ADDITIVE
SELECT warrow
SET RELATION TO Warrow.protab INTO Protab ADDITIVE
=getselect('prttab',0,'serial')
=checkzapfile("prttab")
DO CASE
	CASE m.nat = 1
		SELECT prttab
		APPEND FROM warhea
	CASE m.nat = 2	
		SELECT warhea
		REPLACE ALL bidon WITH ""
		REPLACE BIDON WITH "X" FOR FOUND("WARROW") AND PROTAB.QTY < PROTAB.QTYMIN
		GO TOP
		DO WHILE !EOF()
			IF !EMPTY(bidon)
				SCATTER MEMVAR
				SELECT prttab
				APPEND BLANK
				GATHER MEMVAR
			ENDIF
			SELECT warhea
			SKIP
		ENDDO
	CASE m.nat = 3
		SELECT warhea
		REPLACE ALL bidon WITH ""
		REPLACE BIDON WITH "X" FOR FOUND("WARROW") AND PROTAB.QTY > PROTAB.QTYMAX
		GO TOP
		DO WHILE !EOF()
			IF !EMPTY(bidon)
				SCATTER MEMVAR
				SELECT prttab
				APPEND BLANK
				GATHER MEMVAR
			ENDIF
			SELECT warhea
			SKIP
		ENDDO
ENDCASE				
SELECT prttab
GO TOP
IF !EOF()
	IF LASTKEY() <> 27  AND m.ok = 1
		DO printall.spr
	ENDIF	
ELSE
    WAIT WINDOW "EMPTY FILE" NOWAIT        
ENDIF	
		
		
PROCEDURE invredt
*****************
m.date1=DATE()
m.date2=DATE()
m.nat=1
m.ok=1
DO selarr.spr
alias='invhea'
report1='invredt1'
=getselect('invhea',0,'0')
=getselect('protab',0,'0')
=getselect('prttab',0,'0')
SET ORDER TO serial
=checkzapfile("prttab")
SELECT invhea
GO TOP
DO WHILE !EOF()	
	IF date >= m.date1 AND date <= m.date2 
		SET ORDER TO 
		SELECT prttab
		APPEND BLANK
		REPLACE serial WITH invhea.serial
		REPLACE name WITH protab.name
		REPLACE number WITH ALLTRIM(DTOC(invhea.date))
   	ENDIF	
	SELECT invhea
	SKIP
ENDDO	
SELECT prttab
GO TOP
IF !EOF()
	IF LASTKEY() <> 27  AND m.ok = 1
		DO printall.spr
	ENDIF	
ELSE
    WAIT WINDOW "EMPTY FILE" NOWAIT        
ENDIF	

PROCEDURE print
***************
PARAMETERS tname,tcondition,ttype
REPORT FORMAT &tname ENVIRONMENT FOR &tcondition &ttype
CLOS DATA

PROCEDURE newprint
******************
PARAMETERS tname,tcondition,ttype

REPORT FORMAT &tname FOR &tcondition &ttype

PROCEDURE opendde
*****************
=DDESetService('myservice','DEFINE')
=DDESetService('myservice','ADVISE',.T.)
=DDESetService('myservice','EXECUTE',.T.)
=DDESetService('myservice','POKE',.T.)
=DDESetService('myservice','REQUEST',.T.)

PROCEDURE opencom
*****************
=DDESetTopic('myservice','mycomtopic','comprog')
comchannel = DDEInitiate('myservice','mycomtopic')
*=DDEExecute(dchannel, 'WAIT WINDOW "Command Executed ... "')
*=DDETerminate(dchannel)

PROCEDURE sendcom
*****************
PARAMETERS tdata
DO wait WITH 4
= DDEPOKE(comchannel,'mycomitem',tdata,'CF_TEXT')

PROCEDURE wait 
**************
PARAMETER waiti

FOR ti=1 TO waiti
    FOR ji=-30000 TO 30000
    ENDFOR
ENDFOR

PROCEDURE gettotbal
*******************
CLOS DATA
m.tpbalance = 0
m.tptotal   = 0
SELECT 0
USE  rathea
SELECT 0
USE  ratrow
SELECT 0
USE  pcltab
tppclorder = ORDER()
SET ORDER TO clitab
SELECT 0
USE pmotab
tppmoorder = ORDER()
SET ORDER TO mouhea
SELECT 0
USE mouhea
tpmouorder = ORDER()
SET ORDER TO suptab
SELECT 0
USE clitab
GO TOP
DO WHILE !EOF()
	DO gettotcli 
	SELECT clitab
	REPLACE balance WITH m.tpbalance
	REPLACE tot_pur WITH m.tptotal
	SKIP
ENDDO	
SELECT clitab
m.out=1
DO printrec.spr
SELECT clitab
IF m.out=1
	REPORT FORM creredt FOR !EMPTY(balance) PREVIEW 
ELSE
	REPORT FORM creredt FOR !EMPTY(balance) TO PRINT OFF
ENDIF	
CLOS DATA
RETURN

PROCEDURE getprolst
*******************
PARAMETER mvalx

CLOSE DATA
SELECT 0
USE protab
m.ok = 1
DO warspc.spr
IF m.ok=2
	RETURN
ENDIF
CLOS DATA
m.qtyexit  = 0
m.arrprice = 0
m.out      = 1
m.okbtn    = 1
DO printrec.spr
IF m.okbtn=2
	RETURN
ENDIF
SELECT 0
USE mourow
SELECT 0
USE rathea
SELECT 0
USE ratrow
SELECT 0
USE arrrow
SET ORDER TO protab
SELECT 0
USE arrhea
SET ORDER TO serial
SELECT 0
USE protab
SET FILTER TO !EMPTY(bidon)
SELECT 0
USE famtab1
SET ORDER TO serial
SELECT 0
USE fl2tab
SET ORDER TO serial
SELECT 0
USE fl3tab
SET ORDER TO serial
SELECT 0
USE fl4tab
SET ORDER TO serial
SELECT 0
USE fl1tab
SET ORDER TO serial
SELECT protab
SET RELATION TO Protab.famtab1 INTO Famtab1 ADDITIVE
SET RELATION TO Protab.fl1tab  INTO Fl1tab ADDITIVE
SET RELATION TO Protab.fl2tab  INTO Fl2tab ADDITIVE
SET RELATION TO Protab.fl3tab  INTO Fl3tab ADDITIVE
SET RELATION TO Protab.fl4tab  INTO Fl4tab ADDITIVE
SELECT protab
SET ORDER TO newproray
GO TOP
IF mvalx=1
	IF m.out=1
		REPORT FORM proredt3 PREVIEW 
	ELSE
		REPORT FORM proredt3 TO PRINT OFF
	ENDIF	
ELSE
	IF m.out=1
		REPORT FORM proredt4 PREVIEW 
	ELSE
		REPORT FORM proredt4 TO PRINT OFF
	ENDIF	
ENDIF
CLOSE DATA

PROCEDURE getsel
****************
SELECT arrrow
SET ORDER TO protab
SEEK protab.serial
IF FOUND()
	SELECT arrhea
	SET ORDER TO serial
	SEEK arrrow.arrhea
	IF FOUND()
**		m.arrprice = ROUND(cal_cur(arrrow.price/arrrow.qty,arrhea.date,arrhea.curtab,1),-3)
		m.arrprice = ROUND(cal_cur(arrrow.price/arrrow.qty,arrhea.date,arrhea.curtab,1),0)
		SELECT protab
		RETURN m.arrprice
	ENDIF	
ENDIF	
SELECT protab
RETURN 0


PROCEDURE getspcwar
*******************
CLOS DATA
SELECT 0
USE protab
m.ok = 1
DO warspc.spr
IF m.ok=2
	RETURN
ENDIF	
CLOS DATA
m.qtyexit = 0
m.qtyini  = 0
m.tot     = 0
m.out     = 1
m.okbtn   = 1
DO printrec.spr
IF m.okbtn=2
	RETURN
ENDIF
SELECT 0
USE mourow
SELECT 0
USE arrrow
SELECT 0
USE protab
SET FILTER TO !EMPTY(bidon)
SELECT 0
USE famtab1
SET ORDER TO serial
SELECT 0
USE fl2tab
SET ORDER TO serial
SELECT 0
USE fl1tab
SET ORDER TO serial
SELECT protab
SET RELATION TO Protab.famtab1 INTO Famtab1 ADDITIVE
SET RELATION TO Protab.fl1tab INTO Fl1tab ADDITIVE
SET RELATION TO Protab.fl2tab INTO Fl2tab ADDITIVE
SELECT protab
SET ORDER TO proray
GO TOP
IF m.out=1
	REPORT FORM proredt6 PREVIEW 
ELSE
	REPORT FORM proredt6 TO PRINT OFF
ENDIF	
CLOS DATA

PROCEDURE getspcpro
******************
CLOS DATA
m.date = {//}
SELECT 0
USE protab
SELECT 0
USE arrhea
SELECT 0
USE arrrow
m.ok = 1
DO proredt.spr
IF m.ok=2
	RETURN
ENDIF
CLOS DATA
m.qtyexit = 0
m.out     = 1
m.okbtn   = 1
DO printrec.spr
IF m.okbtn=2
	RETURN
ENDIF
SELECT 0
USE mourow
SELECT 0
USE arrrow
SELECT 0
USE protab
SET FILTER TO !EMPTY(bidon)
SELECT 0
USE famtab1
SET ORDER TO serial
SELECT 0
USE fl2tab
SET ORDER TO serial
SELECT 0
USE fl1tab
SET ORDER TO serial
SELECT 0
USE fl3tab
SET ORDER TO serial
SELECT 0
USE fl4tab
SET ORDER TO serial
SELECT 0
USE rathea
SELECT 0
USE ratrow
SELECT protab
SET RELATION TO Protab.famtab1 INTO Famtab1 ADDITIVE
SET RELATION TO Protab.fl1tab  INTO Fl1tab ADDITIVE
SET RELATION TO Protab.fl2tab  INTO Fl2tab ADDITIVE
SET RELATION TO Protab.fl3tab  INTO Fl3tab ADDITIVE
SET RELATION TO Protab.fl4tab  INTO Fl4tab ADDITIVE
SELECT protab
SET ORDER TO newproray
GO TOP
IF m.out=1
	REPORT FORM proredt2 PREVIEW 
ELSE
	REPORT FORM proredt2 TO PRINT OFF
ENDIF	
CLOS DATA

PROCEDURE gettotwar
*******************
CLOS DATA
m.qtyexit = 0
m.qtyini  = 0
m.tot     = 0
m.out     = 1
m.okbtn   = 1
DO printrec.spr
IF m.okbtn=2
	RETURN
ENDIF
SELECT 0
USE arrrow
SELECT 0
USE mourow
SELECT 0
USE protab
SELECT 0
USE famtab1
SET ORDER TO serial
SELECT 0
USE fl2tab
SET ORDER TO serial
SELECT 0
USE fl1tab
SET ORDER TO serial
SELECT protab
SET RELATION TO Protab.famtab1 INTO Famtab1 ADDITIVE
SET RELATION TO Protab.fl1tab INTO Fl1tab ADDITIVE
SET RELATION TO Protab.fl2tab INTO Fl2tab ADDITIVE
SELECT protab
SET ORDER TO proray
GO TOP
IF m.out=1
	REPORT FORM proredt PREVIEW 
ELSE
	REPORT FORM proredt TO PRINT OFF
ENDIF	
CLOS DATA
RETURN


PROCEDURE texitqty
******************
m.qtyexit = 0
m.protab  = protab.serial
SELECT mourow
SET ORDER TO protab
SEEK m.protab
DO WHILE !EOF() AND mourow.protab = m.protab
*	IF !mouhea.return
	IF EMPTY(mourow.proret) AND !mouhea.return
		m.qtyexit = m.qtyexit + (mourow.qty+mourow.free)
	ELSE
		m.qtyexit = m.qtyexit - mourow.qty
	ENDIF			
	SELECT mourow
	SKIP
ENDDO
SELECT protab
m.pri = (m.qtyini-m.qtyexit)*protab.selling
m.tot = m.tot+m.pri
**RETURN IIF(m.qtyexit=0,"-",m.qtyexit)
RETURN m.qtyexit

PROCEDURE tiniqty
*****************
*m.qtyini = 0
m.qtyini = protab.adjqty
m.protab = protab.serial
SELECT arrrow
SET ORDER TO protab
SEEK m.protab
DO WHILE !EOF() AND arrrow.protab = m.protab
	m.qtyini = m.qtyini + arrrow.qty
	SELECT arrrow
	SKIP
ENDDO
SELECT protab
RETURN IIF(m.qtyini=0,"-",m.qtyini)

PROCEDURE getuprice
*******************
IF m.curtab=0 OR protab.curtab=0
	RETURN 0
ENDIF
IF m.curtab = protab.curtab
   m.suprice = protab.selling
   SELECT protab
   RETURN
ENDIF
**IF m.curtab = 1
**   m.suprice = protab.selling
**   SELECT protab
**  RETURN
**ENDIF
SELECT rathea
LOCATE FOR date>m.date
IF EOF()
   GO BOTTOM
ELSE
   SKIP -1
ENDIF   
m.rathea = serial
SELECT ratrow
SEEK STR(m.rathea)+STR(protab.curtab)
*m.suprice = protab.selling*drate2
m.suprice = protab.selling/rate2
SEEK STR(m.rathea)+STR(m.curtab)
*m.suprice = m.suprice*rate2
m.suprice = m.suprice*rate2
IF m.curtab<>1
	m.suprice = ROUND(m.suprice,2)
ENDIF	
SELECT protab

PROCEDURE checkzapfile
**********************
PARAMETERS tzapalias

SELECT (tzapalias)
IF UPPER(ALIAS()) = UPPER(tzapalias)
   ZAP
ELSE
   WAIT WINDOW 'File '+tzapalias+' Is corrupted' NOWAIT
ENDIF

PROCEDURE calpayed
******************
m.totpayed = 0
m.totpayed1=0 
m.totpayed2=0 
SELECT pmotab
SEEK m.serial
DO WHILE mouhea = m.serial AND !EOF()
	IF m.curtab=4
		m.totpayed = m.totpayed + ROUND(cal_cur(amount,date,curtab,m.curtab),-3)
	ELSE
		m.totpayed = m.totpayed + ROUND(cal_cur(amount,date,curtab,m.curtab),0)
	ENDIF
	SELECT pmotab
**	m.totpayed1 = m.totpayed1 + ROUND(cal_cur(amount,date,curtab,1),-3)
	m.totpayed1 = m.totpayed1 + ROUND(cal_cur(amount,date,curtab,1),0)
	SELECT pmotab
	m.totpayed2 = m.totpayed2 + ROUND(cal_cur(amount,date,curtab,2),0)
	SELECT pmotab
	SKIP
ENDDO
SELECT (alias)

PROCEDURE getitmpri
*******************
PARAMETERS trep

IF LASTKEY()<>27
	CLOS DATA
	m.qtyexit = 0
	m.qtyini  = 0
	m.tot     = 0
	m.tot1    = 0	
	m.pric    = 0
**	m.out     = 1
	m.frdate  = {//}
	m.todate  = DATE()
	m.okbtn   = 1
	m.mouser  = 0
	m.arrser  = 0
	m.mfrser  = 0
	m.afrser  = 0
	m.sea     = .F.
	m.famtab  = 0
	m.famdes  = ''
	SELECT 0
	USE famtab
**	DO printrec.spr
	DO ivtredt1.spr
	IF m.okbtn=4
		RETURN
	ENDIF
	IF !m.sea
		m.famtab = 0
	ENDIF	
	SELECT 0
	USE invtab
	= checkzapfile('invtab')
		SELECT 0
	USE mouhea
	IF EMPTY(m.frdate)
		SET FILTER TO date<=m.todate
	ELSE
		SET FILTER TO date>=m.frdate AND date<=m.todate
		GO TOP
		m.mfrser = serial
	ENDIF	
	GO BOTTOM
	m.mouser = serial
	SET ORDER TO serial
	GO TOP
	SELECT 0
	USE mourow
	IF EMPTY(m.frdate)
		SET FILTER TO mouhea<=m.mouser
	ELSE
		SET FILTER TO mouhea>=m.mfrser AND mouhea<=m.mouser
	ENDIF	
	SELECT 0		
	USE arrhea
	IF EMPTY(m.frdate)
		SET FILTER TO date<=m.todate
	ELSE
		SET FILTER TO date>=m.frdate AND date<=m.todate
		GO TOP
		m.afrser = serial
	ENDIF	
	GO BOTTOM
	m.arrser = serial
	SET ORDER TO serial
	GO TOP
	SELECT 0
	USE arrrow
	IF EMPTY(m.frdate)
		SET FILTER TO arrhea<=m.arrser
	ELSE
		SET FILTER TO arrhea>=m.afrser AND arrhea<=m.arrser
	ENDIF	
	SELECT 0
	USE protab
	SELECT 0
	USE famtab1
	SET ORDER TO serial
	SELECT 0
	USE fl2tab
	SET ORDER TO serial
	SELECT 0
	USE fl1tab
	SET ORDER TO serial
	SELECT 0
	USE fl4tab
	SET ORDER TO serial	
	SELECT 0
	USE rathea
	SELECT 0		
	USE ratrow		
	SELECT protab
	SET RELATION TO Protab.famtab1 INTO Famtab1 ADDITIVE
	SET RELATION TO Protab.fl1tab  INTO Fl1tab ADDITIVE
	SET RELATION TO Protab.fl2tab  INTO Fl2tab ADDITIVE
	SELECT mourow
	SET RELATION TO Mourow.mouhea  INTO Mouhea ADDITIVE
	SELECT arrrow
	SET RELATION TO Arrrow.arrhea  INTO Arrhea ADDITIVE		
	SELECT protab
	SET ORDER TO proray
	GO TOP
	IF trep=1
		= getinvsel()
		SELECT invtab
		SET RELATION TO Invtab.famtab1 INTO Famtab1 ADDITIVE
		SET RELATION TO Invtab.fl1tab  INTO Fl1tab ADDITIVE
		SET RELATION TO Invtab.fl2tab  INTO Fl2tab ADDITIVE
		SET RELATION TO Invtab.fl4tab  INTO Fl4tab ADDITIVE
		IF m.okbtn=1
			REPORT FORM proredt1 PREVIEW 
		ELSE
			IF m.okbtn=2
				REPORT FORM proredt1 TO PRINT OFF
			ELSE
				DO invtxls WITH 'Selling',1
			ENDIF	
		ENDIF	
	ELSE
		IF trep=2
			= getcurcost()
		ELSE
			= getinvcost()
		ENDIF	
		SELECT invtab
		SET ORDER TO fmname
		SET RELATION TO Invtab.famtab1 INTO Famtab1 ADDITIVE
		SET RELATION TO Invtab.fl1tab  INTO Fl1tab ADDITIVE
		SET RELATION TO Invtab.fl2tab  INTO Fl2tab ADDITIVE	
		SET RELATION TO Invtab.fl4tab  INTO Fl4tab ADDITIVE		
		IF m.okbtn=1
			REPORT FORM proredt7 PREVIEW 
		ELSE
			IF m.okbtn=2
				REPORT FORM proredt7 TO PRINT OFF
			ELSE
				DO invtxls WITH 'Cost',trep-1
			ENDIF	
		ENDIF			
	ENDIF	
ENDIF
CLOS DATA

PROCEDURE 11getitmpri
*******************
PARAMETERS trep

IF LASTKEY()<>27
	CLOS DATA
	m.qtyexit = 0
	m.qtyini  = 0
	m.tot     = 0
	m.tot1    = 0	
	m.pric    = 0
**	m.out     = 1
	m.frdate  = {//}
	m.todate  = DATE()
	m.okbtn   = 1
	m.mouser  = 0
	m.arrser  = 0
	m.mfrser  = 0
	m.afrser  = 0
**	DO printrec.spr
	DO ivtredt1.spr
	IF m.okbtn=3
		RETURN
	ENDIF
	SELECT 0
	USE invtab
	= checkzapfile('invtab')
	SELECT 0
	USE mouhea
	IF EMPTY(m.frdate)
		SET FILTER TO date<=m.todate
	ELSE
		SET FILTER TO date>=m.frdate AND date<=m.todate
		GO TOP
		m.mfrser = serial
	ENDIF	
	GO BOTTOM
	m.mouser = serial
	SET ORDER TO serial
	GO TOP
	SELECT 0	
	USE mourow
	IF EMPTY(m.frdate)
		SET FILTER TO mouhea<=m.mouser
	ELSE
		SET FILTER TO mouhea>=m.mfrser AND mouhea<=m.mouser
	ENDIF	
	SELECT 0		
	USE arrhea
	IF EMPTY(m.frdate)
		SET FILTER TO date<=m.todate
	ELSE
		SET FILTER TO date>=m.frdate AND date<=m.todate
		GO TOP
		m.afrser = serial
	ENDIF	
	GO BOTTOM
	m.arrser = serial
	SET ORDER TO serial
	GO TOP
	SELECT 0
	USE arrrow
	IF EMPTY(m.frdate)
		SET FILTER TO arrhea<=m.arrser
	ELSE
		SET FILTER TO arrhea>=m.afrser AND arrhea<=m.arrser
	ENDIF	
	SELECT 0
	USE protab
	SELECT 0
	USE famtab1
	SET ORDER TO serial
	SELECT 0
	USE fl2tab
	SET ORDER TO serial
	SELECT 0
	USE fl1tab
	SET ORDER TO serial
	SELECT 0		
	USE rathea
	SELECT 0		
	USE ratrow		
	SELECT protab
	SET RELATION TO Protab.famtab1 INTO Famtab1 ADDITIVE
	SET RELATION TO Protab.fl1tab  INTO Fl1tab ADDITIVE
	SET RELATION TO Protab.fl2tab  INTO Fl2tab ADDITIVE
	SELECT mourow
	SET RELATION TO Mourow.mouhea  INTO Mouhea ADDITIVE
	SELECT arrrow
	SET RELATION TO Arrrow.arrhea  INTO Arrhea ADDITIVE		
	SELECT protab
	SET ORDER TO proray
	GO TOP
	IF trep=1
		= getinvsel()
		SELECT invtab
		SET RELATION TO Invtab.famtab1 INTO Famtab1 ADDITIVE
		SET RELATION TO Invtab.fl1tab  INTO Fl1tab ADDITIVE
		SET RELATION TO Invtab.fl2tab  INTO Fl2tab ADDITIVE
		IF m.okbtn=1
			REPORT FORM proredt1 PREVIEW 
		ELSE
			REPORT FORM proredt1 TO PRINT OFF
		ENDIF	
	ELSE
		IF trep=2
			= getcurcost()
		ELSE
			= getinvcost()
		ENDIF	
		SELECT invtab
		SET ORDER TO fmname
		SET RELATION TO Invtab.famtab1 INTO Famtab1 ADDITIVE
		SET RELATION TO Invtab.fl1tab  INTO Fl1tab ADDITIVE
		SET RELATION TO Invtab.fl2tab  INTO Fl2tab ADDITIVE
		IF m.okbtn=1
			REPORT FORM proredt7 PREVIEW
		ELSE
			REPORT FORM proredt7 TO PRINT OFF
		ENDIF			
	ENDIF	
ENDIF
CLOS DATA

PROCEDURE tiniqty1
******************
*m.qtyini = 0
m.qtyini = protab.adjqty
m.protab = protab.serial
SELECT arrrow
SET ORDER TO protab
SEEK m.protab
DO WHILE !EOF() AND arrrow.protab = m.protab
	m.qtyini = m.qtyini + arrrow.qty
	SELECT arrrow
	SKIP
ENDDO
SELECT protab
RETURN m.qtyini

PROCEDURE getcostpr
*******************
PARAMETERS tserial

tselect = SELECT()
SELECT arrrow
SUM cal_cur(price,arrhea.date,arrhea.curtab,3) to tprice FOR protab=tserial
SUM qty to tqty FOR protab=tserial
m.pric = (m.qtyini-m.qtyexit)*IIF(tqty<>0,(tprice/tqty),tprice)
m.tot1 = m.tot1+m.pric
SELECT (tselect)
*RETURN IIF(tqty<>0,(tprice/tqty),'-')
RETURN IIF(tqty<>0,(tprice/tqty),0)

PROCEDURE p
***********
sele 0
USE c:\stock\temp.dbf EXCLUSIVE
SELE 0
USE rathea
SELE 0
USE ratrow
SELECT protab
GO TOP
DO WHILE !EOF()
	SCATTER MEMVAR
	SELECT temp
	APPEND BLANK
	REPLACE name   WITH m.name
	REPLACE number WITH m.number
	REPLACE ini    WITH tiniqty1()
	SELECT temp
	REPLACE qexit  WITH texitqty()
	SELECT temp
	REPLACE qty    WITH m.qtyini-m.qtyexit
	REPLACE price  WITH m.selling &&getcostpr(m.serial)
	SELECT temp
	REPLACE total  WITH (m.qtyini-m.qtyexit)*m.selling &&m.pric
	SELECT protab
	SKIP
ENDDO

PROCEDURE prtspecial
********************
IF !USED('clitab')
	SELECT 0
	USE clitab
ELSE
	SELECT clitab	
ENDIF
IF !USED('mouhea')
	SELECT 0
	USE mouhea
ELSE
	SELECT mouhea
ENDIF
IF !USED('rathea')
	SELECT 0
	USE rathea
ELSE
	SELECT rathea
ENDIF
IF !USED('ratrow')
	SELECT 0
	USE ratrow
ELSE
	SELECT ratrow
ENDIF
SELECT clitab
GO TOP
DO WHILE !EOF()
	m.serial = clitab.serial
	SELECT mouhea
	SUM cal_cur(totprice,date,curtab,2) TO n FOR suptab=m.serial
	SELECT CLITAB
	REPLACE bal WITH n
	SKIP
ENDDO
m.out   = 1
m.okbtn = 1
DO printrec.spr
IF m.okbtn=2
	CLOSE DATA
	RETURN
ENDIF	
SELECT clitab
SET FILTER TO bal>=300 &&AND !card
SET ORDER TO name
IF m.out=1
	REPORT FORM cliredt2 PREVIEW
ELSE
	REPORT FORM cliredt2 TO PRINT OFF
ENDIF
CLOSE DATA

PROCEDURE getinvsel
*******************
SELECT protab
SET ORDER TO proray
IF m.famtab<>0
	SET FILTER TO famtab=m.famtab
ENDIF
GO TOP
DO WHILE !EOF()
	IF selling<>0 AND tiniqty2(serial,adjqty)<>0 AND m.qtyini-texitqty1()>0
		SELECT protab
		SCATTER MEMVAR
		SELECT invtab
		APPEND BLANK
		GATHER MEMVAR
		REPLACE qtyini  WITH m.qtyini
		REPLACE qtyexit WITH m.qtyexit
		REPLACE invqty  WITH m.qtyini-m.qtyexit
		REPLACE totsel  WITH invqty*selling
	ENDIF	
	SELECT protab
	SKIP
ENDDO

PROCEDURE getcurcost
********************
SELECT protab
SET ORDER TO proray
IF m.famtab<>0
	SET FILTER TO famtab=m.famtab
ENDIF
GO TOP
DO WHILE !EOF()
	IF selling<>0 AND tiniqty2(serial,adjqty)<>0 AND m.qtyini-texitqty1()<>0
		SELECT protab
		SCATTER MEMVAR
		SELECT invtab
		APPEND BLANK
		GATHER MEMVAR
		REPLACE qtyini  WITH m.qtyini
		REPLACE qtyexit WITH m.qtyexit
		REPLACE invqty  WITH m.qtyini-m.qtyexit
		REPLACE costff  WITH getcostpr(m.serial)
		REPLACE totsel  WITH invqty*costff
		REPLACE totini  WITH qtyini*costff
	ENDIF	
	SELECT protab
	SKIP
ENDDO

PROCEDURE getinvcost
********************
SELECT protab
SET ORDER TO proray
IF m.famtab<>0
	SET FILTER TO famtab=m.famtab
ENDIF
GO TOP
DO WHILE !EOF()
	IF selling<>0 AND tiniqty2(serial,adjqty)<>0
		SELECT protab
		SCATTER MEMVAR
		SELECT invtab
		APPEND BLANK
		GATHER MEMVAR
		REPLACE qtyini  WITH m.qtyini
		REPLACE qtyexit WITH texitqty1()
		SELECT invtab
		REPLACE invqty  WITH m.qtyini-m.qtyexit
		REPLACE costff  WITH getcostpr(m.serial)
		REPLACE totsel  WITH invqty*costff
		REPLACE totini  WITH qtyini*costff
	ENDIF	
	SELECT protab
	SKIP
ENDDO

PROCEDURE texitqty1
*******************
m.qtyexit = 0
m.protab  = protab.serial
SELECT mourow
SET ORDER TO protab
SEEK m.protab
DO WHILE !EOF() AND mourow.protab=m.protab
	IF EMPTY(mourow.proret) AND !mouhea.return
		m.qtyexit = m.qtyexit + (mourow.qty+mourow.free)
	ELSE
		m.qtyexit = m.qtyexit - mourow.qty
	ENDIF			
	SELECT mourow
	SKIP
ENDDO
SELECT protab
RETURN m.qtyexit

PROCEDURE tiniqty2
******************
PARAMETERS pserial,padjqty

m.qtyini = 0
SELECT arrrow
SUM qty TO m.qtyini FOR protab=pserial
m.qtyini = m.qtyini+padjqty
SELECT protab
RETURN m.qtyini

PROCEDURE tgg
*************
PARAMETERS m.price

*fifty = m.price+(m.price*0.5)
fifty = m.price+(m.price*0.4)
eight = fifty+(fifty*0.8)
m.selling = ROUND(((((m.price*0.1)+eight)/6))*1530,-3)
? m.selling

PROCEDURE chgdate
*****************
PARAMETERS vdate

RETURN PADL(DAY(vdate),2,'0')+' '+CMONTH(vdate)+' '+STR(YEAR(vdate),4)

PROCEDURE getmodlst
*******************
CLOSE DATA
m.qtyexit  = 0
m.arrprice = 0
m.out      = 1
m.okbtn    = 1
DO printrec.spr
IF m.okbtn=2
	RETURN
ENDIF
SELECT 0
USE mourow
SELECT 0
USE rathea
SELECT 0
USE ratrow
SELECT 0
USE arrrow
SET ORDER TO protab
SELECT 0
USE arrhea
SET ORDER TO serial
SELECT 0
USE protab
SET FILTER TO !EMPTY(qte)
SELECT 0
USE famtab1
SET ORDER TO serial
SELECT 0
USE fl2tab
SET ORDER TO serial
SELECT 0
USE fl3tab
SET ORDER TO serial
SELECT 0
USE fl4tab
SET ORDER TO serial
SELECT 0
USE fl1tab
SET ORDER TO serial
SELECT protab
SET RELATION TO Protab.famtab1 INTO Famtab1 ADDITIVE
SET RELATION TO Protab.fl1tab  INTO Fl1tab ADDITIVE
SET RELATION TO Protab.fl2tab  INTO Fl2tab ADDITIVE
SET RELATION TO Protab.fl3tab  INTO Fl3tab ADDITIVE
SET RELATION TO Protab.fl4tab  INTO Fl4tab ADDITIVE
SELECT protab
SET ORDER TO mod
GO TOP
IF m.out=1
	REPORT FORM proredt8 PREVIEW 
ELSE
	REPORT FORM proredt8 TO PRINT OFF
ENDIF	
CLOSE DATA


PROCEDURE getinvspr
*******************
IF LASTKEY()<>27
	CLOS DATA
	m.qtyexit = 0
	m.qtyini  = 0
*	m.tot     = 0
*	m.tot1    = 0	
*	m.pric    = 0
	m.frdate  = {//}
	m.todate  = DATE()
	m.okbtn   = 1
	m.mouser  = 0
	m.arrser  = 0
	m.mfrser  = 0
	m.afrser  = 0
	m.sea     = .F.
	m.famtab  = 0
	m.famdes  = ''
	SELECT 0
	USE famtab
	DO ivtredt1.spr
	IF m.okbtn=3
		WAIT WINDOW 'Feature not available' NOWAIT
		RETURN
	ENDIF
	IF m.okbtn=4
		RETURN
	ENDIF
	IF !m.sea
		m.famtab = 0
	ENDIF	
	SELECT 0
	USE invtab
	= checkzapfile('invtab')
		SELECT 0
	USE mouhea
	IF EMPTY(m.frdate)
		SET FILTER TO date<=m.todate
	ELSE
		SET FILTER TO date>=m.frdate AND date<=m.todate
		GO TOP
		m.mfrser = serial
	ENDIF	
	GO BOTTOM
	m.mouser = serial
	SET ORDER TO serial
	GO TOP
	SELECT 0
	USE mourow
	IF EMPTY(m.frdate)
		SET FILTER TO mouhea<=m.mouser
	ELSE
		SET FILTER TO mouhea>=m.mfrser AND mouhea<=m.mouser
	ENDIF	
	SELECT 0		
	USE arrhea
	IF EMPTY(m.frdate)
		SET FILTER TO date<=m.todate
	ELSE
		SET FILTER TO date>=m.frdate AND date<=m.todate
		GO TOP
		m.afrser = serial
	ENDIF	
	GO BOTTOM
	m.arrser = serial
	SET ORDER TO serial
	GO TOP
	SELECT 0
	USE arrrow
	IF EMPTY(m.frdate)
		SET FILTER TO arrhea<=m.arrser
	ELSE
		SET FILTER TO arrhea>=m.afrser AND arrhea<=m.arrser
	ENDIF
	SELECT 0
	USE protab
	SELECT 0
	USE famtab1
	SET ORDER TO serial
	SELECT 0
	USE fl2tab
	SET ORDER TO serial
	SELECT 0
	USE fl1tab
	SET ORDER TO serial
	SELECT 0
	USE fl4tab
	SET ORDER TO serial
	SELECT 0
	USE rathea
	SELECT 0
	USE ratrow
	SELECT 0
	USE sprtab
	SELECT protab
	SET RELATION TO Protab.famtab1 INTO Famtab1 ADDITIVE
	SET RELATION TO Protab.fl1tab  INTO Fl1tab ADDITIVE
	SET RELATION TO Protab.fl2tab  INTO Fl2tab ADDITIVE
	SELECT mourow
	SET RELATION TO Mourow.mouhea  INTO Mouhea ADDITIVE
	SELECT arrrow
	SET RELATION TO Arrrow.arrhea  INTO Arrhea ADDITIVE		
	SELECT protab
	SET ORDER TO proray
	GO TOP
	= getinvgrp()
	SELECT sprtab
	GO TOP
	DO WHILE !EOF()
		SELECT invtab
		SUM qtyini  TO tqtyini  FOR name=ALLTRIM(sprtab.name)+' ' AND famtab1=sprtab.famtab1
		SUM qtyexit TO tqtyexit FOR name=ALLTRIM(sprtab.name)+' ' AND famtab1=sprtab.famtab1
		SUM invqty  TO tinvqty  FOR name=ALLTRIM(sprtab.name)+' ' AND famtab1=sprtab.famtab1
**		SUM qtyini  TO tqtyini  FOR ALLTRIM(sprtab.name) $ name
**		SUM qtyexit TO tqtyexit FOR ALLTRIM(sprtab.name) $ name
**		SUM invqty  TO tinvqty  FOR ALLTRIM(sprtab.name) $ name
		SELECT sprtab
		REPLACE qtyini  WITH tqtyini
		REPLACE qtyexit WITH tqtyexit
		REPLACE invqty  WITH tinvqty
		SKIP
	ENDDO
	SELECT protab
	SET RELATION OFF INTO famtab1
	SELECT sprtab
	SET ORDER TO fgname
	SET RELATION TO sprtab.famtab1 INTO famtab1 ADDITIVE
	IF m.okbtn=1
		REPORT FORM proredt9 PREVIEW
	ELSE
		IF m.okbtn=2
			REPORT FORM proredt9 TO PRINT OFF
		ENDIF	
	ENDIF			
ENDIF
CLOS DATA

PROCEDURE getinvgrp
*******************
SELECT protab
SET ORDER TO proray
IF m.famtab<>0
	SET FILTER TO famtab=m.famtab
ENDIF
GO TOP
DO WHILE !EOF()
	IF selling<>0 AND tiniqty2(serial,adjqty)<>0
		SELECT protab
		SCATTER MEMVAR
		SELECT invtab
		APPEND BLANK
		GATHER MEMVAR
		REPLACE qtyini  WITH m.qtyini
		REPLACE qtyexit WITH texitqty1()
		SELECT invtab
		REPLACE invqty  WITH m.qtyini-m.qtyexit
	ENDIF	
	SELECT protab
	SKIP
ENDDO

******
PROCEDURE getmodls1
*******************
CLOSE DATA
m.qtyexit  = 0
m.arrprice = 0
m.out      = 1
m.okbtn    = 1
m.dspname  = ''
DO printout.spr
IF m.okbtn=2
	RETURN
ENDIF
SELECT 0
USE mourow
SELECT 0
USE rathea
SELECT 0
USE ratrow
SELECT 0
USE arrrow
SET ORDER TO protab
SET FILTER TO bidon="X"
SELECT 0
USE arrhea
SET ORDER TO serial
SELECT 0
USE protab
set filt to !EMPTY(qte2)
SELECT 0
USE famtab1
SET ORDER TO serial
SELECT 0
USE fl2tab
SET ORDER TO serial
SELECT 0
USE fl3tab
SET ORDER TO serial
SELECT 0
USE fl4tab
SET ORDER TO serial
SELECT 0
USE fl1tab
SET ORDER TO serial
SELECT protab
SET RELATION TO Protab.serial  INTO arrrow ADDITIVE
SET RELATION TO Protab.famtab1 INTO Famtab1 ADDITIVE
SET RELATION TO Protab.fl1tab  INTO Fl1tab ADDITIVE
SET RELATION TO Protab.fl2tab  INTO Fl2tab ADDITIVE
SET RELATION TO Protab.fl3tab  INTO Fl3tab ADDITIVE
SET RELATION TO Protab.fl4tab  INTO Fl4tab ADDITIVE
SELECT protab
SET ORDER TO mod
GO TOP
IF m.out=1
	REPORT FORM proredt8 PREVIEW 
ELSE
	REPORT FORM proredt8 TO PRINT OFF
ENDIF	
CLOSE DATA

PROCEDURE getmodls221
*******************
CLOSE DATA
m.qtyexit  = 0
m.arrprice = 0
m.out      = 1
m.okbtn    = 1
DO printrec.spr
IF m.okbtn=2
	RETURN
ENDIF
SELECT 0
USE mourow
SELECT 0
USE rathea
SELECT 0
USE ratrow
SELECT 0
USE arrrow
SELECT arrrow
SET ORDER TO protab
SET FILTER TO arrhea  >= 76
SELECT 0
USE arrhea
SET ORDER TO serial
**sele 0
**use newarr
**set order to barcode
**SET RELATION TO Newarr.barcode INTO Protab ADDITIVE
SELECT 0
USE protab
**SET FILTER TO !EMPTY(qte)
**set filt to new="1"
set filt to newdate={08/11/99}
SELECT 0
USE famtab1
SET ORDER TO serial
SELECT 0
USE fl2tab
SET ORDER TO serial
SELECT 0
USE fl3tab
SET ORDER TO serial
SELECT 0
USE fl4tab
SET ORDER TO serial
SELECT 0
USE fl1tab
SET ORDER TO serial
SELECT protab
SET RELATION TO Protab.serial  INTO arrrow ADDITIVE
SET RELATION TO Protab.famtab1 INTO Famtab1 ADDITIVE
SET RELATION TO Protab.fl1tab  INTO Fl1tab ADDITIVE
SET RELATION TO Protab.fl2tab  INTO Fl2tab ADDITIVE
SET RELATION TO Protab.fl3tab  INTO Fl3tab ADDITIVE
SET RELATION TO Protab.fl4tab  INTO Fl4tab ADDITIVE
SELECT protab
SET ORDER TO mod
GO TOP
IF m.out=1
	REPORT FORM proredt8 PREVIEW 
ELSE
	REPORT FORM proredt8 TO PRINT prompt OFF
ENDIF	
CLOSE DATA

PROCEDURE msg_box
*****************
PARAMETER tstring,tbutton,ttitle

m.tmessage = tstring
m.title    = ttitle
IF tbutton=1
	button1 = IIF(m.publang,'\<OK','\<OK')
	button2 = IIF(m.publang,'\?\<Cancel','\?\<Annuler')
ELSE
	button1 = IIF(m.publang,'\<Yes','\<Oui')
	button2 = IIF(m.publang,'\?\<No','\?\<Non')
ENDIF
msgans = .F.
DO msgbox.spr
RETURN msgans

PROCEDURE getseainv
*******************
CLOSE DATA
m.qtyexit = 0
m.qtyini  = 0
m.tot     = 0
m.tot1    = 0
m.pric    = 0
m.famtab  = 0
m.todate  = DATE()
m.mouser  = 0
m.arrser  = 0
m.intser  = 0
m.okbtn   = 3
DO ivtredt2.spr
IF m.okbtn=3
	RETURN
ENDIF
SELECT 0
USE famtab
SELECT 0
USE invtab
= checkzapfile('invtab')
SELECT 0
USE mouhea
IF m.todate<DATE()
	SET FILTER TO date<=m.todate
	GO BOTTOM
	m.mouser = serial
ENDIF
SET ORDER TO serial
SELECT 0
USE mourow
IF m.todate<DATE()
	SET FILTER TO mouhea<=m.mouser
ENDIF
SELECT 0
USE arrhea
IF m.todate<DATE()
	SET FILTER TO date<=m.todate
	GO BOTTOM
	m.arrser = serial
ENDIF
SET ORDER TO serial
GO TOP
SELECT 0
USE arrrow
SET ORDER TO protab
IF m.todate<DATE()
	SET FILTER TO arrhea<=m.arrser
ENDIF
SELECT 0
USE genpar
SELECT 0
USE protab
SELECT 0
USE rathea
SELECT 0
USE ratrow
SELECT 0
USE inthea
IF m.todate<DATE()
	SET FILTER TO date<=m.todate
	GO BOTTOM
	m.intser = serial
ENDIF
SET ORDER TO serial
GO TOP
SELECT 0		
USE introw
SET ORDER TO protab
IF m.todate<DATE()
	SET FILTER TO inthea<=m.intser
ENDIF
SELECT mourow
SET RELATION TO Mourow.mouhea  INTO Mouhea ADDITIVE
SELECT arrrow
SET RELATION TO Arrrow.arrhea  INTO Arrhea ADDITIVE
SELECT introw
SET RELATION TO Introw.inthea  INTO Inthea ADDITIVE
SELECT protab
SET ORDER TO proray
GO TOP
= getinvcost()
SELECT invtab
SET ORDER TO famtab
SELECT famtab
GO TOP
DO WHILE !EOF()
	m.serial = serial
	m.qtyin  = 0
	m.qtyout = 0
	m.qtycur = 0
	m.totin  = 0
	m.totcur = 0
	SELECT invtab
	SEEK m.serial
	DO WHILE !EOF() AND famtab=m.serial
		m.qtyin  = m.qtyin  + qtyini
		m.qtyout = m.qtyout + qtyexit
		m.qtycur = m.qtycur + invqty
		m.totin  = m.totin  + totini
		m.totcur = m.totcur + totsel
		SELECT invtab
		SKIP
	ENDDO
	SELECT famtab
	REPLACE qtyin  WITH m.qtyin
	REPLACE qtyout WITH m.qtyout
	REPLACE qtycur WITH m.qtycur
	REPLACE totin  WITH m.totin
	REPLACE totcur WITH m.totcur
	SKIP
ENDDO
IF m.okbtn=1
	REPORT FORM ivsredt PREVIEW
ELSE
	REPORT FORM ivsredt TO PRINT OFF
ENDIF
CLOSE DATA
