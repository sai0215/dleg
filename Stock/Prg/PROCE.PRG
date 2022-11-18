PROCEDURE btn_val
*****************
PARAMETER m.tbtnname
DO CASE
	CASE m.tbtnname='NEW'
 		isediting= .T.
		isadding = .T.
		m.issearch=.F.
		=edithand('NEW')
		_curobj=1
		m.notedit=.T.		
		DO refresh
		SHOW GETS
		RETURN
	CASE m.tbtnname='EDIT'
		IF EOF() OR BOF()
		   WAIT WINDOW ENDFILE NOWAIT
		   RETURN
		ENDIF
 		isediting= .T.
		_curobj=1
		m.notedit=.T.				
		DO refresh
		RETURN
	CASE m.tbtnname='DEL' 
		IF EOF() OR BOF()
			WAIT WINDOW ENDFILE NOWAIT
			RETURN
		ENDIF
		IF fox_alert(DELREC)
			DELETE
			IF !EOF() AND DELETED()
				SKIP 1
			ENDIF
			IF EOF()
				WAIT WINDOW ENDFILE NOWAIT
				GO BOTTOM
			ENDIF
		ELSE 
			RETURN	
		ENDIF
		m.notedit=.T.		
		DO refresh
	CASE m.tbtnname='LOCATE'
		m.issearch=.T.
		DO loc
 	    IF !EOF() AND !BOF()
	   		SHOW GET m.top ENABLE
			SHOW GET m.prev ENABLE
			SHOW GET m.next ENABLE
		    SHOW GET m.bottom ENABLE
		ELSE
			IF EOF()
		        WAIT WINDOW ENDFILE NOWAIT
			    SHOW GET m.top ENABLE
				SHOW GET m.prev ENABLE
				SHOW GET m.next DISABLE
				SHOW GET m.bottom DISABLE
			ENDIF
	 		IF BOF()
				WAIT WINDOW TOPFILE NOWAIT
			    SHOW GET m.top DISABLE
			    SHOW GET m.prev DISABLE
			    SHOW GET m.next ENABLE
			    SHOW GET m.bottom ENABLE
			ENDIF			   
		ENDIF
		m.notedit=.F.		
		DO refresh		
        RETURN
	CASE m.tbtnname='PRINT'
		IF EOF() OR BOF()
		   WAIT WINDOW ENDFILE NOWAIT
		   RETURN
		ENDIF
		DO printrec.spr
		RETURN        					
	CASE m.tbtnname='TOP'
   		m.issearch=.T.
		GO TOP
		WAIT WINDOW TOPFILE NOWAIT
		SHOW GET m.top DISABLE
		SHOW GET m.prev DISABLE
		SHOW GET m.next ENABLE
		SHOW GET m.bottom ENABLE
		m.notedit=.F.
		DO refresh
	CASE m.tbtnname='PREV'
   		m.issearch=.T.
	 	IF !BOF()
		   SKIP -1
	 	   SHOW GET m.next ENABLE
		   SHOW GET m.bottom ENABLE
		ENDIF
	 	IF BOF()
		   WAIT WINDOW TOPFILE NOWAIT
		   GO TOP
		   SHOW GET m.top DISABLE
		   SHOW GET m.prev DISABLE
		   SHOW GET m.next ENABLE
		   SHOW GET m.bottom ENABLE
		ENDIF
		m.notedit=.F.
		DO refresh
	CASE m.tbtnname='NEXT'
   		m.issearch=.T.
 	    IF !EOF()
			SKIP 1
	   		SHOW GET m.top ENABLE
			SHOW GET m.prev ENABLE
		ENDIF
		IF EOF()
	        WAIT WINDOW ENDFILE NOWAIT
		    GO BOTTOM
		   	SHOW GET m.top ENABLE
			SHOW GET m.prev ENABLE
			SHOW GET m.next DISABLE
			SHOW GET m.bottom DISABLE
		ENDIF
		m.notedit=.F.		
		DO refresh
	CASE m.tbtnname='BOTTOM'
	   m.issearch=.T.
		GO BOTTOM
		WAIT WINDOW ENDFILE NOWAIT
		SHOW GET m.top ENABLE
		SHOW GET m.prev ENABLE
		SHOW GET m.next DISABLE
		SHOW GET m.bottom DISABLE
		m.notedit=.F.		
		DO refresh
	CASE m.tbtnname='SAVE'
		IF isadding
			=edithand('SAVE')
		ELSE
			GATHER MEMVAR MEMO
		ENDIF				
		isediting=.F.
		isadding=.F.		
		m.notedit=.T.
		_CUROBJ=OBJNUM(m.new)
		DO refresh
	CASE m.tbtnname='UNDO'
		IF isadding
			=edithand('UNDO')
		ENDIF	
		isediting=.F.
		isadding=.F.		
		m.notedit=.T.				
		DO refresh
		IF !full
			SCATTER MEMO MEMVAR BLANK 
			SHOW GETS
			RETURN
		ENDIF
	CASE m.tbtnname='CLOSE'
		CLEAR READ
		RETURN
ENDCASE
SCATTER MEMVAR MEMO
SHOW GETS
RETURN


PROCEDURE refresh
*****************
DO CASE
   CASE m.isediting
   		IF m.notedit
  			SHOW GET m.top DISABLE
			SHOW GET m.prev DISABLE
			SHOW GET m.next DISABLE
			SHOW GET m.bottom DISABLE
		ENDIF
		SHOW GET m.locate DISABLE
		SHOW GET m.new DISABLE
		SHOW GET m.edit DISABLE
		SHOW GET m.del DISABLE
		SHOW GET m.print DISABLE
		SHOW GET m.close DISABLE
		SHOW GET m.save ENABLE
		SHOW GET m.undo ENABLE
		RETURN
   CASE (RECCOUNT()=0 OR EOF()) AND !m.isediting
		SHOW GET m.edit DISABLE
		SHOW GET m.del DISABLE
		SHOW GET m.locate DISABLE
		SHOW GET m.print DISABLE
		SHOW GET m.top DISABLE
		SHOW GET m.prev DISABLE
		SHOW GET m.next DISABLE
		SHOW GET m.bottom DISABLE
		SHOW GET m.save DISABLE
		SHOW GET m.undo DISABLE
		SHOW GET m.new ENABLE
		SHOW GET m.close ENABLE		
	OTHERWISE
   		IF notedit
   			SHOW GET m.top ENABLE
			SHOW GET m.prev ENABLE
			SHOW GET m.next ENABLE
			SHOW GET m.bottom ENABLE
		ENDIF
		SHOW GET m.locate ENABLE
		SHOW GET m.new ENABLE
		IF issearch
	 	   SHOW GET m.edit ENABLE
		   SHOW GET m.del ENABLE
  		   SHOW GET m.print ENABLE
		ENDIF   
		SHOW GET m.close ENABLE
		SHOW GET m.save DISABLE
		SHOW GET m.undo DISABLE
ENDCASE
RETURN


PROCEDURE edithand
******************
PARAMETER m.paction
DO CASE
	CASE m.paction = 'NEW'
		SCATTER MEMVAR MEMO BLANK
	CASE m.paction = 'SAVE'
		INSERT INTO (ALIAS()) FROM MEMVAR
	CASE m.paction = 'UNDO'
		*nothing here
ENDCASE
RETURN				

PROCEDURE fox_alert
*******************
PARAMETER tstring
m.tmessage = tstring
DO message.spr
RETURN m.drec

PROCEDURE pro_pop
*****************
IF lastkey()=13
   DEACTIVATE POPUP 
ENDIF

PROCEDURE pro_popx
******************
IF lastkey()=13
   REPLACE bidon WITH IIF(EMPTY(bidon),'>','')
ENDIF


*PROCEDURE print
***************
*PARAMETERS tname,tcondition,ttype
*REPORT FORMAT &tname FOR &tcondition &ttype

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
tmin  = ABS(tmin)
tthour = INT(tmin/60)
ttmin  = tmin%60
ttsign=IIF(tsign=1,ttsign,'')
RETURN ttsign+PADL(ALLTRIM(STR(tthour)),tlen,'0')+':'+PADL(ALLTRIM(STR(ttmin)),2,'0')

PROCEDURE spop
**************
PARAMETERS tfile,tfield,tx,ty,tsizex,tsizey,tvariable
&tvariable=PROPER(&tvariable)
SELECT (tfile)
DEFINE POPUP ttpop FROM tx,ty TO tx+tsizex,ty+tsizey PROMPT FIELD &tfield SCROLL
ON SELECTION POPUP ttpop DO pro_pop
LOCATE FOR ALLTRIM(&tfield)==ALLTRIM(&tvariable)
IF FOUND()
   STORE &tfield TO &tvariable
   SHOW GETS
   RETURN .T.
ENDIF   
SET FILTER TO ALLTRIM(&tfield)=ALLTRIM(&tvariable)
GO TOP
IF EOF()
   WAIT WINDOW IIF(m.publang,'No matching record',"Enregistrement non trouvé") NOWAIT
   SET FILTER TO
   RETURN .F.
ELSE
   ACTIVATE POPUP ttpop   
ENDIF
SET FILTER TO
RELEASE POPUP ttpop
STORE &tfield TO &tvariable
SHOW GETS
RETURN .T.

PROCEDURE drwbox
****************
PARAMETERS tminx,tminy,tmaxx,tmaxy,tx,tcolor,tlcolor,tccolor,tecolor,ttype,tpattern
tred = VAL(SUBSTR(tcolor,1,3))
tgreen = VAL(SUBSTR(tcolor,4,3))
tblue = VAL(SUBSTR(tcolor,7,3))
tlred = VAL(SUBSTR(tlcolor,1,3))
tlgreen = VAL(SUBSTR(tlcolor,4,3))
tlblue = VAL(SUBSTR(tlcolor,7,3))
tcred = VAL(SUBSTR(tccolor,1,3))
tcgreen = VAL(SUBSTR(tccolor,4,3))
tcblue = VAL(SUBSTR(tccolor,7,3))
tered = VAL(SUBSTR(tecolor,1,3))
tegreen = VAL(SUBSTR(tecolor,4,3))
teblue = VAL(SUBSTR(tecolor,7,3))
tlastx = MROW()
tmrow = tlastx
DO WHILE MDOWN()
   = INKEY(.06,'HM')
   trow = MROW( )
   tcol = MCOL( )
   IF BETWEEN(trow,IIF(ttype =2,tx,tminx),IIF(ttype=1,tx,tmaxx)) AND BETWEEN(tcol,tminy,tmaxy)
      IF ttype = 1 AND  trow <= tx
         tmrow = IIF(trow <tmrow,trow-.1,tmrow)
         @ trow,tminy TO tx,tmaxy PATTERN 1 COLOR RGB(tred,tgreen,tblue,tred,tgreen,tblue)
         @ tmrow,tminy TO trow,tmaxy PATTERN tpattern COLOR RGB(tlred,tlgreen,tlblue,tlred,tlgreen,tlblue)
         tlastx =MROW()
      ENDIF
      IF ttype = 2 AND  trow >= tx
         tmrow = IIF(trow >tmrow,trow+.1,tmrow)
         @ tx,tminy TO trow,tmaxy PATTERN 1 COLOR RGB(tred,tgreen,tblue,tred,tgreen,tblue)
         @ trow,tminy TO tmrow,tmaxy PATTERN tpattern COLOR RGB(tlred,tlgreen,tlblue,tlred,tlgreen,tlblue)
         tlastx =MROW()
      ENDIF
      IF ( BETWEEN(trow,tx-0.24,tx) AND ttype = 1 ) OR ( BETWEEN(trow,tx,tx+0.24) AND ttype = 2 )
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
t_tx=m.tpos+((m.pubbox/24)*ttime)/60
RETURN t_tx

PROCEDURE fillbox2
*****************
PARAMETERS ty,tfrom,tto,tlen,tcolor,tpattern
IF NOT EMPTY(tcolor)
   @ tfrom,ty CLEAR TO tto,ty+tlen
ENDIF
m.trcol = VAL(SUBSTR(tcolor,1,3))
m.tgcol = VAL(SUBSTR(tcolor,4,3))
m.tbcol = VAL(SUBSTR(tcolor,7,3))
@ tfrom,ty TO tto,ty+tlen PATTERN tpattern COLOR RGB(trcol,tgcol,tbcol,trcol,tgcol,tbcol)

*PROCEDURE fillarray
*******************
*SELE grptab
*GO TOP

*FOR i = 1 TO gnc
*	grpname(i) = name
*	IF !EOF()
*		SKIP 1 IN grptab
*	ENDIF	
*ENDFOR

*SELE clutab

PROCEDURE loc
*************
DEFINE WINDOW wlocate FROM 1,1 TO 20,40;
	SYSTEM GROW CLOSE ZOOM FLOAT FONT "Times New Roman",10
MOVE WINDOW wlocate CENTER
BROWSE WINDOW wlocate NOEDIT NODELETE ;
	NOMENU TITLE BRTITLE
RELEASE WINDOW wlocate	
SCATTER MEMVAR MEMO
SHOW GETS	

PROCEDURE chkdate
*****************
SELECT reserv
SET FILTER TO idno=m.idno AND code=m.code AND date>(DATE()-DOW(DATE())) AND date<(DATE()-DOW(DATE())+8)
GO TOP
rescnt=0
DO WHILE !EOF()
	rescnt=rescnt+1
	SKIP 1
ENDDO
IF rescnt>=pubres AND pubres<>0
	WAIT WINDOW ALLTRIM(STR(m.idno)) + IIF(m.publang,' used his hours',' a utilisé ses heures') NOWAIT
	SET FILTER TO
ELSE
	SET FILTER TO
ENDIF

PROCEDURE PTABLE
****************
IF MDOWN()
	@ MROW()-1,MCOL()-1 SAY 'TABLE.BMP' BITMAP
ENDIF	

PROCEDURE spop1
**************
PARAMETERS tfile,tfield,tx,ty,tsizex,tsizey,tvariable
&tvariable=PROPER(&tvariable)
SELECT (tfile)
tfilter=FILTER()
DEFINE POPUP ttpop FROM tx,ty TO tx+tsizex,ty+tsizey PROMPT FIELD &tfield SCROLL
ON SELECTION POPUP ttpop DO pro_pop
LOCATE FOR ALLTRIM(&tfield)==ALLTRIM(&tvariable)
IF FOUND()
   STORE &tfield TO &tvariable
   SHOW GETS
   RETURN .T.
ENDIF   
SET FILTER TO ALLTRIM(&tfield)=ALLTRIM(&tvariable) AND &tfilter
GO TOP
IF EOF()
   WAIT WINDOW IIF(m.publang,'No matching record',"Enregistrement non trouvé") NOWAIT
   SET FILTER TO
   RETURN .F.
ELSE
   ACTIVATE POPUP ttpop   
ENDIF
SET FILTER TO &tfilter
RELEASE POPUP ttpop
STORE &tfield TO &tvariable
SHOW GETS
RETURN .T.