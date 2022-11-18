= AFIELD(varary)
fldcnt = ALEN(varary,1)
FOR if = 1 TO fldcnt-1
	IF varary[if,2] = 'G'
		= ADEL(varary,if)
		varary[fldcnt,1] = ''
	ENDIF
ENDFOR		
DIMENSION varary[9,3]
varary[1,1] = 'LINE'
varary[2,1] = 'FAMILY'
varary[3,1] = 'SUB-FAMILY'
varary[4,1] = 'THEME'
varary[5,1] = 'MODEL'
varary[6,1] = 'COLOR'
varary[7,1] = 'SIZE'
varary[8,1] = 'LABEL'
varary[9,1] = 'SEASON'

varary[1,2] = 'FAMTAB1'
varary[2,2] = 'FL1TAB'
varary[3,2] = 'FL2TAB'
varary[4,2] = 'FIELD1'
varary[5,2] = 'FIELD2'
varary[6,2] = 'FT3TAB'
varary[7,2] = 'FT4TAB'
varary[8,2] = 'NAME'
varary[9,2] = 'FAMTAB'

varary[1,3] = 'T'
varary[2,3] = 'T'
varary[3,3] = 'T'
varary[4,3] = 'F'
varary[5,3] = 'F'
varary[6,3] = 'T'
varary[7,3] = 'T'
varary[8,3] = 'F'
varary[9,3] = 'T'

FOR lp = 1 TO 9
	mm = 'fld'+ STR(lp)
	&mm = varary[lp]
ENDFOR
	
m.res = ''
crtr = ''

DEFINE WINDOW loc_win FROM 23,4 TO 35,60 TITLE 'SEARCH window' FONT 'ms sans serif',10
MOVE WINDOW loc_win CENTER
ACTIVATE WINDOW loc_win

@ 1,1 GET m.fld ;
 	PICTURE "@&N" ;
	FROM varary ;
	SIZE 4,21 ;
	DEFAULT 1 ;
	FONT "MS Sans Serif", 8 ;
	STYLE ""

@ 1,19 GET m.cond ;
 	PICTURE "@^ =; <; >; <=; >=; <>" ;
	SIZE 4,11 ;
	DEFAULT '=' ;
	FONT "MS Sans Serif", 8 ;
	VALID tovar() ;
	STYLE ""

@ 1,30 SAY 'CRITERIA:' ;
	FONT "MS Sans Serif", 8 ;
	STYLE "b" 
	
@ 2,30 GET m.var ;
	SIZE 1,20 ;
	FUNCTION 'k' ;
	DEFAULT " "

@ 1,52.5 GET m.jnct ;
 	PICTURE "@^ ;AND ;OR" ;
	SIZE 4,15 ;
	DEFAULT ' ' ;
	FONT "MS Sans Serif", 8 ;
	VALID chkjnct();
	STYLE ""

@ 6,1 GET m.ok ;
	PICTURE "@*HN \<GO; \<CONTINUE; CLEA\<R; \<QUIT" ;
	SIZE 1.3,14,2 ;
	DEFAULT 1 ;
	VALID _q8s1e61d()

SHOW GET m.ok ,2 DISABLE

READ CYCLE

RELEASE WINDOW loc_win

IF WEXIST("control")
*	ACTIVATE WINDOW control
	IF WEXIST("fil_apr")
		SHOW GETS	
	ELSE
		DO drwall
	ENDIF
	ACTIVATE WINDOW control
ENDIF

FUNCTION tovar
_CUROBJ = OBJNUM(m.var)


FUNCTION _q8s1e61d     &&  m.ok VALID
#REGION 1

DO CASE

	CASE m.ok = 1
		IF EMPTY(crtr)
			DO bldcrtr
		ENDIF	
		SET ESCAPE OFF
		LOCATE FOR &crtr
		IF NOT FOUND() 
			SHOW GET m.ok,2 DISABLE
			WAIT WINDOW "No matching Record" NOWAIT
		ELSE
			SHOW GET m.ok,2 ENABLE
			m.rec = RECNO()
			m.res = &varary[m.fld,1]
			DO upddsp 
			WAIT WINDOW "Record Found: "+ALLTRIM(STR(RECNO())) NOWAIT
		ENDIF
		SET ESCAPE ON
			
	CASE m.ok = 2
			CONTINUE
			IF NOT FOUND() 
				WAIT WINDOW "No More Records Found" NOWAIT
				GO m.rec
				EXIT
			ELSE
				m.rec = RECNO()
				m.res = &varary[m.fld,1]
				DO upddsp 
				WAIT WINDOW "Record Found: "+ALLTRIM(STR(RECNO())) 
			ENDIF
		RETURN 0

	CASE m.ok = 3
		crtr = ''
		m.var = ' '
		@ 2,30 GET m.var ;
			SIZE 1,20 ;
			FUNCTION 'k' ;
			DEFAULT " "

	CASE m.ok = 4
		CLEAR READ
		
ENDCASE		

SHOW GETS

FUNCTION bldcrtr
IF varary(m.fld,3)='F'
	SELECT ALLTRIM(varary(m.fld,2))
	LOCATE FOR name = RTRIM(m.var)
	IF FOUND
		tvar = serial
	ELSE
		tvar = ""
	ENDIF	
	tfld1 = 
ELSE

ENDIF	
tfld = IIF(varary(m.fld,3)='F', ALLTRIM(varary(m.fld,2)), ALLTRIM(ALLTRIM(varary(m.fld,2))+"."+"name"))
tfld2 = IIF(varary(m.fld,3)='F', RTRIM(m.var), ALLTRIM(ALLTRIM(varary(m.fld,2))+"."+"name"))
m.crtr = m.crtr+' '+ tfld +' '+m.cond+" '"+ RTRIM(m.var)+"' "+m.jnct 
@ 4.7,1 SAY m.crtr STYLE 'B'


FUNCTION chkjnct
m.crtr = m.crtr+' '+ALLTRIM(varary(m.fld,2))+' '+m.cond+" '"+RTRIM(m.var)+"' "+m.jnct 
@ 4.7,1 SAY m.crtr STYLE 'B'
IF NOT EMPTY(m.jnct)
	_CUROBJ = OBJNUM(m.fld)
	m.var = ''
	@ 2,30 GET m.var ;
		SIZE 1,20 
	SHOW GET m.fld
ENDIF
	

FUNCTION upddsp
@ 3.3,30 SAY m.res ;
	SIZE 1,35 ;
	FONT "MS Sans Serif", 10;
