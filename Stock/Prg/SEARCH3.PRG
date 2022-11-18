PARAMETERS adbf, afltr, anbr, aslct
DIMENSION tmpary[anbr], objary[anbr], fldary[anbr]

*       *********************************************************
*       *                                                         
*       *               Windows Window definitions                
*       *                                                         
*       *********************************************************
*

IF NOT WEXIST("wsearch") ;
	OR UPPER(WTITLE("WSEARCH")) == "WSEARCH.PJX" ;
	OR UPPER(WTITLE("WSEARCH")) == "WSEARCH.SCX" ;
	OR UPPER(WTITLE("WSEARCH")) == "WSEARCH.MNX" ;
	OR UPPER(WTITLE("WSEARCH")) == "WSEARCH.PRG" ;
	OR UPPER(WTITLE("WSEARCH")) == "WSEARCH.FRX" ;
	OR UPPER(WTITLE("WSEARCH")) == "WSEARCH.QPR"
	DEFINE WINDOW wsearch ;
		AT  1.750, 43.750  ;
		SIZE 36.444,60.571 ;
		TITLE "Search" ;
		FONT "FoxFont", 8 ;
		FLOAT ;
		NOCLOSE ;
		NOMINIMIZE ;
		SYSTEM ;
		COLOR RGB(,,,192,192,192)
	MOVE WINDOW wsearch CENTER
ENDIF


*       *********************************************************
*       *                                                         
*       *         SEARCH3/Windows Setup Code - SECTION 2          
*       *                                                         
*       *********************************************************
*

#REGION 1

fl = ''
ttmp = ''
ttfld = ''

FOR lp = 1 TO anbr
	
	tmpary[lp] = ''
	objary[lp] = 0
ENDFOR	

SELECT *;
 FROM &adbf;
 HAVING &afltr ;
 INTO CURSOR Query

FOR lp = 2 TO anbr
	IF !EMPTY(varary[lp,3])
		relfld = 'query.'+ varary[lp,2]
		SET RELATION TO &relfld INTO varary[lp,2] ADDITIVE
	ENDIF
ENDFOR		
		
brwlst = IIF(EMPTY(varary[1,3]), varary[1,2]+':15', varary[1,2]+'.'+varary[1,3])+':10:H="'+varary[1,1]+'"'
FOR lp = 2 TO anbr
	brwlst = brwlst + ','+ IIF(EMPTY(varary[lp,3]), varary[lp,2]+':10', varary[lp,2]+'.'+varary[lp,3])+':10:H="'+varary[lp,1]+'"'
ENDFOR	

fldnbr = 1
nbrfld = 0
nbrbtn = "@*RVN "
FOR lp = 1 TO anbr-1
	nbrbtn = nbrbtn + "; " 
ENDFOR	

DEFINE POPUP fldpop FROM 0,0 TO 5,34.5 PROMPT FIELDS &varary[1,2] SCROLL
ON SELECTION POPUP fldpop DO pro_pop

m.publang = .T.
IF !m.publang
	MODIFY WINDOW wsearch TITLE 'Chercher'
	MOVE WINDOW wsearch CENTER	
ENDIF	

*       *********************************************************
*       *                                                         
*       *              SEARCH3/Windows Screen Layout              
*       *                                                         
*       *********************************************************
*

#REGION 1
IF WVISIBLE("wsearch")
	ACTIVATE WINDOW wsearch SAME
ELSE
	ACTIVATE WINDOW wsearch NOSHOW
ENDIF

FOR lp = 1 TO anbr
	@ 1.222+lp*3,12.286 TO 3.111+lp*3,31.715 ;
		PEN 1, 8 ;
		COLOR RGB(128,128,128,,,,)
	@ 3.000+lp*3,12.429 TO 3.000+lp*3,31.572 ;
		PEN 1, 8 ;
		STYLE "1" ;
		COLOR RGB(255,255,255,255,255,255)
	@ 1.333+lp*3,31.571 TO 3.111+lp*3,31.571 ;
		PEN 1, 8 ;
		COLOR RGB(255,255,255,255,255,255)
	@ 1.444+lp*3,4.571 SAY varary[lp,1] ;
		SIZE 1.000,10.000 ;
		FONT "MS Sans Serif", 8 ;
		STYLE "T" ;
		COLOR RGB(,,,255,255,255)
	IF EMPTY(varary[lp,3])
		@ 1.444+lp*3,12.714 GET tmpary[lp] ;
			SIZE 1.000,25.800 ;
			DEFAULT " " ;
			FONT "MS Sans Serif", 8 ;
			PICTURE "@K" ;
			VALID actfltr() ;
			WHEN gettmp() ;
			COLOR ,RGB(,,,255,255,255)
		nbrfld = nbrfld +1
		objary[lp] = nbrfld 
	ELSE
		@ 1.444+lp*3,12.714 SAY tmpary[lp] ;
			SIZE 1.000,25.800 ;
			FONT "MS Sans Serif", 8 ;
			COLOR ,RGB(,,,255,255,255)
	ENDIF
ENDFOR

@ 1,2.000 GET m.chk ;
	PICTURE "@*C  "+aslct ;
	SIZE 1.308,5.167 ;
	DEFAULT 1 ;
	FONT "MS Sans Serif", 8 ;
	STYLE "BT" ;
	VALID qtyfltr()

@ 4.2,1.857 GET m.sl ;
	PICTURE nbrbtn ;
	SIZE 1.308,4.833,0.8 ;
	DEFAULT 1 ;
	FONT "MS Sans Serif", 8 ;
	STYLE "BT" ;
	VALID _r9d0yuh2s()

@ 2.333,34.143 TO 34.333,58.143 ;
	PEN 0, 8 ;
	COLOR RGB(128,128,128,,,,)
@ 2.333,58.000 TO 34.222,58.000 ;
	PEN 1, 8 ;
	COLOR RGB(255,255,255,255,255,255)
@ 34.222,34.286 TO 34.222,58.000 ;
	PEN 1, 8 ;
	STYLE "1" ;
	COLOR RGB(255,255,255,255,255,255)

@ 2.444,34.286 GET m.slct ;
 	PICTURE "@&N" ;
	POPUP fldpop ;
	SIZE 21.923,32.800 ;
	DEFAULT " " ;
	FONT "MS Sans Serif", 8 ;
	STYLE ""

@ 31.444,9.000 GET m.loc_btn ;
	PICTURE "@*BHN " + ;
		(LOCFILE("..\bmp\locate.bmp","BMP|ICO|PCT|ICN","Where is locate?")) ;
	SIZE 2.889,3.714,0.571 ;
	DEFAULT 1 ;
	FONT "FoxFont", 8 ;
	VALID brws() ;
	MESSAGE IIF(m.publang,'Locate a record',"Chercher un enregistrement")
@ 31.444,14.714 GET m.add_btn ;
	PICTURE "@*BHN " + ;
		(LOCFILE("..\bmp\division.bmp","BMP|ICO|PCT|ICN","Where is new?")) ;
	SIZE 2.889,3.857,0.571 ;
	DEFAULT 1 ;
	FONT "FoxFont", 8 ;
	VALID asmbl() ;
	MESSAGE IIF(m.publang,'Add a new record',"Ajout d'un nouveau enregistrement")
@ 31.444,20.714 GET m.squit_btn ;
	PICTURE "@*BHN " + ;
		(LOCFILE("..\bmp\close.bmp","BMP|ICO|PCT|ICN","Where is close?")) ;
	SIZE 2.889,3.714,0.571 ;
	DEFAULT 1 ;
	FONT "FoxFont", 8 ;
	VALID _r9d0yugvh()

IF NOT WVISIBLE("wsearch")
	ACTIVATE WINDOW wsearch
ENDIF

READ CYCLE 

RELEASE WINDOW wsearch

RELEASE POPUP fldpop

SELECT &adbf
LOCATE FOR serial = query.serial
m.sfound = .T.

*****************************************************
FUNCTION gettmp
ttmp = ALLTRIM(tmpary[sl])

FUNCTION actfltr
IF ALLTRIM(tmpary[sl]) == ttmp
	RETURN
ENDIF
	
vfld = varary[sl,2]
	
DEFINE POPUP fldpop FROM 0,0 TO 5,34.5 PROMPT FIELDS &vfld SCROLL

fl = "UPPER(&vfld) = UPPER(ALLTRIM(tmpary[sl]))"
SET FILTER TO &fl

@ 2.444,34.286 GET m.slct ;
	PICTURE "@&T" ;
	POPUP fldpop ;
	SIZE 21.923,32.800 ;
	DEFAULT " " ;
	FONT "MS Sans Serif", 8 ;
	STYLE ""

_CUROBJ = OBJNUM(chk)

*********************************************************

FUNCTION _r9d0yugvh     &&  m.squit_btn VALID
#REGION 1
CLEAR READ

*********************************************************

FUNCTION _r9d0yuh2s     &&  m.sl VALID
#REGION 1
IF EMPTY(varary[sl,3])
*	vfld = varary[sl,2]
*	DEFINE POPUP fldpop FROM 0,0 TO 5,34.5 PROMPT FIELDS &vfld SCROLL
*	@ 2.444,34.286 GET m.slct ;
*		PICTURE "@&T" ;
*		POPUP fldpop ;
*		SIZE 21.923,32.800 ;
*		DEFAULT " " ;
*		FONT "MS Sans Serif", 8 ;
*		STYLE ""
	_CUROBJ = objary[sl]
ELSE
	vfld = varary[sl,2]+"."+varary[sl,3]
	
	DEFINE POPUP fldpop FROM 0,0 TO 5,34.5 PROMPT FIELDS &vfld SCROLL

	@ 2.444,34.286 GET m.slct ;
 		PICTURE "@&T" ;
		POPUP fldpop ;
		SIZE 21.923,32.800 ;
		DEFAULT " " ;
		FONT "MS Sans Serif", 8 ;
		STYLE ""
	READ	
	tmpary[sl] = m.slct

	@ 1.444+sl*3,12.714 SAY tmpary[sl] ;
		SIZE 1.000,25.800 ;
		FONT "MS Sans Serif", 8 ;
		COLOR ,RGB(,,,255,255,255)

	vfld = varary[sl,2]
	tfld = varary[sl,2]+".serial"
	ttfld = &tfld
	SET FILTER TO &vfld = ttfld

	vfld = varary[1,2]
	DEFINE POPUP fldpop FROM 0,0 TO 5,34.5 PROMPT FIELDS &vfld SCROLL

	@ 2.444,34.286 GET m.slct ;
		PICTURE "@&T" ;
		POPUP fldpop ;
		SIZE 21.923,32.800 ;
		DEFAULT " " ;
		FONT "MS Sans Serif", 8 ;
		STYLE ""
	_CUROBJ = OBJNUM(chk)

ENDIF	

*********************************************************

FUNCTION qtyfltr
IF chk = 1
	SELECT *;
		 FROM &adbf;
		 HAVING Protab.qty > 0;
		 INTO CURSOR Query
ELSE
	SELECT *;
		 FROM &adbf;
		 INTO CURSOR Query
ENDIF

FOR lp = 2 TO anbr
	IF !EMPTY(varary[lp,3])
		relfld = 'query.'+ varary[lp,2]
		SET RELATION TO &relfld INTO varary[lp,2] ADDITIVE
	ENDIF
ENDFOR		
		
vfld = varary[1,2]
DEFINE POPUP fldpop FROM 0,0 TO 5,34.5 PROMPT FIELDS &vfld SCROLL

@ 2.444,34.286 GET m.slct ;
	PICTURE "@&T" ;
	POPUP fldpop ;
	SIZE 21.923,32.800 ;
	DEFAULT " " ;
	FONT "MS Sans Serif", 8 ;
	STYLE ""

*********************************************************

FUNCTION brws
SET FILTER TO &fl
BROW FIELDS &brwlst
*SET RELATION TO

*********************************************************

FUNCTION asmbl
SET RELATION TO
vfld = varary[1,2]
fl = "&varary[1,2] = ALLTRIM(tmpary[1])"
FOR lp = 2 TO anbr
	vfld = varary[lp,2]
	tfld = IIF(EMPTY(varary[lp,3]), ALLTRIM(tmpary[lp]), varary[lp,2]+".serial")
	IF !EMPTY(ALLTRIM(tmpary[lp]))
		fldary[lp] = &tfld
*		fl = fl + " AND &vfld = &tfld"
		fl = fl + " AND &vfld = fldary["+STR(lp,1)+"]"
	ENDIF	
ENDFOR
	
vfld = varary[1,2]
DEFINE POPUP fldpop FROM 0,0 TO 5,34.5 PROMPT FIELDS &vfld SCROLL

SET FILTER TO &fl
GO TOP 

@ 2.444,34.286 GET m.slct ;
	PICTURE "@&T" ;
	POPUP fldpop ;
	DEFAULT ' ' ;
	SIZE 21.923,32.800 ;
	FONT "MS Sans Serif", 8 ;
	STYLE ""

READ

FOR lp = 2 TO anbr
	IF !EMPTY(varary[lp,3])
		relfld = 'query.'+ varary[lp,2]
		SET RELATION TO &relfld INTO varary[lp,2] ADDITIVE
	ENDIF
ENDFOR		
		
*********************************************************

FUNCTION _r9d0yuh9h     && Read Level Show
PRIVATE currwind
STORE WOUTPUT() TO currwind
*
* Show Code from screen: SEARCH3
*
#REGION 1
IF SYS(2016) = "WSEARCH" OR SYS(2016) = "*"
	ACTIVATE WINDOW wsearch SAME
ENDIF
IF NOT EMPTY(currwind)
	ACTIVATE WINDOW (currwind) SAME
ENDIF