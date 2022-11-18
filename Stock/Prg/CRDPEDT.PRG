= checkzapfile("usr")
k = 0
SELECT labtmp
SET ORDER TO field2
GO TOP
DO WHILE !EOF() 
	SELECT usr
	APPEND BLANK
	k = k+1
	REPLACE number     WITH ALLTRIM(labtmp.number)
	REPLACE descriptio WITH labtmp.name
	REPLACE rayon      WITH ALLTRIM(famtab1.number)
	REPLACE theme      WITH ALLTRIM(protab.field1)
	REPLACE modele     WITH ALLTRIM(protab.field2)
	REPLACE coloris    WITH ALLTRIM(fl3tab.name)
	REPLACE taille     WITH ALLTRIM(fl4tab.name)
	REPLACE status     WITH ALLTRIM(STR(labtmp.qty))
	REPLACE usprice    WITH ALLTRIM(STR(labtmp.aprice))
	REPLACE Llprice    WITH ALLTRIM(STR(labtmp.pprice))
	REPLACE des        WITH '0'+ALLTRIM(number)
	SELECT labtmp
	SKIP
ENDDO	
SELECT usr
SUM VAL(status) TO cnttag

KEYBOARD CHR(13)
m.st    = '1 '
m.nbprx = lay.nbprx
m.nbpry = lay.nbpry

IF m.nbprx+m.nbpry>1
	DEFINE WINDOW prt_win FROM 0,0 TO 8,30 ;
		TITLE "Printing Process" ;
		FONT "MS Sans Serif", 8 STYLE 'B';
		FLOAT ;
		CLOSE ;
		NOMINIMIZE ;
		DOUBLE
	MOVE WINDOW prt_win CENTER
	stloc = 1
ENDIF	

DIMENSION rctary[nbprx]

IF USED("usr")
	SELECT usr
ELSE
	SELECT 0
	USE usr
ENDIF

IF !EOF()
	arec = RECNO()
ELSE
	arec = 1
ENDIF
GO TOP
frstcl = IIF(MOD(stloc,nbprx)=0,nbprx,MOD(stloc,nbprx))
frstrw = INT((stloc-1)/nbprx)+1
lstcl  = frstcl-1

m.mrg  = lay.x0+lay.x2*(frstcl-1)
m.row  = lay.y0+lay.y2*(frstrw-1)
m.stpx = lay.x2
m.stpy = lay.y2

DO WHILE cnttag<>0
	pauok = pauprt()
	IF pauok=2 OR LASTKEY()=27
		SELECT usr
		IF arec=1
			GO TOP
		ELSE	
			GO arec
		ENDIF	
		SET PRINT TO
		RELEASE WINDOW prt_win
		RETURN
	ENDIF
	m.modele = usr.modele	
	FOR lpc=frstrw TO nbpry
		FOR lpr=frstcl TO nbprx
			LOCATE FOR NOT EMPTY(status)
			IF FOUND()
				IF !m.divpage
					m.modele = usr.modele
				ENDIF			
				IF m.modele=usr.modele
					rctary[lpr] = RECNO()
					REPLACE status WITH IIF(VAL(status)=1,'  ',ALLTRIM(STR(VAL(status)-1)))
					lstcl  = lstcl+1
					cnttag = cnttag-1
				ELSE
					EXIT
				ENDIF
			ENDIF
			IF cnttag=0
				EXIT
			ENDIF
		ENDFOR
		DO prtpedt WITH row,mrg,frstcl,lstcl
		SET DEVICE TO SCREEN
		
		FOR lpr=frstcl TO lstcl
			SELECT usr
			GO rctary[lpr]
			REPLACE usr.copies WITH STR(VAL(usr.copies)+1,2)
		ENDFOR
		
 		frstcl = 1
		lstcl  = 0

		IF cnttag=0
			EXIT
		ENDIF		
		m.row = m.row+m.stpy
		m.mrg = lay.x0	
	ENDFOR

	SET PRINT TO
	m.modele = usr.modele
	frstrw   = 1
	m.row    = lay.y0
	IF cnttag=0
		EXIT
	ENDIF
ENDDO
SELECT usr
GO arec
RELEASE WINDOW prt_win

FUNCTION pauprt
***************
IF WVISIBLE("prt_win")
	ACTIVATE WINDOW prt_win SAME
ELSE
	ACTIVATE WINDOW prt_win NOSHOW
ENDIF
CLEAR
@ 1.0,7.600 GET m.cnt ;
	PICTURE "@*HNT \<Continue;\<Abort" ;
	SIZE 1.769,10.500,2.667 ;
	DEFAULT 1 ;
	FONT "MS Sans Serif", 8 ;
	STYLE "B"
@ 3,7.600 SAY "Remaining Prints: " + STR(cnttag,5)+' ' ;
	FONT "MS Sans Serif", 8 ;
	STYLE "B"
	
IF NOT WVISIBLE("prt_win")
	ACTIVATE WINDOW prt_win
ENDIF

IF cal_par.ppause='Paused'
	READ
ELSE
	SHOW GET m.cnt DISABLE
	m.cnt = 1
	IF LASTKEY()=27
		m.cnt = 2
	ENDIF
ENDIF	
ACTIVATE WINDOW prt_win NOSHOW
RETURN m.cnt