PARAMETERS ypos, xstrt, elmt1, elmt2
ACTIVATE screen
MODIFY WINDOW screen FONT 'foxfont',9

wst   = FONTMETRIC(6,'Times New Roman',40) 
twst  = TXTWIDTH('STRING OF CHARACTERS', 'Times New Roman',40 )*wst
twff  = 17.9 && 'STRING OF CHARACTERS' in cm
cmpxl = twff/twst

PUBLIC cod[60], stp[2], pen[2], char[2], clr[2], strt, stop

IF m.out=1
	ACTIVATE screen
ELSE	
	SET DEVICE TO PRINT
ENDIF

SELECT def
SET ORDER TO y 

COUNT TO cntbcd FOR type='BAR CODE'
DIMENSION bcdary[cntbcd+2,4]

COUNT TO m.count
DIMENSION frmpary[m.count+10,15], mempary[m.count+10,1]
COPY TO ARRAY frmpary FIELDS x, y, font, type, lnth, size, style, algn, active, label, hght, rc, gc, bc, lang
GO TOP
FOR j=1 TO m.count
	SCATTER MEMO MEMVAR
	mempary[j,1] = m.constant
	SKIP
ENDFOR

IF USED("usr")
	SELECT usr
ELSE
	SELECT 0
	USE usr
ENDIF

aadj = cal_par.aadj 
xadj = cal_par.xadj 
yadj = cal_par.yadj 

FOR i = 1 TO m.count
	DO prtfld
ENDFOR

SET DEVICE TO SCREEN

RETURN

FUNCTION prtfld
***************
	IF EMPTY(frmpary[i+1,2])
		nexty = 100
	ELSE	
		nexty = (frmpary[i+1,2]+ypos)* yadj 
	ENDIF
	IF frmpary[i,9] = .F.
		RETURN
	ENDIF

	m.x = frmpary[i,1]
	m.y = frmpary[i,2]
	m.dy = (m.y + ypos)* yadj
	m.font = frmpary[i,3]
	m.size = frmpary[i,6]
	m.style = frmpary[i,7]
	m.algn = frmpary[i,8]
	m.lnth = frmpary[i,5]
	m.hght = frmpary[i,11]
	m.rc = frmpary[i,12]*(cal_par.color-1)
	m.gc = frmpary[i,13]*(cal_par.color-1)
	m.bc = frmpary[i,14]*(cal_par.color-1)
	m.lang = frmpary[i,15]
	m.lenth = frmpary[i,5]
	IF frmpary[i,4] = 'BAR CODE'
		xstp = .25/2
		ystp = .1665/2
		m.ratio = VAL(SUBSTR(mempary[i],LEN(mempary[i])-2,1))
		m.wfact = VAL(SUBSTR(mempary[i],LEN(mempary[i]),1))
		m.bornt = SUBSTR(mempary[i],LEN(mempary[i])-4,1)
		vh = IIF(m.bornt='V',1,0)
		bcdhgt = m.hght*IIF(vh=0,yadj,xadj)  &&,3,4.8)
		stp[1] = IIF(vh=0,xstp,ystp)*wfact
		stp[2] = IIF(vh=0,xstp,ystp)*wfact*ratio
		pen[1] = IIF(vh=0,1.25,.8)*wfact/10
		pen[2] = IIF(vh=0,1.25,.8)*wfact/10*ratio
	ENDIF
  
  xadv = xstrt
  FOR ii = elmt1 TO elmt2

	GO rctary[ii]
	m.dx = (m.x + xadv) *xadj
	swchar = ''
	
	DO CASE
		CASE frmpary[i,4] = "CONSTANT"
			= gprint(mempary[i,1])

		CASE frmpary[i,4] = "VARIABLE"
			= gprint(&frmpary[i,10])

		CASE frmpary[i,4] = "PICTURE"
			i1 = m.dy
			d = 0
			DO CASE
				CASE UPPER(m.algn) = 'I'
					d = m.lnth/2*xadj
				CASE UPPER(m.algn) = 'J'
					d = m.lnth*xadj
			ENDCASE
			i2 = m.dx - d
			i3 = m.hght*yadj
			i4 = m.lnth*xadj
			i5 = 'usr.'+frmpary[i,10]
			ph = ALLTRIM(i5)+'n'
			phi = ALLTRIM(&ph)
			IF EMPTY(phi) OR UPPER(SUBSTR(phi, LEN(phi)-2,3)) <> 'BMP'
				@ i1, i2 SAY &i5 ISOMETRIC SIZE i3,i4 STYLE IIF(m.rc=0,'','T')
			ELSE
				@ i1, i2 SAY &ph BITMAP ISOMETRIC SIZE i3,i4 STYLE IIF(m.rc=0,'','T')
			ENDIF	
					
		CASE frmpary[i,4] = "LOGO"
			i1 = m.dy
			d = 0
			DO CASE
				CASE UPPER(m.algn) = 'I'
					d = m.lnth/2*xadj
				CASE UPPER(m.algn) = 'J'
					d = m.lnth*xadj
			ENDCASE
			i2 = m.dx - d
			i3 = m.hght*yadj
			i4 = m.lnth*xadj
			i5 = mempary[i]
			@ i1,i2 SAY i5 BITMAP ISOMETRIC SIZE i3,i4 STYLE IIF(m.rc=0,'','T')
				
		CASE frmpary[i,4] = "BACK"
			i1 = m.dy
			d = 0
			DO CASE
				CASE UPPER(m.algn) = 'I'
					d = m.lnth/2*xadj
				CASE UPPER(m.algn) = 'J'
					d = m.lnth*xadj
			ENDCASE
			i2 = m.dx - d
			i3 = m.hght*yadj
			i4 = m.lnth*xadj
			@ i1,i2 SAY mempary[i] BITMAP SIZE i3,i4 STRETCH STYLE IIF(m.rc=0,'','T')
					
		CASE frmpary[i,4] = "COUNTER"
			m.strnb = VAL(SUBSTR(mempary[i,1],8,m.lenth))
			m.stpsz = VAL(SUBSTR(mempary[i,1],16+m.lenth,1))
			strpt = ALLTRIM(STR(ROUND((RECNO()-1)*m.stpsz+m.strnb,0),m.lenth))
			tl = 3
			strpt = PADL(strpt,tl,'0')
			= gprint(strpt)
			
		CASE frmpary[i,4] = "MAG" 
			= magstr(&frmpary[i,10], mempary[i,1])
				
		CASE frmpary[i,4] = "BAR CODE 2/5"
			= barcd25(&frmpary[i,10])
			
		CASE frmpary[i,4] = "BAR CODE 39" 
			= barcd39(&frmpary[i,10])

	ENDCASE
	
	xadv = xadv + lay.x2
	frstpass = .F.
	
  ENDFOR

RETURN

FUNCTION gprint
***************
PARAMETERS what

what = ALLTRIM(what)
IF lang='ARABIC' AND VAL(what)>0
	swchar = CHR(0240)
	mf = .9
ELSE
	swchar = ''
	mf = 1
ENDIF

delta = adjst(what)
@ m.dy,m.dx+delta SAY swchar+what FONT ALLTRIM(m.font), m.size STYLE m.style+'t' COLOR RGB(m.rc, m.gc, m.bc,0,0,0)

FUNCTION adjst
**************
PARAMETERS astr

IF UPPER(cal_par.ptr_mark) = 'SCREEN'
	aw = FONTMETRIC(6,ALLTRIM(m.font),m.size,m.style)
	twdth = TXTWIDTH(ALLTRIM(astr),ALLTRIM(m.font),m.size,m.style) *aw *mf *cmpxl *aadj *xadj 
	
	fchr = SUBSTR(astr,1,1)
	lchr = SUBSTR(astr,LEN(astr),1)
	IF fchr = 'Ç'
		twdth = twdth - .05 *aw *mf *cmpxl *aadj *xadj
	ENDIF
	IF swchar = CHR(0240) AND (lchr <> '8' AND lchr <> '7')
		twdth = twdth - .09 *aw *mf *cmpxl *aadj *xadj
	ENDIF
ELSE	
	aw = FONTMETRIC(6,ALLTRIM(m.font),100,m.style)
	twdth = TXTWIDTH(ALLTRIM(astr),ALLTRIM(m.font),100,m.style) *aw *mf *cmpxl *aadj *xadj *1.025
	twdth = twdth*m.size/100

	fchr = SUBSTR(astr,1,1)
	lchr = SUBSTR(astr,LEN(astr),1)
	IF fchr = 'Ç'
		twdth = twdth - .05 *aw *mf *cmpxl *aadj *xadj *1.025
	ENDIF
	IF swchar = CHR(0240) AND (lchr <> '8' AND lchr <> '7')
		twdth = twdth - .09 *aw *mf *cmpxl *aadj *xadj *1.025
	ENDIF
ENDIF

d = 0
DO CASE
	CASE UPPER(m.algn) = 'I'
		d = -twdth/2
	CASE UPPER(m.algn) = 'J'
		d = -twdth
ENDCASE

RETURN d


FUNCTION barcd25
****************
PARAMETERS m.nb

m.str = CHR(33)
FOR ib = 1 TO m.lenth STEP 2
	m.temp = VAL(SUBSTR(m.nb,ib,2))+35
	IF m.temp >= 127
		d = 3
	ELSE
		d=0
	ENDIF		
	m.str = m.str + CHR(m.temp+d)
ENDFOR
m.str = m.str + CHR(34)

bchgt = 12*m.hght/.55
IF UPPER(cal_par.ptr_mark) = 'SCREEN'
	aw = FONTMETRIC(6,'bc-25i',bchgt)
	twdth = TXTWIDTH(ALLTRIM(str),'bc-25i',bchgt) *aw *aadj *xadj *cmpxl
ELSE
	aw = FONTMETRIC(6,'bc-25i',100)
	twdth = TXTWIDTH(ALLTRIM(str),'bc-25i',100) *aw *aadj *xadj *cmpxl *1.025
	twdth = twdth*bchgt/100
ENDIF

d = 0
DO CASE
	CASE UPPER(m.algn) = 'I'
		d = -twdth/2
	CASE UPPER(m.algn) = 'J'
		d = -twdth
ENDCASE
m.bdx = m.dx + d
m.bdy = m.dy 

@ m.bdy,m.bdx SAY m.str FONT'BC-25i', bchgt STYLE 'T'

FUNCTION barcd39
****************
PARAMETER m.nb

m.str = '!'
FOR ib = 1 TO m.lenth 
	m.temp = SUBSTR(m.nb,ib,1)
	m.str = m.str + IIF(m.temp = ' ','#',m.temp)
ENDFOR
m.str = m.str + '!'

bchgt = 12*m.hght/.55
IF UPPER(cal_par.ptr_mark) = 'SCREEN'
	aw = FONTMETRIC(6,'bc-39',bchgt)
	twdth = TXTWIDTH(ALLTRIM(str),'bc-39',bchgt) *aw *aadj *xadj *cmpxl
ELSE
	aw = FONTMETRIC(6,'bc-39',100)
	twdth = TXTWIDTH(ALLTRIM(str),'bc-39',100) *aw *aadj *xadj *cmpxl *1.025
	twdth = twdth*bchgt/100
ENDIF

d = 0
DO CASE
	CASE UPPER(m.algn) = 'I'
		d = -twdth/2
	CASE UPPER(m.algn) = 'J'
		d = -twdth
ENDCASE
m.bdx = m.dx + d
m.bdy = m.dy 

@ m.bdy,m.bdx SAY m.str FONT'BC-39', bchgt STYLE 'T'

FUNCTION magstr
***************
PARAMETER m.nb, m.cons

IF EMPTY(m.nb)
	WAIT WINDOW 'Empty Magnetic code Field' NOWAIT
	RETURN
ENDIF	

mag_tr = SUBSTR(m.cons,AT('/',m.cons)+9,1)
mag_ss = SUBSTR(m.cons,AT('/',m.cons)+11,1)
mag_sp = SUBSTR(m.cons,AT('/',m.cons)+13,1)
mag_tl = SUBSTR(m.cons,AT('/',m.cons)+15,2)
trlr   = PADL('',VAL(mag_tl),'0')
mag_es = SUBSTR(m.cons,AT('/',m.cons)+18,1)
*@ 0,0 SAY '~'+ mag_tr+ mag_ss+ m.nb+ mag_sp+ trlr+ mag_es 
@ 0,0 SAY '~'+ mag_tr+ mag_ss+ m.nb+ mag_es