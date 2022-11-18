SET TALK    OFF
SET SAFETY  OFF
SET ESCAPE  OFF
SET CENTURY OFF
SET HELP    OFF
SET DELETED ON
SET CENTURY ON
SET DATE FRENCH
SET HOURS TO 24
SET STATUS BAR ON
SET CLOCK STATUS
CLEAR MACROS
CLEAR
CLEAR ALL
CLOSE DATA
SET PATH TO c:\stock\prg;c:\stock\bmp;c:\stock\rpt;c:\stock\scr;c:\stock\dbf;c:\stock\projects
SET PROC TO proc
ON ERROR DO err_prg WITH ERROR(),MESSAGE(),MESSAGE(1),PROGRAM(),LINENO()

PUBLIC m.publang,add,notedit,full,pubmenu,m.drec,pubcur1,pubcur2,pubfld00,pubfld0,pubfld1,pubfld2,pubfld3,pubfld4,pubfld5,pubfld6,pubcon,pubcom,comchannel
PUBLIC m.pubprog1,m.pubprog2,m.pubprog3,m.pubprog4,m.pubprog5,m.pubver,m.pubdir
m.pubprog1 = .F.
m.pubprog2 = .F.
m.pubprog3 = .F.
m.pubprog4 = .F.
m.pubprog5 = .F.
comchannel = 0
*LANGUAGE
m.publang  = .T.
*Currency
m.pubcur1  = 'DHS'
m.pubcur2  = 'USD'
*Fields
m.pubfld00 = IIF(m.publang,'Season','Saison')
m.pubfld0  = IIF(m.publang,'Category','Categorie')
m.pubfld1  = 'Theme'
m.pubfld2  = IIF(m.publang,'Model','Modele')
m.pubfld3  = IIF(m.publang,'Family','Famille')
m.pubfld4  = IIF(m.publang,'Sub-family','Sous-famille')
m.pubfld5  = IIF(m.publang,'Color','Couleur')
m.pubfld6  = IIF(m.publang,'Size','Taille')
m.pubcon   = IIF(m.publang,'Children','Enfants')
*Communication
m.pubcom   = .T.
*Menu
m.pubmenu  = .F.

IF !USED('genpar')
	SELECT 0
	USE genpar
ELSE
	SELECT genpar
ENDIF
m.pubdir = dir
IF USED('genpar')
	SELECT genpar
	USE
ENDIF	

PUBLIC EDITS,TOPFILE,ENDFILE,BRTITLE,NOLOCK,ECANCEL,DELREC,INVAGE
PUBLIC NOFEAT,NOWIZ,MAKEREPO,NOREPO,DELNOTE,READONLY
PUBLIC NOTABLE,BADEXPR,LOCWIZ,MULTITABLE,DBFEMPTY
IF m.publang
   EDITS	  = 'Please finish your edits.' 
   TOPFILE	  = 'Top of file.'
   ENDFILE	  = 'End of file.'
   BRTITLE	  = 'Locate Record'
   FIEMPTY    = 'Empty File'
   NOLOCK	  = 'Sorry, could not lock record -- try again later.' 
   ECANCEL	  = 'Edits Canceled.'
   DELREC	  = 'Delete selected record?'
   NOFEAT	  = 'Feature not available yet.'
   NOWIZ	  = 'Wizard application is not available.'
   MAKEREPO	  = 'Creating report with Report Wizard.'
   NOREPO	  = 'Could not create report.'
   DELNOTE 	  = 'Deleting records...'
   READONLY   = 'Table is read-only. No editing allowed.'
   NOTABLE 	  = 'No table selected. Open table or run query.'
   BADEXPR	  = 'Invalid expression.'
   LOCWIZ	  = 'Locate WIZARD.APP:'
   MULTITABLE = 'You have multiple related tables. Adding records in not allowed.'
   DBFEMPTY	  = 'The file is empty, add a new record?'
ELSE   
   EDITS	  = 'S.V.P. terminez la saisie' 
   TOPFILE	  = 'Début du fichier'
   ENDFILE	  = 'Fin du fichier'
   BRTITLE	  = 'Chercher un enregistrment'
   FIEMPTY    = 'Fichier Vide'
   NOLOCK	  = "Desolé, l'enregistrement n'est pas disponible--Essayé plus tard"
   ECANCEL	  = 'Saisie éliminée'
   DELREC	  = "Effacer l'enregistrement selectionné?"
   NOFEAT	  = 'Feature not available yet'
   NOWIZ	  = 'Application non trouvé'
   MAKEREPO	  = "Creation d'un rapport"
   NOREPO     = "Rapport non crée"
   DELNOTE 	  = 'Effacer les enregistrements ...'
   READONLY   = 'Table ouverte en lecture seulement'
   NOTABLE 	  = 'Aucune table selectionée'
   BADEXPR	  = 'Expression invalide'
   LOCWIZ	  = 'Locate WIZARD.APP:'
   MULTITABLE = 'Ajout des enregistrements est invalide'
   DBFEMPTY	  = 'Le fichier est vide, ajouter un nouveau?'
ENDIF

PUBLIC isediting,isadding,issearch
PRIVATE isreadonly,is2table

m.isediting = .F.
m.add       = .F.
m.full      = .F.
m.mores     = .T.
m.isadding  = .F.
m.is2table  = .F.
m.issearch  = .F.
m.tempcurs  = SYS(2015)  &&used if General field

m.isreadonly = IIF(ISREAD(),.T.,.F.)
IF m.isreadonly
	WAIT WINDOW C_READONLY TIMEOUT 1
ENDIF

IF !USED('config')
	SELECT 0
	USE config
ELSE
	SELECT config
ENDIF
m.pubver = ver
IF USED('config')
	SELECT config
	USE
ENDIF	
MODIFY WINDOW SCREEN TITLE "CHIP - Lebanon" ICON FILE "..\bmp\stock.ico"
DEFINE WINDOW hidecmd FROM 1,1 TO 2,2
ACTIVATE WINDOW command IN WINDOW hidecmd
ACTIVATE SCREEN
@ 1,3 SAY 'logo.bmp' BITMAP SIZE 45,75 ISOMETRIC
@ 7,3 SAY 'logo.bmp' BITMAP SIZE 45,75 ISOMETRIC
@ 13,3 SAY 'logo.bmp' BITMAP SIZE 45,75 ISOMETRIC
@ 19,3 SAY 'logo.bmp' BITMAP SIZE 45,75 ISOMETRIC
@ 25,3 SAY 'logo.bmp' BITMAP SIZE 45,75 ISOMETRIC
@ 33,50 SAY 'STOCK MANAGEMENT '+m.pubver FONT 'ZapfHumnst BT',12 STYLE 'B' COLOR RGB(0,0,225,255,255,255)

IF m.pubcom
   DO opendde
   FOR i=1 TO 30000
   ENDFOR
   DO opencom
   FOR i=1 TO 30000
   ENDFOR
   ! /N2 C:\STOCK\PROJECTS\com9600.exe
ENDIF

DO stock.mpr