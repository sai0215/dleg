SET TALK OFF
SET DELETED ON
SET ESCAPE OFF
SET DATE FRENCH
SET PROC TO PROC
SET SAFETY OFF
SET HOURS TO 24

CLEAR
CLEAR ALL
SET PATH TO C:\chip\stock\PRG;C:\chip\stock\BMP;C:\chip\stock\RPT;C:\chip\stock\SCR;C:\chip\stock\DBF;

PUBLIC m.publang,add,notedit,full,pubmenu,m.drec,pubcur1,pubcur2,pubfld00,pubfld0,pubfld1,pubfld2,pubfld3,pubfld4,pubfld5,pubfld6,pubcon
*LANGUAGE
m.publang = .T.
m.pubcur1 = 'L.L.'
m.pubcur2 = 'USD'
m.pubfld00 = IIF(m.publang,'Season','Saison')
m.pubfld0 = IIF(m.publang,'Category','Categorie')
m.pubfld1 = 'Theme'
m.pubfld2 = IIF(m.publang,'Model','Modele')
m.pubfld3 = IIF(m.publang,'Family','Famille')
m.pubfld4 = IIF(m.publang,'Sub-family','Sous-famille')
m.pubfld5 = IIF(m.publang,'Color','Couleur')
m.pubfld6 = IIF(m.publang,'Size','Taille')
m.pubmenu = .F.
m.pubcon = IIF(m.publang,'Children','Enfants')

PUBLIC EDITS,TOPFILE,ENDFILE,BRTITLE,NOLOCK,ECANCEL,DELREC,INVAGE
PUBLIC NOFEAT,NOWIZ,MAKEREPO,NOREPO,DELNOTE,READONLY
PUBLIC NOTABLE,BADEXPR,LOCWIZ,MULTITABLE,DBFEMPTY
IF m.publang
   EDITS		='Please finish your edits.' 
   TOPFILE		='Top of file.'
   ENDFILE		='End of file.'
   BRTITLE		='Locate Record'
   FIEMPTY      ='Empty File'
   NOLOCK		='Sorry, could not lock record -- try again later.' 
   ECANCEL		='Edits Canceled.'
   DELREC		='Delete selected record?'
   NOFEAT		='Feature not available yet.'
   NOWIZ		='Wizard application is not available.'
   MAKEREPO		='Creating report with Report Wizard.'
   NOREPO		='Could not create report.'
   DELNOTE 		='Deleting records...'
   READONLY 	='Table is read-only. No editing allowed.'
   NOTABLE 		='No table selected. Open table or run query.'
   BADEXPR		='Invalid expression.'
   LOCWIZ		='Locate WIZARD.APP:'
   MULTITABLE	='You have multiple related tables. Adding records in not allowed.'
   DBFEMPTY	    ='The file is empty,Add a new record ?'
ELSE   
   EDITS	   ='S.V.P. terminez la saisie' 
   TOPFILE	   ='Début du fichier'
   ENDFILE	   ='Fin du fichier'
   BRTITLE	   ='Chercher un enregistrment'
   FIEMPTY     ='Fichier Vide'
   NOLOCK	   ="Desolé, l'enregistrement n'est pas disponible--Essayé plus tard"
   ECANCEL	   ='Saisie éliminée'
   DELREC	   ="Effacer l'enregistrement selectionné?"
   NOFEAT	   ='Feature not available yet'
   NOWIZ	   ='Application non trouvé'
   MAKEREPO	   ="Creation d'un rapport"
   NOREPO      ="Rapport non crée"
   DELNOTE 	   ='Effacer les enregistrements ...'
   READONLY    ='Table ouverte en lecture seulement'
   NOTABLE 	   ='Aucune table selectionée'
   BADEXPR	   ='Expression invalide'
   LOCWIZ	   ='Locate WIZARD.APP:'
   MULTITABLE  ='Ajout des enregistrements est invalide'
   DBFEMPTY	   ='Le fichier est vide, ajouter un nouveau ?'
ENDIF

PUBLIC isediting,isadding,issearch
PRIVATE isreadonly,is2table

m.isediting=.F.
m.add =.F.
m.full=.F.
m.mores=.T.
m.isadding=.F.
m.is2table = .F.
m.issearch=.F.
m.tempcurs=SYS(2015)  &&used if General field

m.isreadonly=IIF(ISREAD(),.T.,.F.)
IF m.isreadonly
	WAIT WINDOW C_READONLY TIMEOUT 1
ENDIF
CLOSE DATA

DO criteria.SPR


PROCEDURE repord
****************
m.date1=DATE()
m.date2=DATE()
m.nat=1
DO selord.spr
alias='ordhea'
report='ordredt'
report1='ordredt1'
=getselect('ordhea',0,'0')
=getselect('arrhea',0,'ordhea')
=getselect('protab',0,'0')
STORE RECNO() TO m.saverec
SELECT 0
USE prttab  
SET ORDER TO serial
ZAP
SELECT ordhea
GO TOP
DO WHILE !EOF()	
	xval=.T.
	IF date >= date1 AND date <= date2 
		IF nat = 2
			xval=.F.
			SELECT arrhea
			SET ORDER TO ordhea
			SEEK ordhea.serial
			IF !FOUND()
				xval=.T.
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
	ENDIF		
	SELECT ordhea
	SKIP
ENDDO	
IF m.saverec <> 0
	GO m.saverec
ENDIF	
SELECT prttab
DO printall.spr


PROCEDURE prtrec
***************
SELECT prttab
DELETE FOR EMPTY(bidon)
SELECT (alias)
SET RELATION TO serial INTO Prttab 
DO print WITH IIF(m.criteria=1,report,report1),IIF(m.criteria=1,'RECNO()=m.saverec',IIF(m.criteria=2,'FOUND("prttab")','.T.')),IIF(m.print=1,'PREVIEW','TO PRINT')
SELECT prttab
USE
SELECT (alias)
IF m.saverec <> 0
	GO m.saverec
ENDIF	
