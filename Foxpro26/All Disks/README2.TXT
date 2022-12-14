                     README.TXT
**********************************************************
                   Release Notes for
 
    Microsoft(R) FoxPro(R) for Windows(TM) Version 2.6

    (C)1994 Microsoft Corporation. All Rights Reserved.

This README provides information not contained in the
FoxPro online Help file or print documentation, and provides
corrections to both. Information in Help is more current than
information in the print documentation. 
**********************************************************

=========
CONTENTS
=========

Part   Description
----   -----------
 1     New Features in FoxPro 2.6
 2     New and Enhanced Commands and Functions
 3     Notes about the Catalog Manager
 4     Notes about Wizards
 5     Macro Key Support in FoxPro 2.6
 6     FoxPro and Other Applications
 7     International and Cross-Platform Recommendations
 8     Installation Questions and Answers
 9     Product Support Questions and Answers
----------------------------------------------------------

===================================
Part 1: New Features in FoxPro 2.6
===================================

Microsoft FoxPro version 2.6 includes the following
new features:

   *  A Catalog Manager interface that provides convenient 
      access to database files and functions, and enables you
      to convert dBASE(R) files easily.
	
   *  A collection of wizards that simplify common database 
      tasks.

   *  Language additions that provide dBASE IV(R) language 
      compatibility.

FoxPro version 2.6 includes existing print
documentation from version 2.5, as well as significant 
updates to the Help file that describe new 2.6
functionality. For more information, see the What's New 
Since FoxPro 2.5 Help topic.
----------------------------------------------------------


================================================
Part 2: New and Enhanced Commands and Functions 
================================================
Several commands and functions are new to FoxPro 2.6, 
and some have been enhanced to provide additional 
compatibility with dBASE IV(R). For additional information 
about these commands and functions, see the New and Enhanced 
Commands and Functions Help topic.

New/Enhanced Commands
---------------------
ON BAR
ON EXIT BAR
ON EXIT MENU
ON EXIT PAD
ON EXIT POPUP
ON MENU
ON MOUSE
ON PAD
ON POPUP
SET CPCOMPILE
SET CPDIALOG
SET COLLATE
SET DIRECTORY
SET FIELDS
SET FORMAT
SET KEY
SET NOCPTRANS
SET VIEW

New/Enhanced Functions
----------------------
BARCOUNT( )
BARPROMPT( )
BLANK
CPCURRENT()
CPDBF()
CPCONVERT()
DESCENDING( )
FLDCOUNT( )
FOR( )
FTIME( )
HOME( )
ID( )
IDXCOLLATE()
ISBLANK( )
ISMOUSE( )
KEYMATCH( )
MEMORY(0)
MEMORY(1)
MEMORY(2)
MEMORY(3)
MEMORY(4)
MEMORY(5)
MEMORY(6)
MEMORY(7)
PADPROMPT( )
PCOUNT( )
PRTINFO( )
RUN( )
SET("ATTRIBUTES")
SET("DIRECTORY")
SET("DISPLAY")
SET("LIBRARY")
SET MBLOCK
SET("MBLOCK")
SYS(1037)
TAG( )
TAGCOUNT( )
TAGNO( )
UNIQUE( )
WINDOW( )

New/Enhanced System Memory Variables
------------------------------------
_ASSIST
_SHELL
----------------------------------------------------------
 
========================================
Part 3: Notes about the Catalog Manager
========================================

Reserved Names in the Catalog Manager
-------------------------------------
Like any other FoxPro application, the Catalog Manager calls 
and runs procedures and programs whose names may interfere 
with other user-defined procedures and programs. 

When you execute a program or procedure, FoxPro searches for
the program or procedure in a specific order.  This order 
includes the current file as well as programs that were 
executed prior to the current one, including the Catalog 
Manager itself.

Below is a list of procedure and program names that can cause
errors if run from within the Catalog Manager.

ABORT
ACTTHERM
ADDALIAS
ADDBS
ADDFILE
ADDTASK
ANALYZE
CLEANUP
CONTROL
DELUSER
DUMMY
ERRHAND
ESCHAND
EXECUTE
EXPORT
FILEEXT
FINDFILE
FORCEEXT
GETSQL
GETSTAMP
GETWPARM
IMPORT
INDEX
ISMAC
ISSPR
JUSTEXT
JUSTFNAME
JUSTPATH
MENUPROC
MIGRATE
MODIFY
MSGBOX
ORDER
PARSFIL
PRINT
REMOVE
SAVEPREF
SETTINGS
SETTOPIC
SHUTDOWN
TOUCHED
TRASHCAN
UPDTHERM
VRUN 


API Library Routines and the Catalog Manager
--------------------------------------------
If you use any routines that are called when an API library 
is opened or closed, you should note the following Catalog 
Manager behaviors.

The Catalog Manager opens and closes API libraries on
these occasions:

* When the Catalog Manager is started, all API 
  libraries are closed.

* When a program is run from within the Catalog
  Manager, all API libraries that were open before
  the Catalog Manager started are opened.  When the program
  ends the API libraries are closed.

* When the Catalog Manager is exited, all API 
  libraries that were open before the Catalog
  Manager started are opened.
----------------------------------------------------------


============================
Part 4: Notes about Wizards
============================

Printing with the Report and Label Wizards
------------------------------------------
To use a Report or a Label Wizard in Windows you must have
a printer installed, otherwise a "Printer not ready" error 
occurs.  To install a printer, select Print Setup from the
File Menu and choose your printer.

Printing Labels with a Laser Printer
------------------------------------
The Label Wizard is designed to print labels that use a
whole sheet of paper, which can go over the printable margins
allowed by Laser printers.  If you use a Laser printer and
your labels print over the printable page you must use the 
Report Writer to modify your label to print only on the 
printable page.  To restrict the printing area of labels 
follow these steps:

1. Modify your label using the Report Writer.
2. From the Report menu, select Page Layout.
3. In the Print Area section of the Page Layout dialog, 
   select the Printable Page radio button.

If you want to modify the label further, you must use the 
Report Writer, not the Label Wizard.

Memo and General Fields and the Label Wizard
--------------------------------------------
To add Memo and General fields you must modify the label 
with the Report Writer. Memo and General fields cannot be 
included in a label created with the Label Wizard.  

Modifying Wizard-Created Items in Microsoft Windows NT(TM)
----------------------------------------------------------
If you are using the Microsoft Windows NT operating system, 
you can use wizards to create tables, queries, 
screens, reports, or labels. To modify items created with a 
wizard you must use the appropriate FoxPro power tool, such
as the Screen Builder, Report Writer or RQBE.
----------------------------------------------------------


========================================
Part 5: Macro Key Support in FoxPro 2.6
========================================
In FoxPro 2.6, the key combinations Alt+1 to Alt+9 cannot
be used to define macros. For a list of key combinations
supported in FoxPro 2.6 see the ON KEY LABEL topic in Help.
----------------------------------------------------------


======================================
Part 6: FoxPro and Other Applications
======================================

Importing Microsoft Excel Spreadsheets to FoxPro
----------------------------------------------------
FoxPro can import spreadsheet files from Microsoft 
Excel version 4.0. To import spreadsheets from Microsoft 
Excel version 5.0, you must first save them as version 
4.0 files. Columns from the spreadsheet become fields 
in the table, and rows from the spreadsheet become records
in the table.

Running dBASE Applications in FoxPro
------------------------------------
To run dBASE III and IV applications in FoxPro, first you 
must convert the individual components, and then re-compile 
them in FoxPro.

To convert dBASE III and dBASE IV files use the FoxPro 
Catalog Manager. For information on converting dBASE files 
in the Catalog Manager, see the Switching from dBASE Help 
topic. 

You can also use the Convert dBASE Files dialog to convert
more than one file at a time.  To bring up this dialog,
select Convert dBASE files... from the Run menu.
----------------------------------------------------------


=========================================================
Part 7: International and Cross-Platform Recommendations 
========================================================= 
See the International and Cross-Platform Features topic in
Help for the latest information on code pages and collation
sequences.

Tagging Codepages to Tables in the Catalog Manager
--------------------------------------------------
Note:  This section only applies if you are using code
pages for cross-platform development.

To ensure that accented characters in a table are displayed 
correctly on the current platform, the table must be tagged 
with a code page. In Catalog Manager, to tag a code page to
a table that does not have one, make sure the following 
are true:

1.  SET EXCLUSIVE is ON.  The exclusive setting can 
    be set in the Catalog Manager Settings dialog.
2.  The configuration file (CONFIG.FPW) includes the
    line CODEPAGE = AUTO

When you try to use the table without a code page in the
Catalog Manager a dialog appears, allowing you to specify 
the platform on which the file was created.  Once a table 
is tagged with a code page you do not have to tag it again. 
If you do not specify a code page, the table is displayed 
with the code page of the current platform.  To remove a 
tag from a table, run CPZERO.PRG, located in the 
FoxPro root directory.
----------------------------------------------------------


===========================================
Part 8: Installation Questions and Answers
===========================================

Overview of Questions

1   What is the easiest way to install FoxPro on a network?
2   What kind of user rights do I need to install FoxPro 
    on the network?
3   Why are workstation users on our network unable to use
    the graphing and spell checker features in FoxPro?
4   During installation I was asked to select either
    DOS-style keystrokes or Windows-style keystrokes. 
    Which should I choose?
5   What is the minimum DOS files setting for FoxPro for 
    Windows?
6   Does FoxPro for Windows update any Dynamic Link 
    Libraries (DLLs) in the Windows subdirectory where 
    system files are kept?
7   I accidentally deleted some of the FoxPro sample files.
    Can I reinstall these files without reinstalling 
    FoxPro? 
8   My computer is having difficulty reading the FoxPro 
    disks because of an alignment problem with the floppy 
    drives. Is there another option for installation?
9   I'm trying to install FoxPro for Windows under Windows 3.0
    and I'm getting the message "Cannot write VER.DLL."  
    Any suggestions?
10  I'm having difficulty installing FoxPro. I've tried
    streamlining both my CONFIG.SYS and AUTOEXEC.BAT and I
    still can't install FoxPro for Windows. What should I 
    do?
11  Can I put a common WIN.INI file up on the server?
12  What is the best way to edit the WIN.INI file?
13  If I modify my WIN.INI file, is it advisable to 
    restart Windows?
14  Can I delete the ~MSSTFQF.T directory?
15  I get the error message "Insufficient memory" when
    trying to load FoxPro for Windows. What setting needs
    to be modified?
16  My computer is running with 4 MB of RAM. How large a
    swap file do I need?
17  When using my LANtastic network card, why do I always
    get a memory conflict when loading FoxPro for Windows?
18  When using a Paradise driver, I get a memory conflict
    when trying to load FoxPro for Windows. What should I
    do?
19  When I use QEMM386 memory manager with FoxPro for
    Windows, I get a memory conflict. What should I do?
20  When I use the Mach32 video card with my Gateway 2000,
    I get a General Protection Fault in module MACH32.DRV
    at 0003:0E22. What should I do?
21  I understand that a number of video drivers don't work
    properly with FoxPro for Windows. Why is this and what
    steps should I take to resolve the situation?
22  I occasionally experience problems when trying to open
    more than 15 tables at one time. What could be wrong?
23  I'm running Windows for Workgroups. Do I need to load
    the MS-DOS SHARE program?
24  Where does FoxPro create temporary files?
25  I frequently use the RUN command to change drives and
    directories within an application. Why isn't this
    working in FoxPro for Windows?
26  Why doesn't the value returned by SYS(2010) match the
    FILES setting in the CONFIG.SYS file?
27  What memory allocation options are available with the
    Windows PIF Editor settings?
----------------------------------------------------------

Question 1

What is the easiest way to install FoxPro on a network?

Answer 1

In a network environment you must first install FoxPro on
the server and then set up each workstation. Installation
on the server is called an Administrative Setup, for which
you use the /A installation switch. This option
decompresses all FoxPro for Windows files and Setup
files. FoxPro for Windows files will be read-only on the
network drive.

To perform an Administrative Setup, choose Run... from the
Windows Program Manager File menu. When the Run dialog 
appears, type A:\SETUP /A or B:\SETUP /A (depending on the
floppy drive you are using) in the Command Line box. Press
Enter and proceed with the installation. You will be
prompted for your user information and an installation
directory, default drive, path name, and so on. For more 
detailed information refer to "Instructions for 
Administrative Setup" in the FoxPro for Windows
Installation and Configuration manual.

Installation on individual workstations is called
Workstation Setup, for which you use the /N installation
switch. Workstation Setup requires a previous 
Administrative Setup. The Workstation Setup copies FoxPro 
initialization files, updates some common .DLL files, and 
creates a Program Manager group and an icon on each 
workstation to run a shared network version of FoxPro 
for Windows.

You can perform the workstation setup in two ways: from 
the network or from disk. To install from a network drive, open
the Windows File Manager and then open a window with the
correct network path for FoxPro for Windows and double-
click SETUP.EXE. Choose the Workstation Installation 
button in the Installation Type dialog box. To install 
from disk, type A:\SETUP /N or B:\SETUP /N in the Run 
dialog Command Line box, then press Enter. For more 
detailed information refer to "Instructions for 
Workstation Setup" in the FoxPro for Windows
Installation and Configuration manual.
----------------------------------------------------------

Question 2

What kind of user rights do I need to install FoxPro on 
the network?

Answer 2

You need full system rights to install on the network.
Contact your system administrator for further information
about your specific installation.
----------------------------------------------------------

Question 3

Why are workstation users on our network unable to use the
graphing and spell checker features in FoxPro?

Answer 3

Workstation users need access to the MSGRAPH and PROOF
subdirectories in order to use these features. Contact
your system administrator regarding user access to these
subdirectories.
----------------------------------------------------------

Question 4

During installation I was asked to select either DOS-style
keystrokes or Windows-style keystrokes. Which should I
choose?

Answer 4

If you aren't sure which style to use, select Windows-
style keystrokes. Windows-style keystrokes correspond to 
standard Windows keystrokes. MS-DOS-style keystrokes cause 
FoxPro for Windows to behave like FoxPro for MS-DOS. 
Later, if you want to use MS-DOS-style keystrokes, add the 
line KEYCOMP = DOS to your CONFIG.FPW file. For more 
information about the styles of keystrokes, see the SET 
KEYCOMP command in the FoxPro Language Reference manual 
or Help.
----------------------------------------------------------

Question 5

What is the minimum DOS files setting for FoxPro for 
Windows?

Answer 5

The DOS FILES command should be set to 40 or higher. 
Buffers should be set to at least 25, or to 10 if you are 
using a disk cache. During installation, Setup checks your 
CONFIG.SYS file. If your FILES setting is less than 40, a 
dialog gives you three choices: have Setup make the 
modifications for you, have Setup make the 
modifications after you have reviewed and edited the 
changes, or allow you to make the modifications later. 
If you allow Setup to modify your CONFIG.SYS file, your 
old file is saved as CONFIG.OLD or as CONFIG.001 if 
CONFIG.OLD already exists. For more information, see 
"Customizing FoxPro" in the Installation and Configuration 
manual.
----------------------------------------------------------

Question 6

Does FoxPro for Windows update any Dynamic Link Libraries
(DLL files) in the Windows subdirectory where system files 
are kept?

Answer 6

Yes. The following DLL files are updated: COMMDLG.DLL,
DDEML.DLL, OLECLI.DLL, OLESVR.DLL, SHELL.DLL, TOOLHELP.DLL
and VER.DLL.

The FoxPro graphing capabilities require that GRAPH.EXE
and GRAPH.HLP be stored in the Windows subdirectory
\MSAPPS\MSGRAPH.

FoxPro spell checking capabilities require that
MSSPELL.DLL, MSSP_AM.LEX, and MSSP_BR.LEX be
stored in the Windows subdirectory \MSAPPS\PROOF.
----------------------------------------------------------

Question 7

I accidentally deleted some of the FoxPro sample files.
Can I reinstall these files without reinstalling FoxPro?

Answer 7

Yes. Run Setup again and select the Custom Installation
option from the Installation Type dialog. With this option
box, you choose which components to install.
----------------------------------------------------------

Question 8

My computer is having difficulty reading the FoxPro disks
because of an alignment problem with the floppy drives. 
Is there another option for installation?

Answer 8

Try copying all the files to a directory on your hard disk
and then install from the hard disk.
----------------------------------------------------------

Question 9

I'm trying to install FoxPro for Windows under Windows 3.0
and I'm getting the message "Cannot write VER.DLL." Any
suggestions?

Answer 9

Rename the VER.DLL file in the Windows System subdirectory.
----------------------------------------------------------

Question 10

I'm having difficulty installing FoxPro. I've tried
streamlining both my CONFIG.SYS and AUTOEXEC.BAT and I
still can't install FoxPro for Windows. What should I do?

Answer 10

First of all, don't try to install from within a third-
party replacement shell, such as Norton Desktop. Make 
sure you use the Windows Program Manager shell. To specify 
the Program Manager shell, change the SYSTEM.INI line

SHELL = <some other shell> 

to 

SHELL = PROGMAN.EXE

Next, make certain you have disabled all TSRs, including
anti-virus software. Check the WIN.INI file for the
LOAD = and RUN = lines. Comment these lines out using a
semicolon (;).
----------------------------------------------------------

Question 11

Can I put a common WIN.INI file up on the server?

Answer 11

No. Each workstation has its own WIN.INI file.
----------------------------------------------------------

Question 12

What is the best way to edit the WIN.INI file?

Answer 12

You can edit the WIN.INI file with Notepad or EDIT.EXE in
MS-DOS 5.0 or greater.
----------------------------------------------------------

Question 13

If I modify my WIN.INI file, is it advisable to restart
Windows?

Answer 13

Yes. As a general rule, always restart Windows when you
have modified the WIN.INI file. There are some
exceptions, but if you are trouble-shooting you want to
be certain that all changes are reflected in the current
environment.
----------------------------------------------------------

Question 14

Can I delete the ~MSSTFQF.T directory?

Answer 14

Yes. The ~MSSTFQF.T directory is created during setup.
Ordinarily this directory is deleted at the end of the
installation process.
----------------------------------------------------------

Question 15

I get the error message "Insufficient memory" when trying
to load FoxPro for Windows. What setting needs to be
modified?

Answer 15

Add a MEMLIMIT statement to your CONFIG.FPW file. The
syntax for this statement is:

MEMLIMIT = %, min, max

The % argument is a calculated amount based on available
memory and min and max represent actual minimum and maximum
allocations.

For example, the following statement allocates 80 percent
of your system's available memory to FoxPro with a minimum
allocation of 1024K and a maximum of 4096K:

MEMLIMIT = 80, 1024, 4096.

----------------------------------------------------------

Question 16

My computer is running with 4 MB of RAM. How large a swap
file do I need?

Answer 16

Use a 2 MB permanent swap file for virtual memory and
don't include a MEMLIMIT setting in the CONFIG.FPW file.
If you are running multiple programs simultaneously, it
may be necessary to specify a swap file larger than 2 MB.
Let FoxPro run with the default setting. In addition, do
not use a RAM drive or an expanded memory manager.
Expanded memory isn't used by FoxPro for Windows. A disk
cache should be no larger than 256K.
----------------------------------------------------------

Question 17

When using my LANtastic network card, I always get a 
memory conflict when loading FoxPro for Windows. What 
should I do?

Answer 17

Try using the following line in the SYSTEM.INI file in
the Windows directory:

EMMExclude = D800-DFFF

----------------------------------------------------------

Question 18

When using a Paradise driver, I get a memory conflict when
trying to load FoxPro for Windows. What should I do?

Answer 18

Try using the following line in the SYSTEM.INI file in
the Windows directory:

EMMExclude = C600-C800

----------------------------------------------------------

Question 19

When I use QEMM386 memory manager with FoxPro for Windows,
I get a memory conflict. What should I do?

Answer 19

Try using the following line in the SYSTEM.INI file in
the Windows directory:

EMMExclude = C0FF-C7FF

----------------------------------------------------------

Question 20

When I use the Mach32 video card with my Gateway 2000, I
get a General Protection Fault in module MACH32.DRV at
0003:0E22. What should I do?

Answer 20

You are experiencing a problem with your video driver.
Contact your video card vendor for the most recent version
of the driver. In the meantime, you can use the VGA.DRV
or the SUPERVGA.DRV drivers that come with Windows.
----------------------------------------------------------

Question 21

I understand that a number of video drivers don't work
properly with FoxPro for Windows. Why is this and what
steps should I take to resolve the situation?

Answer 21

FoxPro for Windows makes intensive use of screen I/O. If 
you are experiencing General Protection Faults (GPFs) or 
display problems (borders, status bar, fonts, unusual 
alignment when previewing reports) you may be experiencing 
a problem with your video driver. To test this situation, 
load either the VGA or VGA version 3.0 drivers and go 
through the steps to reproduce the earlier problem. If the 
problem no longer exists when you are using a standard 
Windows driver, it's very likely that the old driver was 
the cause of the problem. Contact the vendor for an 
updated version of the driver or download the Windows 
Driver Library (WDL) on the Microsoft Download Service 
(MSDL). Complete instructions for obtaining the Windows 
Driver Library is contained in the Product Support section 
of this file.

We have identified problems with the following video 
cards:

   ATI Ultra Graphics Pro (local bus) -- caused GPF, fixed
     with update
   AST Power Premium 486 w/VGA on Motherboard -- fixed
     with update
   Genoa 7900 MM card
   Dell 486/20 with VGA on Motherboard
   Compaq QVision
   Speed Star 24x
   Trident 8900C
   TIGA 1024 x 768 x 256 small fonts
   NEC 2A
   MACH32

Contact your video card vendor for the most recent version
of the driver for your card. In the meantime, you can use
the VGA.DRV or the SUPERVGA.DRV drivers that come with
Windows.
----------------------------------------------------------

Question 22

I occasionally experience problems when trying to open
more than 15 tables at one time. What could be wrong?

Answer 22

Make certain you have the correct SHARE statement in your
AUTOEXEC.BAT file:

   SHARE /F:4096 /L:500
----------------------------------------------------------

Question 23

I'm running Windows for Workgroups. Do I need to load the
MS-DOS SHARE program?

Answer 23

No. Windows for Workgroups automatically loads VSHARE,
which takes the place of SHARE.
----------------------------------------------------------

Question 24

Where does FoxPro create its temporary files?

Answer 24

FoxPro creates its temporary files in the same directory
where Windows stores its temporary files, unless you
specifically designate an alternate location by including
one or more of the EDITWORK, SORTWORK, PROGWORK and/or
TMPFILES statements in the CONFIG.FPW file. The location
of Windows .TMP files is specified with the MS-DOS set
variable TEMP. In most cases, you should let FoxPro write
.TMP files to the same location. However, if this
location is a RAM drive, make sure that enough space is
available for both FoxPro and Windows .TMP files.
----------------------------------------------------------

Question 25

I frequently use the RUN command to change drives and
directories within an application. Why isn't this working
in FoxPro for Windows?

Answer 25

Each time you issue the RUN command to change drives and
directories, Windows opens a new session of MS-DOS,
otherwise known as a virtual machine, and executes the
command within the newly created session. When you exit
from this session or shut down the virtual machine, you
are right back where you started without having changed
drives. You should use FoxPro's own drive and directory
commands instead of using the RUN command. For example,
using SET DEFAULT TO C:\NEWDIR will actually issue the
MS-DOS command CHDIR.
----------------------------------------------------------

Question 26

Why doesn't the value returned by SYS(2010) match the
FILES setting in the CONFIG.SYS file?

Answer 26

Windows interacts with the MS-DOS open file tables in such
a way that the value reported in the MS-DOS box and in
Windows may be different.
----------------------------------------------------------

Question 27

What memory allocation options are available with the
Windows PIF Editor settings?

Answer 27

In the PIF Editor "memory requirements" section the
"KB Required" and "KB Desired" may both be set to "-1".
This tells Windows to give all the available memory it
can to an MS-DOS application. This doesn't mean that the
MS-DOS application will have more than 600K, but it will
get about as much conventional memory as you had available
before starting Windows (minus some conventional memory
that Windows itself occupies).

Changing the "EMS Memory" and "XMS Memory" sections
tells Windows to provide or not provide the MS-DOS
application and EMS (Expanded Memory Services) or XMS
(Extended Memory Services). This won't, however, allow
an application that doesn't support accessing EMS or 
XMS to do so. Only applications that have the capability
will be able to access the additional memory provided. It
is important to note that once you tell Windows to create
EMS or XMS for an application, it does so whether the
application can use it or not.

In the "Execution" section, if Exclusive is selected, no
other application, be it Windows or MS-DOS, will run while
the application with Exclusive is the current application.

In the "Advanced Options" there is a "Multitasking 
Options" section. This section is very important when you 
are trying to get an MS-DOS application to run in the
background. If the MS-DOS application is very I/O
intensive (reads and writes to disk, asynchronous
communication, and do on) it will be necessary to increase 
the "Background Priority." Unfortunately there is no rule 
of thumb when it comes to the exact number that should be 
used. It's important to note that the larger the number 
for "Background Priority" the slower other applications 
will run while the MS-DOS application with the high 
priority is running in the background.
----------------------------------------------------------


==============================================
Part 9: Product Support Questions and Answers
==============================================

Overview of Questions
---------------------

1   Where can I go for additional pre-sale information 
    on Microsoft Fox products, such as pricing, upgrade 
    policies, and release dates? 
2   How do I download the Windows Driver Library?
3   Where can I go for additional FoxPro product support?
4   If I call for support or post a message on the Fox 
    Forum what information should I include?
5   If I am deaf or hard of hearing where can I go for 
    additional FoxPro product support?
6   I have some specific comments and suggestions on 
    Microsoft Fox products. Where can I direct these 
    comments?
7   How do I use @ ... SAY Commands and Control Codes 
    to Print in FP/W.
----------------------------------------------------------

Question 1

Where can I go for additional pre-sale information on 
Microsoft Fox products, such as pricing, upgrade policies, 
and release dates?

Answer 1

Call 800-426-9400.
----------------------------------------------------------

Question 2

How do I download the Windows Driver Library?

Answer 2

Microsoft provides the Windows Driver Library (WDL) on the 
Microsoft Download Service (MSDL) from which users can 
obtain drivers and articles directly via modem. If you 
have a modem, the drivers are available at no charge on 
CompuServe, GEnie, and Microsoft OnLine, as well as the 
MSDL. 

The phone number for the MSDL is 206-936-6735. There is no 
connect-time charge for using this service. However, any 
standard connect-time fees and long-distance telephone 
charges do apply during downloading. After you connect to 
one of these services, please read the WDL.TXT file for 
downloading instructions and a complete list of files 
available on the Download Service.

On CompuServe, the WDL is in the Microsoft Software 
Library, which you can reach by typing GO MSL at the 
command prompt. (To find WDL.TXT, search on S13354.)

On the GEnie and Microsoft OnLine systems, the WDL is in 
the Microsoft Software Library. (To find WDL.TXT, search 
on S13354.)

If you don't have a modem, you can obtain the WDL on disk 
from Microsoft End-User Sales. To order the WDL on disk, 
call 800-426-9400, Monday through Friday, 6:00 a.m.- 5:30 
p.m., Pacific time. There is a $20 fee for ordering the 
WDL on disk.
----------------------------------------------------------

Question 3

Where can I go for additional FoxPro product support?

Answer 3

Call *Microsoft FastTips* to hear recorded responses to 
common questions about Microsoft FoxPro. You can also 
order technical notes that are sent to your fax 
machine. FastTips is available seven days a week, 24 
hours a day. For assistance with FoxPro, dial 
206-635-7190.

Use the *Microsoft Forums* on CompuServe. You can discuss 
current support and design issues with other developers 
and download sample files from the libraries. For an 
introductory CompuServe membership kit specifically for 
Microsoft users, dial 800-848-8199 and ask for operator 
170. If you are already a CompuServe member, type 
GO MICROSOFT at any ! prompt. For specific FoxPro 
information, type GO FOXFORUM, or GO FOXUSER.

Use the *Microsoft Download Service* to access technical 
notes and supplementary files covering common Microsoft 
product support issues. This service is available via 
modem 24 hours a day, seven days a week at 206-936-6735.

Call *Microsoft Product Support Services* between 6 a.m. 
and 6 p.m. Pacific time, Monday through Friday. For 
assistance with Microsoft Fox products for MS-DOS or for 
Windows, call 206-635-7191. For assistance with Microsoft 
Fox products for the Macintosh, call 206-635-7192.

Microsoft support services are subject to Microsoft prices,
terms, and conditions that are in place at the time the 
service is used.
----------------------------------------------------------

Question 4

If I call for support or post a message on the Fox Forum, 
what information should I include?

Answer 4

When you call, you should be at your computer and have the 
appropriate product documentation at hand. Be prepared to 
give the following information:

* The Product Identification Number for your product.  
  You can find this number by choosing About FoxPro... from 
  the Help menu.  The Product ID is displayed on the About 
  FoxPro dialog. The Product ID is a serial number that will 
  be used to auto-route your call to the appropriate support 
  person, to meter any free support you are entitled to, and 
  to display your prior support history.
* The version number of the Microsoft Fox product that you 
  are using. This can be obtained by typing ?VERSION(1) in 
  the Command window for FoxPro.
* The type of hardware you are using, including network 
  hardware if applicable; and the operating environment 
  that you are using.
* The contents of the following files: CONFIG.FPW or 
  CONFIG.SYS, AUTOEXEC.BAT, and any network 
  configuration files. 
* The exact wording of any messages that appeared on your 
  screen.
* A description of what happened and what you were doing 
  when the problem occurred.
* Can you reproduce the problem using a small number of 
  steps?
* What steps have you already taken to solve the problem?
---------------------------------------------------------

Question 5

If I am deaf or hard of hearing where can I go for 
additional FoxPro product support?

Answer 5

Microsoft Product Support Services is available for the 
hearing impaired. Using a special TDD/TT modem, dial 
206-635-4948 between 6 a.m. and 6 p.m. Pacific time, 
Monday through Friday.

Microsoft support services are subject to Microsoft 
prices, terms, and conditions that are in place at the 
time the service is used.
----------------------------------------------------------

Question 6

I have some specific comments and suggestions on Microsoft 
Fox products. Where can I direct these comments?

Answer 6

Help make future versions of Microsoft Fox products even 
better by calling, sending a fax, sending a note, or 
sending electronic mail. Your feedback will be considered 
in developing future versions of Microsoft FoxPro.

Call 206-936-9474 and leave your comments using a 
touch-tone phone 24 hours a day.

Fax 206-936-7329 and provide detailed comments 24 
hours a day.

Send a note to:

Microsoft FoxPro WishLine
One Microsoft Way
Redmond, WA 98052

If you have access to the Internet, send electronic mail 
at:

   foxwish@microsoft.com.

Or, address CompuServe mail to:

   >INTERNET:foxwish@microsoft.com.

Note that charges apply to Internet if you use CompuServe.
----------------------------------------------------------
                   =================
                   End of README.TXT
                   =================


