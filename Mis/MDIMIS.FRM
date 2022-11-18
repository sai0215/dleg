VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.MDIForm MDIMIS 
   BackColor       =   &H8000000C&
   Caption         =   "MIS Booking & Accounting Software"
   ClientHeight    =   7995
   ClientLeft      =   -15
   ClientTop       =   600
   ClientWidth     =   11850
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   11790
      TabIndex        =   1
      Top             =   0
      Width           =   11850
      Begin VB.CommandButton cmdBo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Booking Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   0
         Picture         =   "MDIMIS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   2895
      End
      Begin VB.CommandButton cmdRep 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   6000
         Picture         =   "MDIMIS.frx":0190
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   3015
      End
      Begin VB.CommandButton cmdFA 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Financial Accounting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   2880
         Picture         =   "MDIMIS.frx":0290
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   3135
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   9000
         Picture         =   "MDIMIS.frx":06D2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6495
      Left            =   0
      ScaleHeight     =   6435
      ScaleWidth      =   11790
      TabIndex        =   0
      Top             =   2055
      Width           =   11850
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   6495
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   11895
         ExtentX         =   20981
         ExtentY         =   11456
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Menu mnubo 
      Caption         =   "&Booking Order"
      Begin VB.Menu mnugen 
         Caption         =   "&General Data"
         Begin VB.Menu mnuGenDeal 
            Caption         =   "&Deal"
         End
         Begin VB.Menu mnudm 
            Caption         =   "Deal M&odification"
         End
         Begin VB.Menu mnuBudget 
            Caption         =   "&Budget"
         End
         Begin VB.Menu mnuAgency 
            Caption         =   "&Agency"
         End
         Begin VB.Menu mnuclient 
            Caption         =   "&Cliient"
         End
         Begin VB.Menu mnumedia 
            Caption         =   "&Media type"
         End
         Begin VB.Menu mnuProduct 
            Caption         =   "&Product"
         End
         Begin VB.Menu mnumaterial 
            Caption         =   "Ma&terial"
         End
      End
      Begin VB.Menu mnuborates 
         Caption         =   "&Rates"
         Begin VB.Menu mnucinemarates 
            Caption         =   "&Cinema"
         End
         Begin VB.Menu mnuCnn 
            Caption         =   "C&NN"
         End
         Begin VB.Menu mnuMagazinerates 
            Caption         =   "&Magazine"
         End
      End
      Begin VB.Menu mnubonew 
         Caption         =   "&New Entry"
         Begin VB.Menu mnuNewCinema 
            Caption         =   "&Cinema"
         End
         Begin VB.Menu mnuNewMagazine 
            Caption         =   "&Magazine"
         End
         Begin VB.Menu mnuNewTelevision 
            Caption         =   "&Television"
         End
         Begin VB.Menu mnuNewOnline 
            Caption         =   "&Online"
         End
      End
      Begin VB.Menu mnubomod 
         Caption         =   "&Modify"
         Begin VB.Menu mnuCinemaModification 
            Caption         =   "&Cinema"
         End
         Begin VB.Menu mnuMagazineModification 
            Caption         =   "&Magazine"
         End
         Begin VB.Menu mnuTelevisionModification 
            Caption         =   "&Television"
         End
         Begin VB.Menu mnuOnlineMod 
            Caption         =   "&Online"
         End
      End
      Begin VB.Menu mnubodel 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnufa 
      Caption         =   "&Financial Accounting"
      Begin VB.Menu mnufagd 
         Caption         =   "&General Data"
         Begin VB.Menu mnufagdac 
            Caption         =   "&Chart of Accounts"
         End
         Begin VB.Menu mnufagdbanks 
            Caption         =   "&Banks"
         End
         Begin VB.Menu mnufagdsupp 
            Caption         =   "&Supplier"
         End
         Begin VB.Menu mnuremarks 
            Caption         =   "&Remarks"
         End
      End
      Begin VB.Menu mnufatr 
         Caption         =   "&Transactions"
         Begin VB.Menu mnufatrcash 
            Caption         =   "&Cash"
            Begin VB.Menu mnufatrcashpay 
               Caption         =   "&Payment"
               Begin VB.Menu mnufatrcashpaynew 
                  Caption         =   "&New"
               End
               Begin VB.Menu mnufatrcashmod 
                  Caption         =   "&Modify"
               End
            End
            Begin VB.Menu mnufatrcashrpt 
               Caption         =   "&Receipt"
               Begin VB.Menu mnufatrcashrecpnew 
                  Caption         =   "&New"
               End
               Begin VB.Menu mnufatrcashrecpmod 
                  Caption         =   "&Modify"
               End
            End
         End
         Begin VB.Menu mnufatrbank 
            Caption         =   "&Bank"
            Begin VB.Menu mnufatrbankpay 
               Caption         =   "&Payment"
               Begin VB.Menu mnufatrbankpaynew 
                  Caption         =   "&New"
               End
               Begin VB.Menu mnufatrbankpaymod 
                  Caption         =   "&Modify"
               End
            End
            Begin VB.Menu mnufatrbankrecp 
               Caption         =   "&Receipt"
               Begin VB.Menu mnufatrbankrecpnew 
                  Caption         =   "&New"
               End
               Begin VB.Menu mnufatrbankrecpmod 
                  Caption         =   "&Modify"
               End
            End
         End
         Begin VB.Menu mnufatrpdcrpt 
            Caption         =   "PDC &Receipt"
            Begin VB.Menu mnufatrpdcreptnew 
               Caption         =   "&New"
            End
            Begin VB.Menu mnufatrpdcrecpmod 
               Caption         =   "&Modify"
            End
            Begin VB.Menu mnufatrpdcrptpost 
               Caption         =   "&Posting"
            End
         End
         Begin VB.Menu mnufatrPdcIss 
            Caption         =   "PDC &Issue"
            Begin VB.Menu mnufatrpdcissnew 
               Caption         =   "&New"
            End
            Begin VB.Menu mnufatrpdcissmod 
               Caption         =   "&Modify"
            End
            Begin VB.Menu mnufatrPdcIssposting 
               Caption         =   "&Posting"
            End
         End
         Begin VB.Menu mnufatrjv 
            Caption         =   "&Journal Voucher"
            Begin VB.Menu mnufatrjvnew 
               Caption         =   "&New"
            End
            Begin VB.Menu mnufatrjvprnt 
               Caption         =   "&Print"
            End
         End
         Begin VB.Menu mnufatrcn 
            Caption         =   "C&reditNote"
            Begin VB.Menu mnufatrcnnew 
               Caption         =   "&New"
            End
            Begin VB.Menu mnufatrcnmod 
               Caption         =   "&Modify"
            End
            Begin VB.Menu mnufatracnprnt 
               Caption         =   "&Print"
            End
         End
         Begin VB.Menu mnufatradn 
            Caption         =   "&Debit Note"
            Begin VB.Menu mnufatrdnnew 
               Caption         =   "&New"
            End
            Begin VB.Menu mnufatrdnmod 
               Caption         =   "&Modify"
            End
            Begin VB.Menu mnufatrdnprnt 
               Caption         =   "&Print"
            End
         End
         Begin VB.Menu mnufatrpur 
            Caption         =   "&Purchase"
            Begin VB.Menu mnufatrpurcash 
               Caption         =   "&Cash"
               Begin VB.Menu mnucpaddition 
                  Caption         =   "&Addition"
               End
               Begin VB.Menu mnucpmodification 
                  Caption         =   "&Modification"
               End
            End
            Begin VB.Menu mnufatrpurcredit 
               Caption         =   "C&redit"
               Begin VB.Menu mnucrpaddition 
                  Caption         =   "&Addition"
               End
               Begin VB.Menu mnucrpmodification 
                  Caption         =   "&Modification"
               End
            End
         End
         Begin VB.Menu mnufatrdatafreez 
            Caption         =   "&Data Freezing"
         End
      End
   End
   Begin VB.Menu mnurep 
      Caption         =   "&Reports"
      Begin VB.Menu mnuborep 
         Caption         =   "&Booking Orders"
         Begin VB.Menu mnuborepAGTO 
            Caption         =   "&Agency /Turnover"
         End
         Begin VB.Menu mnuborepagclto 
            Caption         =   "&Client/Turnover"
         End
         Begin VB.Menu mnuAgClPr 
            Caption         =   "Agency-Client-Product / &TO"
         End
         Begin VB.Menu mnurepdeal 
            Caption         =   "&Deal Monitoring"
         End
         Begin VB.Menu mnuFlatPlan 
            Caption         =   "&Flat Plan"
         End
         Begin VB.Menu mnupaging 
            Caption         =   "&Paging"
         End
         Begin VB.Menu mnuCinemaDtls 
            Caption         =   "&Cinema Screen TO"
         End
         Begin VB.Menu mnucinscreenmon 
            Caption         =   "Cinema Screen Monitoring"
         End
      End
      Begin VB.Menu mnufarep 
         Caption         =   "&Accounting Statements"
         Begin VB.Menu mnufarepcs 
            Caption         =   "Cash Statement - DHS"
         End
         Begin VB.Menu mnucashstatementusd 
            Caption         =   "Cash Statement - &USD"
         End
         Begin VB.Menu mnubankstmt 
            Caption         =   "Bank Statement"
         End
         Begin VB.Menu mnupdcrecd 
            Caption         =   "PDC &Received"
            Begin VB.Menu mnupdcrecdagncy 
               Caption         =   "&Agency "
            End
            Begin VB.Menu mnupdcrecdDate 
               Caption         =   "&Date "
            End
         End
         Begin VB.Menu mnupdciss 
            Caption         =   "PDC &Issued"
            Begin VB.Menu mnupdcisssup 
               Caption         =   "&Supplier "
            End
            Begin VB.Menu mnupdcissdate 
               Caption         =   "&Date "
            End
         End
         Begin VB.Menu mnumonthinv 
            Caption         =   "Print I&nvoices"
         End
         Begin VB.Menu mnuinvlst 
            Caption         =   "&Invoice List"
            Begin VB.Menu mnuinvagncy 
               Caption         =   "&Agency "
            End
            Begin VB.Menu mnuInvLstRepmedia 
               Caption         =   "&Media"
            End
            Begin VB.Menu mnuinvdate 
               Caption         =   "&Date"
            End
         End
         Begin VB.Menu mnuSaleAnalysis 
            Caption         =   "&Sales Analysis"
         End
         Begin VB.Menu mnusagross 
            Caption         =   "Sales &Analysis Gross"
         End
         Begin VB.Menu mnuSaleAnalysisSubMedia 
            Caption         =   "Sales Analysis - Sub &Media"
         End
         Begin VB.Menu mnuagncystmt 
            Caption         =   "&Agency Statment"
            Begin VB.Menu mnuagnstmtdhs 
               Caption         =   "&DHS"
            End
            Begin VB.Menu mnuagnstmtUSD 
               Caption         =   "&USD"
            End
         End
         Begin VB.Menu mnuttlagncyos 
            Caption         =   "&Total Agency Out Standing"
            Begin VB.Menu mnuTotAgencydhs 
               Caption         =   "&DHS"
            End
            Begin VB.Menu mnuTotAgencyUSD 
               Caption         =   "&USD"
            End
         End
         Begin VB.Menu mnusuppstmt 
            Caption         =   "&Supplier Stament"
         End
         Begin VB.Menu mnutotsuppos 
            Caption         =   "Total S&upplier out standing"
         End
         Begin VB.Menu mnuaccstmt 
            Caption         =   "A&ccount statment"
         End
         Begin VB.Menu mnutb 
            Caption         =   "T&rail Balance"
         End
         Begin VB.Menu mnufinstmnts 
            Caption         =   "F&inanacial Statements"
         End
      End
      Begin VB.Menu mnurepclose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "MDIMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdFA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then PopupMenu mnufa
End Sub
Private Sub cmdgen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then PopupMenu mnugen
End Sub
Private Sub cmdBo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then PopupMenu mnubo
End Sub
Private Sub cmdRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Z = 1 Then
  mnurep.Enabled = True
  If Button = vbLeftButton Then PopupMenu mnurep
Else
  mnurep.Enabled = False
  If Button = vbLeftButton Then frmLoginreports.Show
End If
  
End Sub
Private Sub MDIForm_Load()
mnurep.Enabled = False
WebBrowser1.Navigate App.Path & "\UDAYMA.HTML"
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
If Cancel = 1 Then End
End Sub
Private Sub mnuaccstmt_Click()
frmAcctReport.Show
End Sub

Private Sub mnuAgClPr_Click()
frmAgencyClientPrdTo.Show
End Sub

Private Sub mnuagency_Click()
frmAgency.Show
End Sub
Private Sub mnuagnstmtdhs_Click()
frmCustomerRepdhs.Show
End Sub
Private Sub mnuagnstmtUSD_Click()
frmCustomerRep.Show
End Sub
Private Sub mnubankstmt_Click()
frmBankBalRep.Show
End Sub
Private Sub mnuborepagclto_Click()
frmClientTorev.Show
End Sub
Private Sub mnuborepAGTO_Click()
frmAgencyTorev.Show
End Sub
Private Sub mnuCashstmt_Click()
frmCashBalRep.Show
End Sub
Private Sub mnuBudget_Click()
frmagencyBudget.Show
End Sub
Private Sub mnucashstatementusd_Click()
frmCashBalRepusd.Show
End Sub
Private Sub mnuCinemaDtls_Click()
frmCinRep.Show
End Sub
Private Sub mnuCinemaModification_Click()
frmBOModCin.Show
End Sub
Private Sub mnucinemarates_Click()
frmCinemaRates.Show
End Sub
Private Sub mnucinscreenmon_Click()
frmCinRepScMonRevOriginal.Show
End Sub
Private Sub mnuClient_Click()
frmClient.Show
End Sub
Private Sub mnuCnn_Click()
frmcnnrates.Show
End Sub

Private Sub mnucpaddition_Click()
frmCashPurchaseAddition.Show
End Sub
Private Sub mnucpmodification_Click()
frmCashPurchaseModifcation.Show
End Sub
Private Sub mnucrpaddition_Click()
frmCreditPurchaseAddition.Show
End Sub
Private Sub mnucrpmodification_Click()
frmCreditPurchaseModifcation.Show
End Sub

Private Sub mnudm_Click()
frmDealMod.Show
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnufagdac_Click()
frmacctmas.Show
End Sub

Private Sub mnufagdbanks_Click()
frmBankMas.Show
End Sub

Private Sub mnufagdsupp_Click()
frmSuppmas.Show
End Sub

Private Sub mnufarepcs_Click()
frmCashBalRep.Show
End Sub

Private Sub mnufatracnprnt_Click()
frmCrntPrint.Show

End Sub

Private Sub mnufatrbankpaymod_Click()
frmBankPaymentModification.Show
End Sub

Private Sub mnufatrbankpaynew_Click()
frmBankPaymentAddition.Show
End Sub

Private Sub mnufatrbankrecpmod_Click()
frmBankReceiptModification.Show
End Sub

Private Sub mnufatrbankrecpnew_Click()
frmBankReceiptAddition.Show
End Sub
Private Sub mnufatrcashmod_Click()
 frmCashPaymentModification.Show
End Sub
Private Sub mnufatrcashpaynew_Click()
 frmCashPaymentAddition.Show
End Sub
Private Sub mnufatrcashrecpmod_Click()
 frmCashReceiptModification.Show
End Sub
Private Sub mnufatrcashrecpnew_Click()
 frmCashReceiptAddition.Show
End Sub
Private Sub mnufatrcnmod_Click()
 frmCreditNoteMod.Show
End Sub
Private Sub mnufatrcnnew_Click()
 frmCreditNoteAdd.Show
End Sub
Private Sub mnufatrdatafreez_Click()
 frmPosting.Show
End Sub
Private Sub mnufatrdnmod_Click()
frmDebitNoteMod.Show
End Sub
Private Sub mnufatrdnnew_Click()
frmDebitNoteAdd.Show
End Sub

Private Sub mnufatrdnprnt_Click()
frmDBNTPrint.Show
End Sub

Private Sub mnufatrjvnew_Click()
frmJrnlAdd.Show
End Sub

Private Sub mnufatrjvprnt_Click()
frmJrnlPrint.Show
End Sub

Private Sub mnufatrpdcissmod_Click()
 frmPDCPaymentModification1.Show
End Sub
Private Sub mnufatrpdcissnew_Click()
 frmPDCPaymentAddition1.Show
End Sub
Private Sub mnufatrPdcIssposting_Click()
 frmPdcPaymentPosting.Show
End Sub
Private Sub mnufatrpdcrecpmod_Click()
 frmPdcReceiptModification1.Show
End Sub
Private Sub mnufatrpdcreptnew_Click()
 frmPdcReceiptAddition1.Show
End Sub
Private Sub mnufatrpdcrptpost_Click()
 frmPdcRecPosting1.Show
End Sub

Private Sub mnufinstmnts_Click()
frmFinStmnts.Show
End Sub

Private Sub mnuFlatPlan_Click()
frmFlatPlan.Show
End Sub

Private Sub mnuGenDeal_Click()
frmDeal.Show
End Sub

Private Sub mnuinvagncy_Click()
frmSaleRepCust.Show
End Sub

Private Sub mnuinvdate_Click()
frmSaleRepdt.Show
End Sub

Private Sub mnuInvLstRepmedia_Click()
frmSaleRepMedia.Show
End Sub

Private Sub mnuMagazineModification_Click()
frmBOModMag.Show
End Sub

Private Sub mnuMagazinerates_Click()
frmMagRates.Show
End Sub

Private Sub mnuMaterial_Click()
 frmmaterial.Show
End Sub
Private Sub mnuMedia_Click()
 frmMediaType.Show
End Sub

Private Sub mnumonthinv_Click()
frmInvoicing.Show
End Sub

Private Sub mnuNewCinema_Click()
frmBOCin.Show
End Sub

Private Sub mnuNewMagazine_Click()
frmBOMag.Show
End Sub

Private Sub mnuNewOnline_Click()
frmBOONL.Show
End Sub

Private Sub mnuNewTelevision_Click()
frmBOTEL.Show
End Sub
Private Sub mnuOnlineMod_Click()
frmBOModol.Show
End Sub
Private Sub mnupaging_Click()
frmPaging.Show
End Sub
Private Sub mnupdcissdate_Click()
frmPdcPmtRepDt.Show
End Sub

Private Sub mnupdcisssup_Click()
 frmPdcPaySupp.Show
End Sub
Private Sub mnupdcrecdagncy_Click()
 frmPdcRptRepCust.Show
End Sub
Private Sub mnupdcrecdDate_Click()
 frmPdcRptRepDt.Show
End Sub
Private Sub mnuProduct_Click()
 frmProducts.Show
End Sub

Private Sub mnuremarks_Click()
frmAcRemarks.Show
End Sub

Private Sub mnurep_Click()
If Z = 1 Then
 mnurep.Enabled = True
Else
 frmLoginreports.Show
End If
End Sub

Private Sub mnurepclose_Click()
 mnurep.Enabled = False
 Z = 0
End Sub
Private Sub mnurepdeal_Click()
frmDealReport.Show
End Sub

Private Sub mnusagross_Click()
frmTOTANALYSISgr.Show
End Sub

Private Sub mnuSaleAnalysis_Click()
frmTOTANALYSIS.Show
End Sub

Private Sub mnuSaleAnalysisSubMedia_Click()
frmSubMediaAnalysis.Show
End Sub

Private Sub mnusuppstmt_Click()
frmSupplierRep.Show
End Sub
Private Sub mnutb_Click()
frmTBReport.Show
End Sub

Private Sub mnuTelevisionModification_Click()
frmBOModTel.Show
End Sub

Private Sub mnuTotAgencydhs_Click()
frmCustomerListRepDhs.Show
End Sub

Private Sub mnuTotAgencyUSD_Click()
frmCustomerListRep.Show
End Sub

Private Sub mnutotsuppos_Click()
frmSupplierListRep.Show
End Sub
