VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmAcctReport 
   BackColor       =   &H80000005&
   Caption         =   "Statement of Account"
   ClientHeight    =   8775
   ClientLeft      =   -105
   ClientTop       =   285
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Statement of Account "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   7335
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   8415
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   7320
         Top             =   5160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<<&Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4800
         Picture         =   "frmAcctReport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6360
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3720
         Picture         =   "frmAcctReport.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6360
         Width           =   1095
      End
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         Picture         =   "frmAcctReport.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6360
         Width           =   1095
      End
      Begin VB.ListBox lstAcctCodes 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   3900
         Left            =   600
         TabIndex        =   0
         Top             =   600
         Width           =   7335
      End
      Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   4800
         Width           =   1455
         _Version        =   65541
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   244
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         Text            =   ""
         Mask            =   "##/##/####"
         PromptCharacter =   ""
         BackColor       =   16777215
         ForeColor       =   0
         HighlightColor  =   -2147483640
         Alignment       =   1
      End
      Begin PVMaskEditLib.PVMaskEdit txtdateto 
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   5400
         Width           =   1455
         _Version        =   65541
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   244
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         Text            =   ""
         Mask            =   "##/##/####"
         PromptCharacter =   ""
         BackColor       =   16777215
         ForeColor       =   0
         HighlightColor  =   -2147483640
         Alignment       =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404080&
         BorderWidth     =   2
         X1              =   8400
         X2              =   -120
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Date To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   2640
         TabIndex        =   6
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Date From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   2640
         TabIndex        =   5
         Top             =   4920
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAcctReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim ws As Workspace
Dim db As Database
Dim Opbal As Currency
Dim Opcrptcr As Currency
Dim Opcrptdb As Currency
Dim Opcpmtcr As Currency
Dim Opcpmtdb As Currency
Dim Opbrptcr As Currency
Dim Opbrptdb As Currency
Dim Opbpmtdb As Currency
Dim Opbpmtcr As Currency
Dim Opprptcr As Currency
Dim Opprptdb As Currency
Dim Opppmtcr As Currency
Dim Opppmtdb As Currency
Dim Opjrnldb As Currency
Dim Opjrnlcr As Currency
Dim opdbntcr As Currency
Dim opdbntdb As Currency
Dim Opcrntcr As Currency
Dim Opcrntdb As Currency
Dim Opcrslcr As Currency
Dim Opcrsldb As Currency
Dim Opcrprdb As Currency
Dim OpcaslCr As Currency
Dim Opcasldb As Currency
Dim Opcaprcr As Currency
Dim Opcaprdb As Currency
Dim Ttlopbal As Currency
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim Sqlqry3 As String
Dim Sqlqry4 As String
Dim Sqlqry5 As String
Dim Sqlqry6 As String
Dim Sqlqry7 As String
Dim Sqlqry8 As String
Dim Sqlqry9 As String
Dim Sqlqry10 As String
Dim SQLQRY11 As String
Dim SQLQRY12 As String
Dim Sqlqry13 As String
Dim Sqlqry14 As String
Dim Sqlqry15 As String
Dim sqlqry16 As String
Dim sqlqry17 As String
Dim sqlqry18 As String
Dim sqlqry19 As String
Dim sqlqry20 As String
Dim Sqlqry21 As String
Dim Sqlqry22 As String
Dim Sqlqry23 As String
Dim Sqlqry24 As String
Dim Sqlqry25 As String
Dim Sqlqry26 As String
Dim Sqlqry27 As String
Dim Sqlqry28 As String
Dim Sqlqry29 As String
Dim Sqlqry30 As String
Dim Sqlqry31 As String
Dim Sqlqry32 As String
Dim Sqlqry33 As String
Dim Sqlqry34 As String
Dim Sqlqry35 As String
Dim Sqlqry36 As String
Dim Sqlqry37 As String
Dim Sqlqry38 As String
Dim Sqlqry39 As String
Dim Sqlqry40 As String
Dim Sqlqry41 As String
Dim Sqlqry42 As String
Dim Sqlqry43 As String
Dim Sqlqry44 As String
Dim Sqlqry45 As String
Dim Sqlqry46 As String
Dim Sqlqry47 As String
Dim Sqlqry48 As String
Dim Sqlqry49 As String
Dim Sqlqry50 As String
Dim Sqlqry51 As String
Dim Sqlqry52 As String
Dim Sqlqry53 As String
Dim Sqlqry54 As String
Dim Sqlqry55 As String
Dim Sqlqry56 As String
Dim Sqlqry57 As String
Dim Sqlqry58 As String
Dim Sqlqry59 As String
Dim Sqlqry60 As String
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim rs3 As Recordset
Dim rs4 As Recordset
Dim rs5 As Recordset
Dim rs6 As Recordset
Dim rs7 As Recordset
Dim rs8 As Recordset
Dim rs9 As Recordset
Dim rs10 As Recordset
Dim rs11 As Recordset
Dim rs12 As Recordset
Dim rs13 As Recordset
Dim rs14 As Recordset
Dim rs15 As Recordset
Dim rs16 As Recordset
Dim rs17 As Recordset
Dim rs18 As Recordset
Dim rs19 As Recordset
Dim rs20 As Recordset
Dim rs21 As Recordset
Dim rs22 As Recordset
Dim rs23 As Recordset
Dim rs24 As Recordset
Dim rs25 As Recordset
Dim rs26 As Recordset
Dim rs27 As Recordset
Dim rs28 As Recordset
Dim rs29 As Recordset
Dim rs30 As Recordset
Dim rs31 As Recordset
Dim rs32 As Recordset
Dim rs33 As Recordset
Dim rs34 As Recordset
Dim rs35 As Recordset
Dim rs36 As Recordset
Dim rs37 As Recordset
Dim rs38 As Recordset
Dim rs39 As Recordset
Dim rs40 As Recordset
Dim rs41 As Recordset
Dim rs42 As Recordset
Dim rs43 As Recordset
Dim rs44 As Recordset
Dim rs45 As Recordset
Dim rs46 As Recordset
Dim rs47 As Recordset
Dim rs48 As Recordset
Dim rs49 As Recordset
Dim rs50 As Recordset
Private Sub cmdBack_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub
 
Private Sub cmdClear_Click()
 textclear
End Sub
 Private Sub cmdDisplay_Click()
  If ValidateData = True Then
             
             
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Delete * from ACCREPORT"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
       
Opbal = 0
Opcrptcr = 0
Opcrptdb = 0
Opcpmtcr = 0
Opcpmtdb = 0
Opbrptcr = 0
Opbrptdb = 0
Opbpmtdb = 0
Opbpmtcr = 0
Opprptcr = 0
Opprptdb = 0
Opppmtcr = 0
Opppmtdb = 0
Opjrnldb = 0
Opjrnlcr = 0
opdbntcr = 0
opdbntdb = 0
Opcrntcr = 0
Opcrntdb = 0
Opcrslcr = 0
Opcrsldb = 0
Opcrprdb = 0
OpcaslCr = 0
Opcasldb = 0
Opcaprcr = 0
Opcaprdb = 0
Ttlopbal = 0
       
       ' Op. from Account Master
        Sqlqry = " select * from acct_mas where acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
          MsgBox " Selected Code not found in Account Register"
          Exit Sub
        Else
          rs.MoveFirst
          Opbal = rs!open_bal
        End If
   
        ' Cash Receipt(credit) before from Date
        Sqlqry1 = " select sum(amount) from crpt_tra where TDATE< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Opcrptcr = rs1.Fields(0)
        
        ' Cash Receipt(debit) before from date
        Sqlqry2 = " select sum(ttl_amount) from crpt_mas where TDATE< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Cash_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs2.Fields(0)) = False Then Opcrptdb = rs2.Fields(0)
                     
        ' Cash Payment(debit) before from date
        Sqlqry3 = " select sum(amount) from cpmt_tra where TDATE< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs3 = db.OpenRecordset(Sqlqry3, dbOpenDynaset)
        If IsNull(rs3.Fields(0)) = False Then Opcpmtdb = rs3.Fields(0)
         
        ' Cash Payment(credit) before from date
        Sqlqry4 = " select sum(ttl_Amount) from cpmt_mas where TDATE< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and cash_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs4 = db.OpenRecordset(Sqlqry4, dbOpenDynaset)
        If IsNull(rs4.Fields(0)) = False Then Opcpmtcr = rs4.Fields(0)
                
        ' Bank Receipt(credit) before from date.
        Sqlqry5 = "select sum(amount) from brpt_tra where TDATE< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs5 = db.OpenRecordset(Sqlqry5, dbOpenDynaset)
        If IsNull(rs5.Fields(0)) = False Then Opbrptcr = rs5.Fields(0)
        
        
        ' Bank Receipt(debit) before from date.
        Sqlqry6 = "select Sum(ttl_amount) from brpt_mas where TDATE< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs6 = db.OpenRecordset(Sqlqry6, dbOpenDynaset)
        If IsNull(rs6.Fields(0)) = False Then Opbrptcr = rs6.Fields(0)
        
        ' Bank Payment(debit) before From date
        Sqlqry7 = "select Sum(amount) from bpmt_tra where TDATE< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs7 = db.OpenRecordset(Sqlqry7, dbOpenDynaset)
        If IsNull(rs7.Fields(0)) = False Then Opbpmtdb = rs7.Fields(0)
        
                        
        ' Bank Payment(credit) before From date
        Sqlqry8 = "select sum(ttl_Amount) from bpmt_mas where TDATE< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs8 = db.OpenRecordset(Sqlqry8, dbOpenDynaset)
        If IsNull(rs8.Fields(0)) = False Then Opbpmtcr = rs8.Fields(0)
                
       ' Pdc Receipts (Debit) before From date
        Sqlqry9 = "select sum(amount) from prpt_mas1 where Cheque_dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and NOT isnull(posting_dt) and BANK_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs9 = db.OpenRecordset(Sqlqry9, dbOpenDynaset)
        If IsNull(rs9.Fields(0)) = False Then Opprptdb = rs9.Fields(0)
                   
       ' Pdc Receipts (credit) before From Date
        Sqlqry10 = "Select sum(amount) from Prpt_MAS1 where Cheque_dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_dt) and acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs10 = db.OpenRecordset(Sqlqry10, dbOpenDynaset)
        If IsNull(rs10.Fields(0)) = False Then Opprptcr = rs10.Fields(0)
        
       ' Pdc Payments (Credit) before From Date
        SQLQRY11 = "select sum(ttl_amount) from Ppmt_mas where Cheque_Dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_Dt)"
        Set rs11 = db.OpenRecordset(SQLQRY11, dbOpenDynaset)
        If IsNull(rs11.Fields(0)) = False Then Opppmtcr = rs11.Fields(0)
        
       ' Pdc Payments (Debit) before From Date
        SQLQRY12 = "Select sum(amount) from Prpt_tra where Cheque_Dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_Dt) and acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs12 = db.OpenRecordset(SQLQRY12, dbOpenDynaset)
        If IsNull(rs12.Fields(0)) = False Then Opppmtdb = rs12.Fields(0)
             
       ' Journal Debit Amount before From Date
        Sqlqry13 = "select sum(damount) from Jrnl_tra where TDATE<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "' and dc_code ='D'"
        Set rs13 = db.OpenRecordset(Sqlqry13, dbOpenDynaset)
        If IsNull(rs13.Fields(0)) = False Then Opjrnldb = rs13.Fields(0)
        
                 
       ' Journal Credit Amount before From Date
        Sqlqry14 = "select sum(camount) from Jrnl_tra where TDATE<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "' and dc_code ='C'"
        Set rs14 = db.OpenRecordset(Sqlqry14, dbOpenDynaset)
        If IsNull(rs14.Fields(0)) = False Then Opjrnlcr = rs14.Fields(0)
        
       ' Debit note (credit) before From Date
        Sqlqry15 = "select sum(amount) from debt_mas where TDATE<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs15 = db.OpenRecordset(Sqlqry15, dbOpenDynaset)
        If IsNull(rs15.Fields(0)) = False Then opdbntcr = rs15.Fields(0)
        
                 
       ' Debit note (debit)  before From Date
        sqlqry16 = "select sum(amount) from debt_mas where TDATE<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and cust_no='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs16 = db.OpenRecordset(sqlqry16, dbOpenDynaset)
        If IsNull(rs16.Fields(0)) = False Then opdbntdb = rs16.Fields(0)
        
       ' Credit note (debit) before From Date
        sqlqry17 = "select Sum(amount) from crdt_mas where TDATE<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs17 = db.OpenRecordset(sqlqry17, dbOpenDynaset)
        If IsNull(rs17.Fields(0)) = False Then Opcrntdb = rs17.Fields(0)
        
        
       ' Credit note (credit)  before From Date
        sqlqry18 = "select sum(amount) from crdt_mas where TDATE<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and supp_no='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs18 = db.OpenRecordset(sqlqry18, dbOpenDynaset)
        If IsNull(rs18.Fields(0)) = False Then Opcrntcr = rs18.Fields(0)
        
                  
       ' Credit Sales  before From Date
        sqlqry19 = "select sum(net_amount) from bo_mas where invoice_date<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs19 = db.OpenRecordset(sqlqry19, dbOpenDynaset)
        If IsNull(rs19.Fields(0)) = False Then Opcrslcr = rs19.Fields(0) * convertion
                                    
       ' Credit Purchases  before From Date
        sqlqry20 = "select sum(gamount) from crpr_mas where TDATE<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs20 = db.OpenRecordset(sqlqry20, dbOpenDynaset)
        If IsNull(rs20.Fields(0)) = False Then Opcrprdb = rs20.Fields(0)
        
        ' Cash Purchases(debit)  before From Date
        Sqlqry23 = "select Sum(gamount) from capr_mas where TDATE<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs23 = db.OpenRecordset(Sqlqry23, dbOpenDynaset)
        If IsNull(rs23.Fields(0)) = False Then Opcaprdb = rs23.Fields(0)
        
        ' Cash Purchases(Credit)  before From Date
        Sqlqry24 = "select sum(gamount) from capr_mas where TDATE<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and cash_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs24 = db.OpenRecordset(Sqlqry24, dbOpenDynaset)
        If IsNull(rs24.Fields(0)) = False Then Opcaprcr = rs24.Fields(0)
        
        
        Ttlopbal = Opbal - Opcrptcr + Opcrptdb - Opcpmtcr + Opcpmtdb - Opbrptcr + Opbrptdb + Opbpmtdb - Opbpmtcr + Opprptdb - Opprptcr + Opppmtdb _
                 - Opppmtcr + Opjrnldb - Opjrnlcr + opdbntdb - opdbntcr + Opcrntdb - Opcrntcr - Opcrslcr + Opcrprdb - OpcaslCr + Opcasldb - Opcaprcr + Opcasldb
                
       If Ttlopbal > 0 Then
        Sqlqry25 = "Insert into ACCREPORT values(" & 0 & ",'','" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','Opening Balance'," & Trim(Ttlopbal) & "," & 0 & ")"
        ws.BeginTrans
        db.Execute (Sqlqry25)
        ws.CommitTrans
       ElseIf Ttlopbal < 0 Then
        Sqlqry25 = "Insert into ACCREPORT values(" & 0 & ",'','" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','Opening Balance'," & 0 & "," & Abs(Trim(Ttlopbal)) & ")"
        ws.BeginTrans
        db.Execute (Sqlqry25)
        ws.CommitTrans
       Else
        Sqlqry25 = "Insert into ACCREPORT values(" & 0 & ",'','" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','Opening Balance'," & 0 & "," & 0 & ")"
        ws.BeginTrans
        db.Execute (Sqlqry25)
        ws.CommitTrans
       End If
       
       ' Cash Receipt after From date and before to date
        Sqlqry26 = "Select * from Crpt_Tra where TDATE>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and TDATE<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs26 = db.OpenRecordset(Sqlqry26, dbOpenDynaset)
        If rs26.RecordCount <> 0 Then
         rs26.MoveFirst
         Do Until rs26.EOF
          Sqlqry27 = "Insert into ACCREPORT values('" & rs26!VOUC_NO & "','" & rs26!vouc_type & "','" & Trim(rs26!tDate) & "','" & Trim(rs26!Description) & "'," & 0 & "," & Trim(rs26!Amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry27)
          ws.CommitTrans
          rs26.MoveNext
         Loop
        End If
        
       ' Cash Payment after From date and before to date
        Sqlqry28 = "Select * from Cpmt_Tra where TDATE>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and TDATE<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs27 = db.OpenRecordset(Sqlqry28, dbOpenDynaset)
        If rs27.RecordCount <> 0 Then
         rs27.MoveFirst
         Do Until rs27.EOF
          Sqlqry29 = "Insert into ACCREPORT values('" & rs27!VOUC_NO & "','" & rs27!vouc_type & "','" & Trim(rs27!tDate) & "','" & Trim(rs27!Description) & "'," & Trim(rs27!Amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry29)
          ws.CommitTrans
          rs27.MoveNext
         Loop
        End If
        
        
       ' Bank Receipt after From date and before to date
        Sqlqry30 = "select * from brpt_tra where TDATE>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and TDATE<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and ACCT_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs28 = db.OpenRecordset(Sqlqry30, dbOpenDynaset)
        If rs28.RecordCount <> 0 Then
         rs28.MoveFirst
         Do Until rs28.EOF
          Sqlqry31 = "Insert into ACCREPORT values('" & rs28!VOUC_NO & "','" & rs28!vouc_type & "','" & Trim(rs28!tDate) & "','" & Trim(rs28!Description) & "'," & 0 & "," & Trim(rs28!Amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry31)
          ws.CommitTrans
          rs28.MoveNext
         Loop
        End If
        
        ' Bank Payment after From date and before to date
        Sqlqry32 = "select * from bpmt_tra where TDATE>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and TDATE<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs29 = db.OpenRecordset(Sqlqry32, dbOpenDynaset)
        If rs29.RecordCount <> 0 Then
         rs29.MoveFirst
         Do Until rs29.EOF
          Sqlqry33 = "Insert into ACCREPORT values('" & rs29!VOUC_NO & "','" & rs29!vouc_type & "','" & Trim(rs29!tDate) & "','" & Trim(rs29!Description) & "'," & Trim(rs29!Amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry33)
          ws.CommitTrans
          rs29.MoveNext
         Loop
        End If
        
        ' Pdc Receipts after From date and before To date
        Sqlqry34 = "select * from prpt_mas1 where Cheque_dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Cheque_dt<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(Posting_Dt) and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "' "
        Set rs31 = db.OpenRecordset(Sqlqry34, dbOpenDynaset)
        If rs31.RecordCount <> 0 Then
         rs31.MoveFirst
         Do Until rs31.EOF
            
            Sqlqry35 = "Insert into ACCREPORT values('" & rs31!VOUC_NO & "','" & rs31!vouc_type & "','" & Trim(rs31!Cheque_Dt) & "','" & Trim(rs31!Description) & "'," & 0 & "," & Trim(rs31!Amount) & ")"
             ws.BeginTrans
             db.Execute (Sqlqry35)
             ws.CommitTrans
             rs31.MoveNext
           Loop
         End If
                  
        ' Pdc Payments after From date and before to date
        Sqlqry36 = "select * from Ppmt_mas where Cheque_Dt>=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and TDATE<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and not Isnull(Posting_Dt) "
        Set rs32 = db.OpenRecordset(Sqlqry36, dbOpenDynaset)
        If rs32.RecordCount <> 0 Then
         rs32.MoveFirst
         Do Until rs32.EOF
           Sqlqry37 = "Select * from ppmt_tra where Vouc_no=" & rs32!VOUC_NO & " and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
           Set rs33 = db.OpenRecordset(Sqlqry37, dbOpenDynaset)
             If rs33.RecordCount <> 0 Then
              rs33.MoveFirst
              Do Until rs33.EOF
               Sqlqry38 = "Insert into ACCREPORT values('" & rs33!VOUC_NO & "','" & rs33!vouc_type & "','" & Trim(rs32!Cheque_Dt) & "','" & Trim(rs33!Description) & "'," & Trim(rs33!Amount) & "," & 0 & ")"
               
                ws.BeginTrans
                db.Execute (Sqlqry38)
                ws.CommitTrans
               rs33.MoveNext
              Loop
            End If
          rs32.MoveNext
          Loop
         End If
         
        ' Journal Debit after From date and before to date
        Sqlqry39 = "select * from jrnl_tra where TDATE>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and TDATE<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Val(Mid(lstAcctCodes, 1, 6)) & "' and dc_code='D' "
        Set rs34 = db.OpenRecordset(Sqlqry39, dbOpenDynaset)
         If rs34.RecordCount <> 0 Then
          rs34.MoveFirst
          Do Until rs34.EOF
           Sqlqry40 = "Insert into ACCREPORT values('" & rs34!VOUC_NO & "','" & rs34!vouc_type & "','" & Trim(rs34!tDate) & "','" & Trim(rs34!Description) & "'," & Trim(rs34!damount) & "," & 0 & ")"
           ws.BeginTrans
           db.Execute (Sqlqry40)
           ws.CommitTrans
           rs34.MoveNext
          Loop
        End If
        
        ' Journal Credit after From date and before to date
        Sqlqry41 = "select * from Jrnl_tra where TDATE>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "' and dc_code='C'"
        Set rs35 = db.OpenRecordset(Sqlqry41, dbOpenDynaset)
        If rs35.RecordCount <> 0 Then
         rs35.MoveFirst
        Do Until rs35.EOF
          Sqlqry41 = "Insert into ACCREPORT values('" & rs35!VOUC_NO & "','" & rs35!vouc_type & "','" & Trim(rs35!tDate) & "','" & Trim(rs35!Description) & "'," & 0 & "," & Trim(rs35!camount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry41)
          ws.CommitTrans
          rs35.MoveNext
         Loop
        End If
        
       ' DebitNote - credit after From date and before to date
        Sqlqry42 = "select * from debt_mas where Tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs36 = db.OpenRecordset(Sqlqry42, dbOpenDynaset)
        If rs36.RecordCount <> 0 Then
         rs36.MoveFirst
        Do Until rs36.EOF
          Sqlqry43 = "Insert into ACCREPORT values('" & rs36!VOUC_NO & "','" & rs36!vouc_type & "','" & Trim(rs36!tDate) & "','" & Trim(rs36!Description) & "'," & 0 & "," & Trim(rs36!Amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry43)
          ws.CommitTrans
          rs36.MoveNext
         Loop
        End If
        
        ' DebitNote - debit after From date and before to date
        Sqlqry43 = "select * from debt_mas where Tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Cust_no='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs37 = db.OpenRecordset(Sqlqry43, dbOpenDynaset)
        If rs37.RecordCount <> 0 Then
         rs37.MoveFirst
        Do Until rs37.EOF
          Sqlqry44 = "Insert into ACCREPORT values('" & rs37!VOUC_NO & "','" & rs37!vouc_type & "','" & Trim(rs37!tDate) & "','" & Trim(rs37!Description) & "'," & Trim(rs37!Amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry44)
          ws.CommitTrans
          rs37.MoveNext
         Loop
        End If
        
        ' CreditNote - Debit after From date and before to date
        Sqlqry45 = "select * from crdt_mas where Tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Supp_no='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs38 = db.OpenRecordset(Sqlqry45, dbOpenDynaset)
        If rs38.RecordCount <> 0 Then
         rs38.MoveFirst
        Do Until rs38.EOF
          Sqlqry46 = "Insert into ACCREPORT values('" & rs38!VOUC_NO & "','" & rs38!vouc_type & "','" & Trim(rs38!tDate) & "','" & Trim(rs38!Description) & "'," & 0 & "," & Trim(rs38!Amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry46)
          ws.CommitTrans
          rs38.MoveNext
         Loop
        End If
        
        ' Credit Note - Credit after From date and before to date
        Sqlqry47 = "select * from crdt_mas where Tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs39 = db.OpenRecordset(Sqlqry47, dbOpenDynaset)
        If rs39.RecordCount <> 0 Then
         rs39.MoveFirst
        Do Until rs39.EOF
          Sqlqry48 = "Insert into ACCREPORT values('" & rs39!VOUC_NO & "','" & rs39!vouc_type & "','" & Trim(rs39!tDate) & "','" & Trim(rs39!Description) & "'," & Trim(rs39!Amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry48)
          ws.CommitTrans
          rs39.MoveNext
         Loop
        End If
        
         ' Credit Sale after From date and before to date
        Sqlqry49 = "select * from bo_mas where invoice_date>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and invoice_date<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs40 = db.OpenRecordset(Sqlqry49, dbOpenDynaset)
        If rs40.RecordCount <> 0 Then
         rs40.MoveFirst
        Do Until rs40.EOF
          Sqlqry50 = "Insert into ACCREPORT values('" & rs40!serial_no & "','INV','" & Trim(rs40!invoice_date) & "','Sale'," & 0 & "," & Val(rs40!net_Amount) * convertion & ")"
          ws.BeginTrans
          db.Execute (Sqlqry50)
          ws.CommitTrans
          rs40.MoveNext
         Loop
        End If
        
       ' Credit Purchase after From date and before to date
        Sqlqry51 = "select * from crpr_mas where Tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs41 = db.OpenRecordset(Sqlqry51, dbOpenDynaset)
        If rs41.RecordCount <> 0 Then
         rs41.MoveFirst
        Do Until rs41.EOF
          Sqlqry51 = "Insert into ACCREPORT values('" & rs41!VOUC_NO & "','" & rs41!vouc_type & "','" & Trim(rs41!tDate) & "','Purchase'," & Trim(rs41!namount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry51)
          ws.CommitTrans
          rs41.MoveNext
         Loop
        End If
       
         ' Cash Purchase after From date and before to date
        Sqlqry52 = "select * from capr_mas where Tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
        Set rs42 = db.OpenRecordset(Sqlqry52, dbOpenDynaset)
        If rs42.RecordCount <> 0 Then
         rs42.MoveFirst
        Do Until rs42.EOF
          Sqlqry52 = "Insert into ACCREPORT values('" & rs42!VOUC_NO & "','" & rs42!vouc_type & "','" & Trim(rs42!tDate) & "','Cash Purchase '," & Trim(rs42!namount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry52)
          ws.CommitTrans
          rs42.MoveNext
         Loop
        End If
       
         ' Cash  Sale after From date and before to date
        'Sqlqry53 = "select * from casl_mas where Tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstAcctCodes, 1, 6) & "'"
       ' Set rs43 = db.OpenRecordset(Sqlqry53, dbOpenDynaset)
       ' If rs43.RecordCount <> 0 Then
       '  rs43.MoveFirst
       ' Do Until rs43.EOF
       '   Sqlqry53 = "Insert into ACCREPORT values('" & rs43!VOUC_NO & "','" & rs43!vouc_type & "','" & Trim(rs43!tDate) & "','Cash Sale'," & 0 & "," & Trim(rs43!namount) & ")"
       '   ws.BeginTrans
       '   db.Execute (Sqlqry53)
       '   ws.CommitTrans
       '   rs43.MoveNext
       '  Loop
       ' End If
        
        
    With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\ACCREPORT.rpt"
     .Formulas(0) = "zzz='" & " From " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
     .Formulas(1) = "yyy='" & Mid(lstAcctCodes, 1, 40) & "'"
     .WindowState = crptMaximized
     .Action = 1
    End With
    
   Else
        MsgBox "Improper Dates", vbDefaultButton1, "Invalid entry"
        Exit Sub
  End If
 End Sub

Private Sub Form_Load()
PopulateAcctCodes
txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
txtdateto.TextWithMask = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub PopulateAcctCodes()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select * from Acct_mas where acct_code<'" & 103001 & "' order by acct_code"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

lstAcctCodes.Clear

 If rs.RecordCount = 0 Then
      MsgBox "No Records found in the Account Register"
 Else
      rs.MoveFirst
      Do Until rs.EOF
       lstAcctCodes.AddItem rs!acct_code & " : " & rs!acct_name
      rs.MoveNext
   Loop
 End If
 
Sqlqry1 = "Select * from Acct_mas where acct_code>'" & 103200 & "' order by acct_code"
Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)

 If rs1.RecordCount = 0 Then
      MsgBox "No Records found in the Account Register"
 Else
      rs1.MoveFirst
      Do Until rs1.EOF
       lstAcctCodes.AddItem rs1!acct_code & " : " & rs1!acct_name
       rs1.MoveNext
   Loop
 End If
End Sub
Private Sub lstAcctCodes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdatefrom.SetFocus
End Sub
Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdateto.SetFocus
End Sub

Private Sub txtdatefrom_LostFocus()
If IsDate(txtdatefrom.TextWithMask) = False Then
    MsgBox " Invalid From Date", vbInformation, "Invalid Entry"
    txtdatefrom.SetFocus
    SendKeys " {Home} + { End} "
End If
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdDisplay.SetFocus
End Sub
Private Function ValidateData()
 ValidateData = False

If IsDate(txtdatefrom.TextWithMask) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
 ElseIf lstAcctCodes.SelCount = 0 Then
   MsgBox "Select Account Code", vbInformation, "Invalid Entry"
   lstAcctCodes.SetFocus
   SendKeys " {Home} + {end} "
   Exit Function
 ElseIf IsDate(txtdateto.TextWithMask) = False Then
   MsgBox "Invalid To Date", vbInformation, "Invalid Entry"
   txtdateto.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
 Else
  ValidateData = True
 End If
End Function
Private Sub textclear()
 lstAcctCodes.ListIndex = -1
 txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
 txtdateto.TextWithMask = Format(Now, "dd/mm/YYYY")
End Sub

Private Sub txtdateto_LostFocus()
    If IsDate(txtdateto.TextWithMask) = False Then
       MsgBox "Invalid To Date", vbInformation, "Invalid Entry"
       txtdateto.SetFocus
       SendKeys " {Home} + {End} "
    End If
   
End Sub
