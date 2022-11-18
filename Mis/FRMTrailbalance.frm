VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmTBReport 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Statement of Trial Balance"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Trail Balance As On"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4455
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   5775
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Height          =   3495
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   4815
         Begin VB.CommandButton cmdBack 
            BackColor       =   &H00FFFF80&
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
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H00FFFF80&
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
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CommandButton cmdDisplay 
            BackColor       =   &H00FFFF80&
            Caption         =   "P&review"
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
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00FFFF80&
            Caption         =   "&Print"
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
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2280
            Width           =   1095
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   4080
            Top             =   600
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   262150
         End
         Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
            Height          =   375
            Left            =   1680
            TabIndex        =   0
            Top             =   840
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
            ForeColor       =   8388608
            Alignment       =   1
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   4800
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   960
            TabIndex        =   7
            Top             =   960
            Width           =   600
         End
      End
   End
   Begin VB.Label lblWait 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Please Wait -- -- -- TB In Process"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   6600
      Width           =   5295
   End
End
Attribute VB_Name = "frmTBReport"
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
Dim opcpmt As Currency
Dim opcrpt As Currency
Dim Opcasl As Currency
Dim Opbrpta As Currency
Dim Opbpmta As Currency
Dim Ttlopbal As Currency
Dim TTLCUST As Currency
Dim Advclbal As Currency
Dim VehClbal As Currency
Dim MacclBal As Currency
Dim ProClbal As Currency
Dim PurClbal As Currency
Dim SalClbal As Currency
Dim GoodClbal As Currency
Dim TMachinery As Currency
Dim TCostPrice As Currency
Dim X As Currency
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
Dim Sqlqry61 As String
Dim Sqlqry62 As String
Dim Sqlqry63 As String
Dim Sqlqry64 As String
Dim Sqlqry65 As String
Dim Sqlqry66 As String
Dim Sqlqry67 As String
Dim Sqlqry68 As String
Dim Sqlqry69 As String
Dim Sqlqry70 As String
Dim Sqlqry71 As String
Dim Sqlqry72 As String
Dim Sqlqry73 As String
Dim Sqlqry74 As String
Dim Sqlqry75 As String
Dim Sqlqry76 As String
Dim Sqlqry77 As String
Dim Sqlqry78 As String
Dim Sqlqry79 As String
Dim Sqlqry80 As String
Dim Sqlqry81 As String
Dim Sqlqry82 As String
Dim Sqlqry83 As String
Dim Sqlqry84 As String
Dim Sqlqry85 As String
Dim Sqlqry86 As String
Dim Sqlqry87 As String
Dim Sqlqry88 As String
Dim Sqlqry89 As String
Dim Sqlqry90 As String
Dim Sqlqry91 As String
Dim Sqlqry92 As String
Dim Sqlqry93 As String
Dim Sqlqry94 As String
Dim Sqlqry95 As String
Dim Sqlqry96 As String
Dim Sqlqry97 As String
Dim Sqlqry98 As String
Dim Sqlqry99 As String
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
Dim rs51 As Recordset
Dim rs52 As Recordset
Dim rs53 As Recordset
Dim rs54 As Recordset
Dim rs55 As Recordset
Dim rs56 As Recordset
Dim rs57 As Recordset
Dim rs58 As Recordset
Dim rs59 As Recordset
Dim rs60 As Recordset
Dim rs61 As Recordset
Dim rs62 As Recordset
Dim rs63 As Recordset
Dim rs64 As Recordset
Dim rs65 As Recordset
Dim rs66 As Recordset
Dim rs67 As Recordset
Dim rs68 As Recordset
Dim rs69 As Recordset
Dim rs70 As Recordset
Dim rs71 As Recordset
Dim rs72 As Recordset
Dim rs73 As Recordset
Dim rs74 As Recordset
Dim rs75 As Recordset
Dim rs76 As Recordset
Dim rs77 As Recordset
Dim rs78 As Recordset
Dim rs79 As Recordset
Dim rs80 As Recordset
Dim rs81 As Recordset
Dim rs82 As Recordset
Dim opsal As Currency
Dim oprec As Currency
Dim oppur As Currency
Dim oppay As Currency
Dim Opbrpt As Currency
Dim Opprpt As Currency
Dim Opjdb As Currency
Dim Opbpmt As Currency
Dim Opppmt As Currency
Dim Opjcr As Currency
Dim opjbd As Currency
Dim Opjbc As Currency
Dim Opcrsl As Currency
Dim Opcrpr As Currency
Dim Ttlsupp As Currency
Dim Tpdc As Currency

Private Sub cmdBack_Click()
 Unload Me
End Sub

Private Sub CmdPrint_Click()
 With CrystalReport1
       .DataFiles(0) = App.Path & "\misov.mdb"
       .ReportFileName = App.Path & "\TrialBalance1.rpt"
       .SelectionFormula = "{acct_mas.Close_bal}<>" & 0 & ""
       .Formulas(0) = "xxx1='" & txtdatefrom.TextWithMask & "'"
       .WindowMaxButton = True
       .WindowState = crptMaximized
       .Action = 1
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
 End Sub

Private Sub cmdClear_Click()
 textclear
End Sub
 
 Private Sub cmdDisplay_Click()
 lblWait.Visible = True
 lblWait.Caption = "Please Wait -- -- -- TB In Process "
 
    
    If ValidateData = True Then
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    doclosnil
    Sqlqry = " select * from acct_mas order by acct_code"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount = 0 Then
       MsgBox " No Records Found In The Account Register"
       Exit Sub
     Else
      rs.MoveFirst
      Do Until rs.EOF
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
        TTLCUST = 0
        Ttlsupp = 0
        Opcasl = 0
        Tpdc = 0
        opsal = 0
        oprec = 0
        oppur = 0
        oppay = 0
        Opbrpt = 0
        Opbpmt = 0
        opcpmt = 0
        opcrpt = 0
        Opprpt = 0
        Opppmt = 0
        Opjdb = 0
        Opjcr = 0
        Opcrsl = 0
        Opcrpr = 0
        If IsNull(rs!open_bal) = True Then
         Opbal = 0
        Else
         Opbal = Val(rs!open_bal)
        End If
        
     If Val(rs!acct_code) = 103001 Then
         docashbal
     ElseIf Val(rs!acct_code) = 103002 Then
        docashbalusd
     ElseIf Val(rs!acct_code) >= 103101 And Val(rs!acct_code) < 103200 Then
         dobankbal
     Else
         
     '  Cash Receipt(credit) before From date
        Sqlqry1 = " select sum(amount) from crpt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Opcrptcr = rs1.Fields(0)
        
     ' Cash Receipt(debit) before From date
        Sqlqry2 = " select sum(ttl_amount)from crpt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Cash_code='" & rs!acct_code & "'"
        Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs2.Fields(0)) = False Then Opcrptdb = rs2.Fields(0)
        
     '  Cash Payment(debit) before from date
        Sqlqry3 = " select sum(amount) from cpmt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs3 = db.OpenRecordset(Sqlqry3, dbOpenDynaset)
        If IsNull(rs3.Fields(0)) = False Then Opcpmtdb = rs3.Fields(0)
                
     '  Cash Payment(credit) before from date
        Sqlqry4 = " select sum(ttl_Amount) from cpmt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and cash_code='" & rs!acct_code & "'"
        Set rs4 = db.OpenRecordset(Sqlqry4, dbOpenDynaset)
        If IsNull(rs4.Fields(0)) = False Then Opcpmtcr = rs4.Fields(0)
              
     '  Bank Receipt(credit) before from date.
        Sqlqry5 = "select sum(amount) from brpt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "'"
        Set rs5 = db.OpenRecordset(Sqlqry5, dbOpenDynaset)
        If IsNull(rs5.Fields(0)) = False Then Opbrptcr = rs5.Fields(0)
        
     '  Bank Receipt(debit) before from date.
        Sqlqry6 = "select sum(ttl_amount) from brpt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & rs!acct_code & "'"
        Set rs6 = db.OpenRecordset(Sqlqry6, dbOpenDynaset)
        If IsNull(rs6.Fields(0)) = False Then Opbrptdb = rs6.Fields(0)
                       
     '  Bank Payment(debit) before From date
        Sqlqry7 = "select sum(amount) from bpmt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs7 = db.OpenRecordset(Sqlqry7, dbOpenDynaset)
        If IsNull(rs7.Fields(0)) = False Then Opbpmtdb = rs7.Fields(0)
                        
     '  Bank Payment(credit) before From date
        Sqlqry8 = "select sum(ttl_amount) from bpmt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & rs!acct_code & "'"
        Set rs8 = db.OpenRecordset(Sqlqry8, dbOpenDynaset)
        If IsNull(rs8.Fields(0)) = False Then Opbpmtcr = rs8.Fields(0)
                
      ' Pdc Receipts (debit) before From date
        Sqlqry9 = "select sum(amount) from prpt_mas1 where Cheque_dt<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_dt) and bank_code = '" & rs!acct_code & "' "
        Set rs9 = db.OpenRecordset(Sqlqry9, dbOpenDynaset)
        If IsNull(rs9.Fields(0)) = False Then Opprptdb = rs9.Fields(0)
        
      ' Pdc Receipts (credit) before From date
        Sqlqry10 = "Select sum(amount) from Prpt_mas1 where  Cheque_dt<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_dt) and acct_code='" & rs!acct_code & "'"
        Set rs10 = db.OpenRecordset(Sqlqry10, dbOpenDynaset)
        If IsNull(rs10.Fields(0)) = False Then Opprptcr = rs10.Fields(0)
        
       ' Pdc Payments before From Date
        SQLQRY11 = "select sum(ttl_amount) from Ppmt_mas where Cheque_dt<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_Dt) and Bank_Code='" & rs!acct_code & "'"
        Set rs11 = db.OpenRecordset(SQLQRY11, dbOpenDynaset)
        If IsNull(rs11.Fields(0)) = False Then Opppmtcr = rs11.Fields(0)
                   
       ' Opening Balance Pdc Payments before From Date
        SQLQRY12 = "Select sum(amount) from Prpt_tra where Cheque_dt<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_Dt) and acct_code='" & rs!acct_code & "'"
        Set rs12 = db.OpenRecordset(SQLQRY12, dbOpenDynaset)
        If IsNull(rs12.Fields(0)) = False Then Opppmtdb = rs12.Fields(0)
        
       ' Journal Debit Amount before From Date
        Sqlqry13 = "select sum(damount) from Jrnl_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "' and dc_code ='D'"
        Set rs13 = db.OpenRecordset(Sqlqry13, dbOpenDynaset)
        If IsNull(rs13.Fields(0)) = False Then Opjrnldb = rs13.Fields(0)
         
       ' Journal Credit Amount before From Date
        Sqlqry14 = "select sum(camount) from Jrnl_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "' and dc_code ='C'"
        Set rs14 = db.OpenRecordset(Sqlqry14, dbOpenDynaset)
        If IsNull(rs14.Fields(0)) = False Then Opjrnlcr = rs14.Fields(0)
        
       ' Debit note (credit) before From Date
        Sqlqry15 = "select sum(amount) from debt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "'"
        Set rs15 = db.OpenRecordset(Sqlqry15, dbOpenDynaset)
        If IsNull(rs15.Fields(0)) = False Then opdbntcr = rs15.Fields(0)
        
       ' Debit note (debit)  before From Date
        sqlqry16 = "select sum(amount) from debt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and cust_no='" & rs!acct_code & "'"
        Set rs16 = db.OpenRecordset(sqlqry16, dbOpenDynaset)
        If IsNull(rs16.Fields(0)) = False Then opdbntdb = rs16.Fields(0)
        
       ' Credit note (debit) before From Date
        sqlqry17 = "select sum(amount) from crdt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "'"
        Set rs17 = db.OpenRecordset(sqlqry17, dbOpenDynaset)
        If IsNull(rs17.Fields(0)) = False Then Opcrntdb = rs17.Fields(0)
         
       ' Credit note (credit)  before From Date
         sqlqry18 = "select sum(amount) from crdt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and supp_no='" & rs!acct_code & "'"
        Set rs18 = db.OpenRecordset(sqlqry18, dbOpenDynaset)
        If IsNull(rs18.Fields(0)) = False Then Opcrntcr = rs18.Fields(0)
                  
       ' Credit Sales  before From Date
        sqlqry19 = "select sum(net_amount) from bo_mas where Invoice_date<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs19 = db.OpenRecordset(sqlqry19, dbOpenDynaset)
        If IsNull(rs19.Fields(0)) = False Then Opcrslcr = rs19.Fields(0)
                
       ' Credit Purchases  before From Date
        sqlqry20 = "select sum(gamount) from crpr_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs20 = db.OpenRecordset(sqlqry20, dbOpenDynaset)
        If IsNull(rs20.Fields(0)) = False Then Opcrprdb = rs20.Fields(0)
                        
       ' Cash Sales(credit)  before From Date
        'Sqlqry21 = "select sum(namount) from casl_mas where tdate<=#" & DateValue(Format(txtdatefrom.textwithmask, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        'Set rs21 = db.OpenRecordset(Sqlqry21, dbOpenDynaset)
        'If IsNull(rs21.Fields(0)) = False Then OpcaslCr = rs21.Fields(0)
        
       ' Cash Sales(debit)  before From Date
        'Sqlqry22 = "select sum(namount) from casl_mas where tdate<=#" & DateValue(Format(txtdatefrom.textwithmask, "dd/mm/yyyy")) & "# and cash_code='" & rs!acct_code & "'"
        'Set rs22 = db.OpenRecordset(Sqlqry22, dbOpenDynaset)
        'If IsNull(rs22.Fields(0)) = False Then Opcasldb = rs22.Fields(0)
                
       ' Cash Purchases(debit)  before From Date
        Sqlqry23 = "select sum(gamount) from capr_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs23 = db.OpenRecordset(Sqlqry23, dbOpenDynaset)
        If IsNull(rs23.Fields(0)) = False Then Opcaprdb = rs23.Fields(0)
                                 
       ' Cash Purchases(Credit)  before From Date
        Sqlqry24 = "select sum(gamount) from capr_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and cash_code='" & rs!acct_code & "'"
        Set rs24 = db.OpenRecordset(Sqlqry24, dbOpenDynaset)
        If IsNull(rs24.Fields(0)) = False Then Opcaprcr = rs24.Fields(0)
                          
        Ttlopbal = Opbal - Opcrptcr + Opcrptdb - Opcpmtcr + Opcpmtdb - Opbrptcr + Opbrptdb + Opbpmtdb - Opbpmtcr + Opprptdb - Opprptcr + Opppmtdb _
               - Opppmtcr + Opjrnldb - Opjrnlcr + opdbntdb - opdbntcr + Opcrntdb - Opcrntcr - Opcrslcr + Opcrprdb - OpcaslCr + Opcasldb - Opcaprcr + Opcaprdb
                
        
       End If
                
          Sqlqry25 = "Update acct_mas set close_bal =" & Ttlopbal & " where Acct_code ='" & rs!acct_code & "'"
          ws.BeginTrans
          db.Execute (Sqlqry25)
          ws.CommitTrans
       
       rs.MoveNext
       Loop
       
       DOCUSTPOST
       DOSUPPPOST
       SalClbal = 0
       PurClbal = 0
       ProClbal = 0
       VehClbal = 0
       MacclBal = 0
       Advclbal = 0
       TMachinery = 0
       TCostPrice = 0
        Sqlqry80 = " Select SUM(CLOSE_BAL) from  acct_mas where acct_code >='105000' and acct_code<='105499' "
        Set rs73 = db.OpenRecordset(Sqlqry80, dbOpenDynaset)
        If IsNull(rs73.Fields(0)) = False Then Advclbal = rs73.Fields(0)
        
            Sqlqry81 = " Update acct_mas set close_bal=" & Advclbal & " where acct_code ='105000'"
                ws.BeginTrans
                db.Execute (Sqlqry81)
                ws.CommitTrans
             Sqlqry82 = " update acct_mas set close_bal =" & 0 & " where acct_code>='105001' and acct_code<='105499'"
                ws.BeginTrans
                db.Execute (Sqlqry82)
                ws.CommitTrans
            
        Sqlqry83 = " Select SUM(CLOSE_BAL) from  acct_mas where acct_code >='417000' and acct_code<'417100' "
        Set rs74 = db.OpenRecordset(Sqlqry83, dbOpenDynaset)
        If IsNull(rs74.Fields(0)) = False Then VehClbal = rs74.Fields(0)
        
        
            Sqlqry84 = " Update acct_mas set close_bal=" & VehClbal & " where acct_code ='417000'"
            ws.BeginTrans
            db.Execute (Sqlqry84)
            ws.CommitTrans
            
            Sqlqry85 = " update acct_mas set close_bal =" & 0 & " where acct_code>='417001' and acct_code<'417100'"
            ws.BeginTrans
            db.Execute (Sqlqry85)
            ws.CommitTrans
         
        Sqlqry86 = " Select SUM(CLOSE_BAL) from  acct_mas where acct_code >='204000' and acct_code<'204100' "
        Set rs75 = db.OpenRecordset(Sqlqry86, dbOpenDynaset)
        If IsNull(rs75.Fields(0)) = False Then ProClbal = rs75.Fields(0)
        
              
            Sqlqry87 = " Update acct_mas set close_bal=" & ProClbal & " where acct_code ='204000'"
            ws.BeginTrans
            db.Execute (Sqlqry87)
            ws.CommitTrans
            
            Sqlqry88 = " Update acct_mas set close_bal =" & 0 & " where acct_code>='204001' and acct_code<'204100'"
            ws.BeginTrans
            db.Execute (Sqlqry88)
            ws.CommitTrans
            
          
           Sqlqry89 = " Select SUM(CLOSE_BAL) from  acct_mas where acct_code >='401000' and acct_code<'404000' "
            Set rs76 = db.OpenRecordset(Sqlqry89, dbOpenDynaset)
            If IsNull(rs76.Fields(0)) = False Then PurClbal = rs76.Fields(0)
            Sqlqry90 = " Update acct_mas set close_bal=" & PurClbal & " where acct_code ='401000'"
            ws.BeginTrans
            db.Execute (Sqlqry90)
            ws.CommitTrans
            Sqlqry91 = " update acct_mas set close_bal =" & 0 & " where acct_code>='401001' and acct_code<'404000'"
            ws.BeginTrans
            db.Execute (Sqlqry91)
            ws.CommitTrans
         
       ' Sqlqry92 = " Select SUM(CLOSE_BAL) from  acct_mas where acct_code >='301000' and acct_code<'304000' "
       ' Set rs77 = db.OpenRecordset(Sqlqry92, dbOpenDynaset)
       ' If IsNull(rs77.Fields(0)) = False Then SalClbal = rs77.Fields(0)
        
       '     Sqlqry93 = " Update acct_mas set close_bal=" & SalClbal & " where acct_code ='301000'"
       '     ws.BeginTrans
       '     db.Execute (Sqlqry93)
       '     ws.CommitTrans
       '     Sqlqry94 = " Update acct_mas set close_bal =" & 0 & " where acct_code>='301001' and acct_code<'304000'"
       '     ws.BeginTrans
       '     db.Execute (Sqlqry94)
       '     ws.CommitTrans
    End If
          
     With CrystalReport1
       .DataFiles(0) = App.Path & "\misov.mdb"
       .ReportFileName = App.Path & "\TrialBalance1.rpt"
       .Formulas(0) = "xxx1='" & txtdatefrom.TextWithMask & "'"
       .SelectionFormula = "{acct_mas.Close_bal}<> " & 0 & ""
       .WindowMaxButton = True
       .WindowState = crptMaximized
       .Action = 1
     End With
     
    lblWait.Visible = False
   Else
        lblWait.Visible = False
        MsgBox "Improper Dates", vbDefaultButton1, "Invalid entry"
        Exit Sub
  End If
  
End Sub

Private Sub Form_Load()
    txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
    lblWait.Visible = False
End Sub


Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdDisplay.SetFocus
End Sub

Private Function ValidateData()
 ValidateData = False

If IsDate(txtdatefrom.TextWithMask) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
   ValidateData = True
End If

End Function

Private Sub textclear()
  txtdatefrom.TextWithMask = Format(Now(), "dd/mm/yyyy")
End Sub

Private Sub docashbal()
   
       Opbpmt = 0
       Opbrpt = 0
       opsal = 0
       oprec = 0
       oppur = 0
       oppay = 0
       Opjdb = 0
       Opjcr = 0
       Ttlopbal = 0
       ' Total Amount of Sales before From date
        'Sqlqry26 = "select sum(namount) from casl_mas where Cash_code='103001' and tdate<=#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "#"
        'Set rs25 = db.OpenRecordset(Sqlqry26, dbOpenDynaset)
        'If IsNull(rs25.Fields(0)) = False Then opsal = rs25.Fields(0)
              
       ' Total Amount of Purchase before From date
        Sqlqry27 = "select sum(gamount) from capr_mas where Cash_code='103001' and tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs26 = db.OpenRecordset(Sqlqry27, dbOpenDynaset)
        If IsNull(rs26.Fields(0)) = False Then oppur = rs26.Fields(0)
                       
       ' Total cash Receipts before From date
        Sqlqry28 = "select sum(ttl_Amount) from crpt_mas where Cash_code='103001' and tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs27 = db.OpenRecordset(Sqlqry28, dbOpenDynaset)
        If IsNull(rs27.Fields(0)) = False Then oprec = rs27.Fields(0)
                
       ' Total cash Payments before To date
        Sqlqry29 = "select sum(ttl_amount) from cpmt_mas where Cash_code ='103001' and tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs28 = db.OpenRecordset(Sqlqry29, dbOpenDynaset)
        If IsNull(rs28.Fields(0)) = False Then oppay = rs28.Fields(0)
        
                
        ' Bank Payment  debit Before before From Date
        sqlqry16 = "select sum(amount) from bpmt_tra where Acct_Code ='103001' and  tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs11 = db.OpenRecordset(sqlqry16, dbOpenDynaset)
        If IsNull(rs11.Fields(0)) = False Then Opbpmt = rs11.Fields(0)
        
                 
        ' Bank Receipt Credit Before before From Date
        sqlqry17 = "select Sum(amount) from brpt_tra where Acct_Code ='103001' and  tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs12 = db.OpenRecordset(sqlqry17, dbOpenDynaset)
        If IsNull(rs12.Fields(0)) = False Then Opbrpt = rs12.Fields(0)
         
         ' Journal Debit Amount before From Date
        Sqlqry78 = "select SUM(DAMOUNT) from Jrnl_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='103001' and dc_code ='D'"
        Set rs71 = db.OpenRecordset(Sqlqry78, dbOpenDynaset)
        If IsNull(rs71.Fields(0)) = False Then Opjdb = rs71.Fields(0)
        
                 
        ' Journal Credit Amount before From Date
        Sqlqry79 = "select SUM(CAMOUNT) from Jrnl_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='103001' and dc_code ='C'"
        Set rs72 = db.OpenRecordset(Sqlqry79, dbOpenDynaset)
        If IsNull(rs72.Fields(0)) = False Then Opjcr = rs72.Fields(0)
                        
         Ttlopbal = Opbal + Opbpmt - Opbrpt + opsal + oprec - oppur - oppay + Opjdb - Opjcr
      'MsgBox
      
End Sub
Private Sub docashbalusd()
   
       Opbpmt = 0
       Opbrpt = 0
       opsal = 0
       oprec = 0
       oppur = 0
       oppay = 0
       Opjdb = 0
       Opjcr = 0
       Ttlopbal = 0
       
       ' Total Amount of Sales before From date
        'Sqlqry26 = "select sum(namount) from casl_mas where Cash_code='103002' and tdate<=#" & DateValue(Format(txtdatefrom.textwithmask, "dd/mm/yyyy")) & "#"
        'Set rs25 = db.OpenRecordset(Sqlqry26, dbOpenDynaset)
        'If IsNull(rs25.Fields(0)) = False Then opsal = rs25.Fields(0)
              
       ' Total Amount of Purchase before From date
        Sqlqry27 = "select sum(gamount) from capr_mas where Cash_code='103002' and tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs26 = db.OpenRecordset(Sqlqry27, dbOpenDynaset)
        If IsNull(rs26.Fields(0)) = False Then oppur = rs26.Fields(0)
                       
       ' Total cash Receipts before From date
        Sqlqry28 = "select sum(ttl_Amount) from crpt_mas where Cash_code='103002' and tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs27 = db.OpenRecordset(Sqlqry28, dbOpenDynaset)
        If IsNull(rs27.Fields(0)) = False Then oprec = rs27.Fields(0)
                
       ' Total cash Payments before To date
        Sqlqry29 = "select sum(ttl_amount) from cpmt_mas where Cash_code ='103002' and tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs28 = db.OpenRecordset(Sqlqry29, dbOpenDynaset)
        If IsNull(rs28.Fields(0)) = False Then oppay = rs28.Fields(0)
        
                
        ' Bank Payment  debit Before before From Date
        sqlqry16 = "select sum(amount) from bpmt_tra where Acct_Code ='103002' and  tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs11 = db.OpenRecordset(sqlqry16, dbOpenDynaset)
        If IsNull(rs11.Fields(0)) = False Then Opbpmt = rs11.Fields(0)
        
                 
        ' Bank Receipt Credit Before before From Date
        sqlqry17 = "select Sum(amount) from brpt_tra where Acct_Code ='103002' and  tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs12 = db.OpenRecordset(sqlqry17, dbOpenDynaset)
        If IsNull(rs12.Fields(0)) = False Then Opbrpt = rs12.Fields(0)
         
         ' Journal Debit Amount before From Date
        Sqlqry78 = "select SUM(DAMOUNT) from Jrnl_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='103002' and dc_code ='D'"
        Set rs71 = db.OpenRecordset(Sqlqry78, dbOpenDynaset)
        If IsNull(rs71.Fields(0)) = False Then Opjdb = rs71.Fields(0)
        
                 
        ' Journal Credit Amount before From Date
        Sqlqry79 = "select SUM(CAMOUNT) from Jrnl_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='103002' and dc_code ='C'"
        Set rs72 = db.OpenRecordset(Sqlqry79, dbOpenDynaset)
        If IsNull(rs72.Fields(0)) = False Then Opjcr = rs72.Fields(0)
                        
         Ttlopbal = Opbal + Opbpmt - Opbrpt + opsal + oprec - oppur - oppay + Opjdb - Opjcr
      
End Sub


Public Sub dobankbal()
               
        Ttlopbal = 0
        Opbal = 0
        Opcasl = 0
        opcpmt = 0
        opcrpt = 0
        Opbrpt = 0
        Opprpt = 0
        Opbrpta = 0
        Opjdb = 0
        Opbpmt = 0
        Opbpmta = 0
        Opppmt = 0
        Opjcr = 0
        opdbntcr = 0
        Opcrntdb = 0
        
        'Sqlqry28 = "select SUM(NAMOUNT) from casl_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and cash_code='" & rs!acct_code & "'"
        'Set rs13 = db.OpenRecordset(Sqlqry28, dbOpenDynaset)
        'If IsNull(rs13.Fields(0)) = False Then Opcasl = rs13.Fields(0)
        ' Opening Balance
        Sqlqry28 = " select * from bank_mas where bank_code='" & rs!acct_code & "'"
        Set rs13 = db.OpenRecordset(Sqlqry28, dbOpenDynaset)
        If rs13.RecordCount <> 0 Then
           Opbal = Round(Val(rs13!Open_baldhs) + Val(rs13!open_balUSD) * convertion, 2)
        End If
               
      ' Bank Receipt bank code  before From date
        Sqlqry30 = "select SUM(TTL_AMOUNT) from brpt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & rs!acct_code & "'"
        Set rs18 = db.OpenRecordset(Sqlqry30, dbOpenDynaset)
        If IsNull(rs18.Fields(0)) = False Then Opbrpt = rs18.Fields(0)
        
        
      ' Bank Receipt Acct Code before From date
        Sqlqry30 = "select SUM(AMOUNT) from brpt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "'"
        Set rs29 = db.OpenRecordset(Sqlqry30, dbOpenDynaset)
        If IsNull(rs29.Fields(0)) = False Then Opbrpta = rs29.Fields(0)
        
               
      ' Debit note Credit before From date
        Sqlqry5 = "select SUM(AMOUNT) from debt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs5 = db.OpenRecordset(Sqlqry5, dbOpenDynaset)
        If IsNull(rs5.Fields(0)) = False Then opdbntcr = rs5.Fields(0)
     
        
      ' Credit note Debit before From date
        Sqlqry5 = "select SUM(AMOUNT) from Crdt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs5 = db.OpenRecordset(Sqlqry5, dbOpenDynaset)
        If IsNull(rs5.Fields(0)) = False Then Opcrntdb = rs5.Fields(0)
                 
        
      ' Bank Payment Bank Code before From date
        Sqlqry31 = "select SUM(TTL_AMOUNT) from bpmt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & rs!acct_code & "'"
        Set rs30 = db.OpenRecordset(Sqlqry31, dbOpenDynaset)
        If IsNull(rs30.Fields(0)) = False Then Opbpmt = rs30.Fields(0)
        
                
      ' Bank Payment Account Code before From date
        Sqlqry31 = "select SUM(AMOUNT) from bpmt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs30 = db.OpenRecordset(Sqlqry31, dbOpenDynaset)
        If IsNull(rs30.Fields(0)) = False Then Opbpmta = rs30.Fields(0)
        
                            
      ' Cash receipt (credit) before From date
        sqlqry20 = "select SUM(AMOUNT) from crpt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs13 = db.OpenRecordset(sqlqry20, dbOpenDynaset)
        If IsNull(rs13.Fields(0)) = False Then opcrpt = rs13.Fields(0)
       
        
      ' Cash Payment (Debit) before From date
        Sqlqry21 = "select SUM(AMOUNT) from cpmt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs14 = db.OpenRecordset(Sqlqry21, dbOpenDynaset)
        If IsNull(rs14.Fields(0)) = False Then opcpmt = rs14.Fields(0)
        
                
       ' Pdc Receipts before From date
        Sqlqry32 = "select SUM(AMOUNT) from prpt_mas1 where Cheque_dt<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & rs!acct_code & "' and not isnull(posting_dt)"
        Set rs31 = db.OpenRecordset(Sqlqry32, dbOpenDynaset)
        If IsNull(rs31.Fields(0)) = False Then Opprpt = rs31.Fields(0)
                        
        ' Pdc Payments before From Date
        Sqlqry33 = "select SUM(TTL_AMOUNT) from Ppmt_mas where Cheque_dt<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & rs!acct_code & "' and not isnull(posting_Dt)"
        Set rs32 = db.OpenRecordset(Sqlqry33, dbOpenDynaset)
        If IsNull(rs32.Fields(0)) = False Then Opppmt = rs32.Fields(0)
                 
       ' Journal Debit Amount before From Date
        Sqlqry34 = "select SUM(DAMOUNT) from Jrnl_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "' and dc_code ='D'"
        Set rs33 = db.OpenRecordset(Sqlqry34, dbOpenDynaset)
        If IsNull(rs33.Fields(0)) = False Then Opjdb = rs33.Fields(0)
                       
       ' Journal Credit Amount before From Date
        Sqlqry35 = "select SUM(CAMOUNT) from Jrnl_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "' and dc_code ='C'"
        Set rs34 = db.OpenRecordset(Sqlqry35, dbOpenDynaset)
        If IsNull(rs34.Fields(0)) = False Then Opjcr = rs34.Fields(0)
        
                 
        Ttlopbal = Opbal + Opcasl + opcpmt + Opcrntdb - opdbntcr - opcrpt + Opbrpt - Opbrpta + Opprpt + Opjdb - Opbpmt + Opbpmta - Opppmt - Opjcr
End Sub
Private Sub DOCUSTPOST()
    
   TTLCUST = 0
   Sqlqry36 = " select * from agndtls order by agentname"
   Set rs35 = db.OpenRecordset(Sqlqry36, dbOpenDynaset)
   If rs35.RecordCount = 0 Then
     MsgBox " Agency Code not found in Agency Register"
     Exit Sub
    Else
     rs35.MoveFirst
      Do Until rs35.EOF
         
                Opbal = 0
                opcrpt = 0
                opcpmt = 0
                Opbrpt = 0
                Opbpmt = 0
                Opprpt = 0
                Opppmt = 0
                opjbd = 0
                Opjbc = 0
                opdbntdb = 0
                opdbntcr = 0
                Opcrntdb = 0
                Opcrntcr = 0
                Opcrsl = 0
                Opcrpr = 0
                Ttlopbal = 0
                Tpdc = 0
                
        If IsNull(rs35!op_USD) = True Then
          Opbal = 0
        Else
           Opbal = rs35!op_USD * convertion
        End If
        
        If IsNull(rs35!op_DHS) = True Then
          Opbal = Opbal
        Else
         Opbal = Opbal + rs35!op_DHS
        End If
        
       ' Cash Receipt before From date
         Sqlqry37 = "select SUM(AMOUNT) from crpt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs35!agentname) & "'"
         Set rs36 = db.OpenRecordset(Sqlqry37, dbOpenDynaset)
         If IsNull(rs36.Fields(0)) = False Then opcrpt = rs36.Fields(0)
        
       ' Cash Payment before From date
         Sqlqry38 = "select SUM(AMOUNT) from cpmt_Tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs35!agentname) & "'"
         Set rs37 = db.OpenRecordset(Sqlqry38, dbOpenDynaset)
         If IsNull(rs37.Fields(0)) = False Then opcpmt = rs37.Fields(0)
                        
       ' Bank Receipt before From date
         Sqlqry39 = "select SUM(AMOUNT) from brpt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs35!agentname) & "'"
         Set rs38 = db.OpenRecordset(Sqlqry39, dbOpenDynaset)
         If IsNull(rs38.Fields(0)) = False Then Opbrpt = rs38.Fields(0)
        
       ' Bank Payment before From date
         Sqlqry40 = "select SUM(AMOUNT) from bpmt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs35!agentname) & "'"
         Set rs39 = db.OpenRecordset(Sqlqry40, dbOpenDynaset)
         If IsNull(rs39.Fields(0)) = False Then Opbpmt = rs39.Fields(0)
        
          
       ' Pdc Receipts before From date
         Sqlqry42 = "Select sum(amount) from Prpt_mas1 where Cheque_dt<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_dt) and acct_name='" & findfirstfixup(rs35!agentname) & "'"
         Set rs41 = db.OpenRecordset(Sqlqry42, dbOpenDynaset)
         If IsNull(rs41.Fields(0)) = False Then Opprpt = rs41.Fields(0)
         
        
        ' Pdc Payments before From Date
          Sqlqry44 = "Select sum(amount) from Ppmt_tra where Cheque_dt<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_Dt) and acct_name='" & findfirstfixup(rs35!agentname) & "'"
          Set rs43 = db.OpenRecordset(Sqlqry44, dbOpenDynaset)
          If IsNull(rs43.Fields(0)) = False Then Opppmt = rs43.Fields(0)
         
       ' Journal Debit Amount before From Date
          Sqlqry45 = "select sum(damount) from Jrnl_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs35!agentname) & "' and dc_code ='D'"
          Set rs44 = db.OpenRecordset(Sqlqry45, dbOpenDynaset)
          If IsNull(rs44.Fields(0)) = False Then opjbd = rs44.Fields(0)
              
         
       ' Journal Credit Amount before From Date
         Sqlqry46 = "select sum(camount) from Jrnl_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs35!agentname) & "' and dc_code ='C'"
         Set rs45 = db.OpenRecordset(Sqlqry46, dbOpenDynaset)
         If IsNull(rs45.Fields(0)) = False Then Opjbc = rs45.Fields(0)
        
                 
        ' Opening balance  debit note (credit) before From Date
        Sqlqry47 = "select sum(amount) from debt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs35!agentname) & "'"
        Set rs46 = db.OpenRecordset(Sqlqry47, dbOpenDynaset)
        If IsNull(rs46.Fields(0)) = False Then opdbntcr = rs46.Fields(0)
        
         
       ' Opening balance debit note (debit)  before From Date
        Sqlqry48 = "select sum(amount) from debt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and cust_name='" & findfirstfixup(rs35!agentname) & "'"
        Set rs47 = db.OpenRecordset(Sqlqry48, dbOpenDynaset)
        If IsNull(rs47.Fields(0)) = False Then opdbntdb = rs47.Fields(0)
        
              
       ' Opening balance  Credit note (debit) before From Date
        Sqlqry49 = "select sum(amount) from crdt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs35!agentname) & "'"
        Set rs48 = db.OpenRecordset(Sqlqry49, dbOpenDynaset)
        If IsNull(rs48.Fields(0)) = False Then Opcrntdb = rs48.Fields(0)
        
       ' Opening balance credit note (credit)  before From Date
        Sqlqry50 = "select sum(amount) from crdt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and supp_name='" & findfirstfixup(rs35!agentname) & "'"
        Set rs49 = db.OpenRecordset(Sqlqry50, dbOpenDynaset)
        If IsNull(rs49.Fields(0)) = False Then Opcrntcr = rs49.Fields(0)
        
       ' Opening balance credit Sales  before From Date
        Sqlqry51 = "select sum(net_amount) from bo_mas where invoice_date<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and agency='" & findfirstfixup(rs35!agentname) & "'"
        Set rs50 = db.OpenRecordset(Sqlqry51, dbOpenDynaset)
        If IsNull(rs50.Fields(0)) = False Then Opcrsl = rs50.Fields(0)
        
       ' Opening balance credit Purchases  before From Date
        Sqlqry52 = "select sum(gamount) from crpr_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and supp_name='" & findfirstfixup(rs35!agentname) & "'"
        Set rs51 = db.OpenRecordset(Sqlqry52, dbOpenDynaset)
        If IsNull(rs51.Fields(0)) = False Then Opcrpr = rs51.Fields(0)
                 
        Ttlopbal = Opbal - opcrpt + opcpmt - Opbrpt + Opbpmt - Opprpt + Opppmt + opjbd - Opjbc + opdbntdb - opdbntcr _
                    + Opcrntdb - Opcrntcr + Opcrsl - Opcrpr
                    
        TTLCUST = TTLCUST + Ttlopbal
        
     rs35.MoveNext
     Loop
         
    ' Pending Post Dated Cheques Received
     Sqlqry53 = "select sum(Amount) from prpt_mas1 where isnull(posting_dt) "
     Set rs52 = db.OpenRecordset(Sqlqry53, dbOpenDynaset)
     If IsNull(rs52.Fields(0)) = False Then Tpdc = rs52.Fields(0)
     TTLCUST = TTLCUST - Tpdc
          
     Sqlqry54 = "Update acct_mas set close_bal=" & TTLCUST & " where acct_code ='102000'"
     ws.BeginTrans
     db.Execute (Sqlqry54)
     ws.CommitTrans
       
       ' 103501 = Bills Receivable
  
        Sqlqry54 = "Select * FROM ACCT_MAS WHERE acct_code='103501'"
        Set rs54 = db.OpenRecordset(Sqlqry54, dbOpenDynaset)
         If rs54.RecordCount <> 0 Then
          rs54.MoveFirst
          If IsNull(rs54!open_bal) = True Then
            X = 0
          Else
            X = rs54!open_bal
          End If
         End If
         
        Sqlqry55 = "Update acct_mas set Close_bal=" & X + Tpdc & " where acct_code='103501'"
        ws.BeginTrans
        db.Execute (Sqlqry55)
        ws.CommitTrans
       
     End If
     
           
            
End Sub
Private Sub DOSUPPPOST()
   Ttlsupp = 0
   Sqlqry57 = " select * from Supp_Fin order by Supp_no"
   Set rs54 = db.OpenRecordset(Sqlqry57, dbOpenDynaset)
    If rs54.RecordCount = 0 Then
      MsgBox " Supplier Code not found in Supp_Fin"
      Exit Sub
    Else
      rs54.MoveFirst
      Do Until rs54.EOF
                Opbal = 0
                Opbrpt = 0
                Opbpmt = 0
                Opprpt = 0
                Opppmt = 0
                opjbd = 0
                Opjbc = 0
                opcrpt = 0
                opcpmt = 0
                opdbntdb = 0
                opdbntcr = 0
                Opcrntdb = 0
                Opcrntcr = 0
                Opcrsl = 0
                Opcrpr = 0
                Ttlopbal = 0
                Tpdc = 0
                
                
         Opbal = rs54!open_bal
        
        ' Cash Receipt before From date
        Sqlqry58 = "select sum(amount) from crpt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs54!Supp_no & "'"
        Set rs55 = db.OpenRecordset(Sqlqry58, dbOpenDynaset)
        If IsNull(rs55.Fields(0)) = False Then opcrpt = rs55.Fields(0)
                
        ' Cash Payment before From date
        Sqlqry59 = "select sum(amount) from cpmt_Tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs54!Supp_no & "'"
        Set rs56 = db.OpenRecordset(Sqlqry59, dbOpenDynaset)
        If IsNull(rs56.Fields(0)) = False Then opcpmt = rs56.Fields(0)
        
        ' Bank Receipt before From date
        Sqlqry60 = "select sum(amount) from brpt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs54!Supp_no & "'"
        Set rs57 = db.OpenRecordset(Sqlqry60, dbOpenDynaset)
        If IsNull(rs57.Fields(0)) = False Then Opbrpt = rs57.Fields(0)
        
        ' Bank Payment before From date
        Sqlqry61 = "select sum(amount) from bpmt_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs54!Supp_no & "'"
        Set rs58 = db.OpenRecordset(Sqlqry61, dbOpenDynaset)
        If IsNull(rs58.Fields(0)) = False Then Opbpmt = rs58.Fields(0)
                
        ' Pdc Receipts before From date
        Sqlqry63 = "Select sum(amount) from Prpt_mas1 where Cheque_dt<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_dt) and acct_code='" & rs54!Supp_no & "'"
        Set rs60 = db.OpenRecordset(Sqlqry63, dbOpenDynaset)
        If IsNull(rs60.Fields(0)) = False Then Opprpt = rs60.Fields(0)
        
       ' Pdc Payments before From Date
        Sqlqry65 = "Select sum(amount) from Ppmt_tra where  Cheque_dt<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_Dt) and acct_code='" & rs54!Supp_no & "'"
        Set rs62 = db.OpenRecordset(Sqlqry65, dbOpenDynaset)
        If IsNull(rs62.Fields(0)) = False Then Opppmt = rs62.Fields(0)
                     
       ' Journal Debit Amount before From Date
        Sqlqry66 = "select sum(damount) from Jrnl_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs54!Supp_no & "' and dc_code ='D'"
        Set rs63 = db.OpenRecordset(Sqlqry66, dbOpenDynaset)
        If IsNull(rs63.Fields(0)) = False Then opjbd = rs63.Fields(0)
                    
       ' Journal Credit Amount before From Date
        Sqlqry67 = "select sum(camount) from Jrnl_tra where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs54!Supp_no & "' and dc_code ='C'"
        Set rs64 = db.OpenRecordset(Sqlqry67, dbOpenDynaset)
        If IsNull(rs64.Fields(0)) = False Then Opjbc = rs64.Fields(0)
        
      ' Debit note (credit) before From Date
        Sqlqry68 = "select sum(amount) from debt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs54!Supp_no & "'"
        Set rs65 = db.OpenRecordset(Sqlqry68, dbOpenDynaset)
        If IsNull(rs65.Fields(0)) = False Then opdbntcr = rs65.Fields(0)
        
                 
       ' Debit note (debit)  before From Date
        Sqlqry69 = "select sum(amount) from debt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and cust_no='" & rs54!Supp_no & "'"
        Set rs66 = db.OpenRecordset(Sqlqry69, dbOpenDynaset)
        If IsNull(rs66.Fields(0)) = False Then opdbntdb = rs66.Fields(0)
        
                
       ' Credit note (debit) before From Date
        Sqlqry70 = "select sum(amount) from crdt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs54!Supp_no & "'"
        Set rs67 = db.OpenRecordset(Sqlqry70, dbOpenDynaset)
        If IsNull(rs67.Fields(0)) = False Then Opcrntdb = rs67.Fields(0)
        
                     
       ' Credit note (credit)  before From Date
        Sqlqry71 = "select SUM(AMOUNT) from crdt_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and supp_no='" & rs54!Supp_no & "'"
        Set rs68 = db.OpenRecordset(Sqlqry71, dbOpenDynaset)
        If IsNull(rs68.Fields(0)) = False Then Opcrntcr = rs68.Fields(0)
        
                  
        ' Credit Sales  before From Date
        Sqlqry72 = "select SUM(Net_AMOUNT) from bo_mas where invoice_date<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and agency='" & findfirstfixup(Trim(rs54!Supp_name)) & "'"
        Set rs69 = db.OpenRecordset(Sqlqry72, dbOpenDynaset)
        If IsNull(rs69.Fields(0)) = False Then Opcrsl = rs69.Fields(0) * convertion
        
                                   
        ' Opening balance credit Purchases  before From Date
        Sqlqry73 = "select SUM(gAMOUNT) from crpr_mas where tdate<=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and supp_no='" & rs54!Supp_no & "'"
        Set rs70 = db.OpenRecordset(Sqlqry73, dbOpenDynaset)
        If IsNull(rs70.Fields(0)) = False Then Opcrpr = rs70.Fields(0)
               
        Ttlopbal = Opbal - opcrpt + opcpmt - Opbrpt + Opbpmt - Opprpt + Opppmt + opjbd - Opjbc + opdbntdb - opdbntcr _
                    + Opcrntdb - Opcrntcr + Opcrsl - Opcrpr
        
        Ttlsupp = Ttlsupp + Ttlopbal
     
     rs54.MoveNext
     Loop
     
       ' Pending Post Dated Cheques Paid
        'Sqlqry74 = "select SUM(TTL_AMOUNT) from ppmt_mas where isnull(posting_dt) AND CHEQUE_DT>=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# "
        Sqlqry74 = "select SUM(TTL_AMOUNT) from ppmt_mas where isnull(posting_dt)"
        Set rs71 = db.OpenRecordset(Sqlqry74, dbOpenDynaset)
        If IsNull(rs71.Fields(0)) = False Then Tpdc = rs71.Fields(0)
        
        Ttlsupp = Ttlsupp + Tpdc
          
        Sqlqry75 = "Update acct_mas set close_bal=" & Ttlsupp & " where acct_code ='202000'"
        ws.BeginTrans
        db.Execute (Sqlqry75)
        ws.CommitTrans
        X = 0
        Sqlqry75 = "Select * from Acct_Mas Where Acct_code='202100'"
        Set rs71 = db.OpenRecordset(Sqlqry75, dbOpenDynaset)
         If rs71.RecordCount <> 0 Then
          rs71.MoveFirst
          If IsNull(rs71!open_bal) = True Then
            X = 0
          Else
            X = rs71!open_bal
          End If
         End If
        Sqlqry76 = "Update acct_mas set Close_bal=" & X - Tpdc & " where acct_code='202100'"
        ws.BeginTrans
        db.Execute (Sqlqry76)
        ws.CommitTrans
        
     End If
 
End Sub
Private Sub doclosnil()
        Sqlqry77 = "Update acct_mas set close_bal=" & 0 & ""
        ws.BeginTrans
        db.Execute (Sqlqry77)
        ws.CommitTrans
End Sub

Private Sub txtdatefrom_LostFocus()
If IsDate(txtdatefrom.TextWithMask) = False Then
      MsgBox "Invalid Date from ", vbInformation, "Invalid Entry"
      txtdatefrom.SetFocus
      SendKeys "{Home} + {End}"
    End If
End Sub
