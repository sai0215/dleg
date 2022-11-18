VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmSupplierListRep 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Supplier List"
   ClientHeight    =   8775
   ClientLeft      =   15
   ClientTop       =   300
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8160
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "List of Suppliers"
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
      Height          =   4815
      Left            =   2640
      TabIndex        =   4
      Top             =   1320
      Width           =   6015
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Height          =   1215
         Left            =   240
         TabIndex        =   9
         Top             =   3120
         Width           =   5535
         Begin VB.CommandButton cmdAging 
            BackColor       =   &H0080C0FF&
            Caption         =   "Aging "
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
            Left            =   1560
            Picture         =   "frmSupplierListRep.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdDisplay 
            BackColor       =   &H0080C0FF&
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
            Left            =   480
            Picture         =   "frmSupplierListRep.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H0080C0FF&
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
            Left            =   2640
            Picture         =   "frmSupplierListRep.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdBack 
            BackColor       =   &H0080C0FF&
            Caption         =   "<< &Back"
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
            Picture         =   "frmSupplierListRep.frx":0986
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   2415
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   4695
         Begin VB.Frame Frame3 
            Caption         =   "Frame2"
            Height          =   15
            Left            =   0
            TabIndex        =   6
            Top             =   2520
            Width           =   4335
         End
         Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
            Height          =   375
            Left            =   2160
            TabIndex        =   10
            Top             =   720
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
         Begin PVMaskEditLib.PVMaskEdit txtdateto 
            Height          =   375
            Left            =   2160
            TabIndex        =   11
            Top             =   1440
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Date From"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   600
            TabIndex        =   8
            Top             =   840
            Width           =   1365
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Date To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   600
            TabIndex        =   7
            Top             =   1560
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmSupplierListRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim Opbal As Currency
Dim opcrpt As Currency
Dim opcpmt As Currency
Dim Opbrpt As Currency
Dim Opbpmt As Currency
Dim Opprpt As Currency
Dim Opppmt As Currency
Dim opjbd As Currency
Dim Opjbc As Currency
Dim opdbntdb As Currency
Dim opdbntcr As Currency
Dim Opcrntdb As Currency
Dim Opcrntcr As Currency
Dim Opcrsl As Currency
Dim Opcrpr As Currency
Dim Ttlopbal As Currency
Dim tcrpt As Currency
Dim tcpmt As Currency
Dim tbrpt As Currency
Dim tbpmt As Currency
Dim tprpt As Currency
Dim tppmt As Currency
Dim tjrnd As Currency
Dim tjrnc As Currency
Dim tdbntc As Currency
Dim tdbntd As Currency
Dim tcrntc As Currency
Dim tcrntd As Currency
Dim tcrsl As Currency
Dim tcrpr As Currency
Dim treceipts As Currency
Dim tsales As Currency
Dim tgr As Currency
Dim tnt As Currency
Dim ttlpdc As Currency
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
Dim vaddr
Dim vcity
Dim vcountry
Dim vtel
Dim vfax
Dim dbb30 As Currency
Dim db30 As Currency
Dim db60 As Currency
Dim db90 As Currency
Dim B30 As Currency
Dim A30 As Currency
Dim A60 As Currency
Dim A90 As Currency

Private Sub cmdAging_Click()
  With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\sUPPAGEREP.RPT"
     .SelectionFormula = "{Supplstrep.gr_bal}<>0"
     .WindowState = crptMaximized
     .Action = 1
   End With
       
End Sub

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
           
      Sqlqry53 = " Delete * from Supplstrep"
            ws.BeginTrans
            db.Execute (Sqlqry53)
            ws.CommitTrans
   
   Sqlqry = "Select * from Supp_fin order by Supp_no"
   Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
   If rs.RecordCount = 0 Then
     MsgBox " Supplier Code not found in Supp_fin"
     Exit Sub
    Else
     rs.MoveFirst
      Do Until rs.EOF
            Sqlqry = " Delete * from Suppreport1"
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
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
                ttlpdc = 0
                tcrpt = 0
                tcpmt = 0
                tbrpt = 0
                tbpmt = 0
                tprpt = 0
                tppmt = 0
                tjrnd = 0
                tjrnc = 0
                tdbntc = 0
                tdbntd = 0
                tcrntc = 0
                tcrntd = 0
                tcrsl = 0
                tcrpr = 0
                treceipts = 0
                tsales = 0
                tgr = 0
                tnt = 0
                dbb30 = 0
                db30 = 0
                db60 = 0
                db90 = 0
                B30 = 0
                A30 = 0
                A60 = 0
                A90 = 0
           Opbal = rs!open_bal
         
    ' Cash Receipt before From date
        Sqlqry1 = "select sum(amount) from crpt_tra where tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!Supp_no & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then opcrpt = rs1.Fields(0)
        
        
    ' Cash Payment before From date
        Sqlqry2 = "select sum(amount) from cpmt_Tra where tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!Supp_no & "'"
        Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then opcpmt = rs1.Fields(0)
                
    ' Bank Receipt before From date
        Sqlqry3 = "select sum(amount) from brpt_tra where tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!Supp_no & "'"
        Set rs3 = db.OpenRecordset(Sqlqry3, dbOpenDynaset)
        If IsNull(rs3.Fields(0)) = False Then Opbrpt = rs3.Fields(0)
            
    ' Bank Payment before From date
        Sqlqry4 = "select sum(amount) from bpmt_tra where tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!Supp_no & "'"
        Set rs4 = db.OpenRecordset(Sqlqry4, dbOpenDynaset)
        If IsNull(rs4.Fields(0)) = False Then Opbpmt = rs4.Fields(0)
             
    ' Pdc Receipts before From date
        Sqlqry6 = "Select sum(amount) from Prpt_mas1 where Cheque_Dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_dt) and acct_code='" & rs!Supp_no & "'"
        Set rs6 = db.OpenRecordset(Sqlqry6, dbOpenDynaset)
        If IsNull(rs6.Fields(0)) = False Then Opprpt = rs6.Fields(0)
        
    'Pdc Payments before From Date
        Sqlqry8 = "Select sum(amount) from Ppmt_tra where Cheque_Dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_dt) and acct_code='" & rs!Supp_no & "'"
        Set rs8 = db.OpenRecordset(Sqlqry8, dbOpenDynaset)
        If IsNull(rs8.Fields(0)) = False Then Opppmt = rs8.Fields(0)
          
    ' Journal Debit Amount before From Date
        Sqlqry9 = "select sum(damount) from Jrnl_tra where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!Supp_no & "' and dc_code ='D'"
        Set rs9 = db.OpenRecordset(Sqlqry9, dbOpenDynaset)
        If IsNull(rs9.Fields(0)) = False Then opjbd = rs9.Fields(0)
        
    ' Journal Credit Amount before From Date
        Sqlqry10 = "select sum(camount) from Jrnl_tra where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!Supp_no & "' and dc_code ='C'"
        Set rs10 = db.OpenRecordset(Sqlqry10, dbOpenDynaset)
        If IsNull(rs10.Fields(0)) = False Then Opjbc = rs10.Fields(0)
        
     ' Debit note (credit) before From Date
        SQLQRY11 = "select sum(amount) from debt_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!Supp_no & "'"
        Set rs11 = db.OpenRecordset(SQLQRY11, dbOpenDynaset)
        If IsNull(rs11.Fields(0)) = False Then opdbntcr = rs11.Fields(0)
             
     ' Debit note (debit)  before From Date
        SQLQRY12 = "select Sum(amount) from debt_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and cust_no='" & rs!Supp_no & "'"
        Set rs12 = db.OpenRecordset(SQLQRY12, dbOpenDynaset)
        If IsNull(rs12.Fields(0)) = False Then opdbntdb = rs12.Fields(0)
           
      ' Credit note (debit) before From Date
        Sqlqry13 = "select sum(amount) from crdt_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!Supp_no & "'"
        Set rs13 = db.OpenRecordset(Sqlqry13, dbOpenDynaset)
        If IsNull(rs12.Fields(0)) = False Then Opcrntdb = rs12.Fields(0)
         
      ' Credit note (credit)  before From Date
        Sqlqry14 = "select sum(Amount) from crdt_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and supp_no='" & rs!Supp_no & "'"
        Set rs14 = db.OpenRecordset(Sqlqry14, dbOpenDynaset)
        If IsNull(rs14.Fields(0)) = False Then Opcrntcr = rs14.Fields(0)
        
       ' Opening balance credit Purchases  before From Date
        sqlqry16 = "select sum(gamount) from crpr_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and supp_no='" & rs!Supp_no & "'"
        Set rs16 = db.OpenRecordset(sqlqry16, dbOpenDynaset)
        If IsNull(rs16.Fields(0)) = False Then Opcrpr = rs16.Fields(0)
         
          Ttlopbal = Opbal - opcrpt + opcpmt - Opbrpt + Opbpmt - Opprpt + Opppmt + opjbd - Opjbc + opdbntdb - opdbntcr _
                    + Opcrntdb - Opcrntcr + Opcrsl - Opcrpr
                    
       If Ttlopbal < 0 Then
         sqlqry19 = "Insert into Suppreport1 values('','',#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#,'Opening Balance'," & 0 & "," & Abs(Ttlopbal) & ")"
          ws.BeginTrans
          db.Execute (sqlqry19)
          ws.CommitTrans
       Else
         sqlqry19 = "Insert into Suppreport1 values('','',#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#,'Opening Balance'," & Ttlopbal & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (sqlqry19)
          ws.CommitTrans
       End If
       
       
        ' Cash Receipt after From date and before to date
        sqlqry18 = "select * from crpt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!Supp_no & "'"
        Set rs18 = db.OpenRecordset(sqlqry18, dbOpenDynaset)
        If rs18.RecordCount <> 0 Then
         rs18.MoveFirst
         Do Until rs18.EOF
          sqlqry19 = "Insert into Suppreport1 values('" & rs18!VOUC_NO & "','" & rs18!vouc_type & "','" & Trim(rs18!tDate) & "','" & Trim(rs18!Description) & "'," & 0 & "," & Trim(rs18!Amount) & ")"
          ws.BeginTrans
          db.Execute (sqlqry19)
          ws.CommitTrans
          tcrpt = tcrpt + rs18!Amount
          rs18.MoveNext
         Loop
        End If
        
        ' Cash Payment after From date and before to date
        sqlqry20 = "select * from Cpmt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!Supp_no & "'"
        Set rs19 = db.OpenRecordset(sqlqry20, dbOpenDynaset)
        If rs19.RecordCount <> 0 Then
         rs19.MoveFirst
         Do Until rs19.EOF
          Sqlqry21 = "Insert into Suppreport1 values('" & rs19!VOUC_NO & "','" & rs19!vouc_type & "','" & Trim(rs19!tDate) & "','" & Trim(rs19!Description) & "'," & Trim(rs19!Amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry21)
          ws.CommitTrans
          tcpmt = tcpmt + rs19!Amount
          rs19.MoveNext
         Loop
        End If
                
        ' Bank Receipt after From date and before to date
        Sqlqry22 = "select * from brpt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!Supp_no & "'"
        Set rs20 = db.OpenRecordset(Sqlqry22, dbOpenDynaset)
        If rs20.RecordCount <> 0 Then
         rs20.MoveFirst
         Do Until rs20.EOF
          Sqlqry23 = "Insert into Suppreport1 values('" & rs20!VOUC_NO & "','" & rs20!vouc_type & "','" & Trim(rs20!tDate) & "','" & Trim(rs20!Description) & "'," & 0 & "," & Trim(rs20!Amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry23)
          ws.CommitTrans
          tbrpt = tbrpt + rs20!Amount
          rs20.MoveNext
         Loop
        End If
        
        ' Bank Payment after From date and before to date
        Sqlqry24 = "select * from bpmt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & rs!Supp_no & "'"
        Set rs21 = db.OpenRecordset(Sqlqry24, dbOpenDynaset)
        If rs21.RecordCount <> 0 Then
         rs21.MoveFirst
         Do Until rs21.EOF
          Sqlqry25 = "Insert into Suppreport1 values('" & rs21!VOUC_NO & "','" & rs21!vouc_type & "','" & Trim(rs21!tDate) & "','" & Trim(rs21!Description) & "'," & Trim(rs21!Amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry25)
          ws.CommitTrans
           tbpmt = tbpmt + rs21!Amount
          rs21.MoveNext
         Loop
        End If
        
        ' Pdc Receipts after From date and before to date
           
         Sqlqry27 = "Select * from Prpt_mas1 where Cheque_Dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Cheque_dt<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#  and not Isnull(Posting_Dt) and acct_code ='" & rs!Supp_no & "'"
         Set rs23 = db.OpenRecordset(Sqlqry27, dbOpenDynaset)
              If rs23.RecordCount <> 0 Then
                rs23.MoveFirst
                 Do Until rs23.EOF
                   Sqlqry28 = "Insert into Suppreport1 values('" & rs23!VOUC_NO & "','" & rs23!vouc_type & "','" & Trim(rs23!Cheque_Dt) & "','" & Trim(rs23!Description) & "'," & 0 & "," & Trim(rs23!Amount) & ")"
                     ws.BeginTrans
                     db.Execute (Sqlqry28)
                     ws.CommitTrans
                     tprpt = tprpt + rs23!Amount
                    rs23.MoveNext
                 Loop
                End If
          
        
        ' Pdc Payments after From date and before to date
          Sqlqry30 = "Select * from Ppmt_tra where Cheque_Dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#  and not Isnull(Posting_Dt) and acct_code ='" & rs!Supp_no & "'"
          Set rs25 = db.OpenRecordset(Sqlqry30, dbOpenDynaset)
          If rs25.RecordCount <> 0 Then
            rs25.MoveFirst
            Do Until rs25.EOF
            Sqlqry31 = "Insert into Suppreport1 values('" & rs25!VOUC_NO & "','" & rs25!vouc_type & "','" & Trim(rs24!Cheque_Dt) & "','" & Trim(rs25!Description) & "'," & Trim(rs25!Amount) & "," & 0 & ")"
                ws.BeginTrans
                db.Execute (Sqlqry31)
                ws.CommitTrans
                tppmt = tppmt + rs25!Amount
                rs25.MoveNext
                Loop
         End If
          
         
       ' Journal Debit after From date and before to date
        Sqlqry32 = "select * from jrnl_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!Supp_no & "' and dc_code='D' "
        Set rs26 = db.OpenRecordset(Sqlqry32, dbOpenDynaset)
        If rs26.RecordCount <> 0 Then
          rs26.MoveFirst
          Do Until rs26.EOF
          Sqlqry33 = "Insert into Suppreport1 values('" & rs26!VOUC_NO & "','" & rs26!vouc_type & "','" & Trim(rs26!tDate) & "','" & Trim(rs26!Description) & "'," & Trim(rs26!damount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry33)
          ws.CommitTrans
          tjrnd = tjrnd + rs26!damount
          rs26.MoveNext
         Loop
        End If
        
        ' Journal Credit after From date and before to date
        Sqlqry33 = "select * from Jrnl_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!Supp_no & "' and dc_code='C'"
        Set rs27 = db.OpenRecordset(Sqlqry33, dbOpenDynaset)
        If rs27.RecordCount <> 0 Then
         rs27.MoveFirst
        Do Until rs27.EOF
          Sqlqry34 = "Insert into Suppreport1 values('" & rs27!VOUC_NO & "','" & rs27!vouc_type & "','" & Trim(rs27!tDate) & "','" & Trim(rs27!Description) & "'," & 0 & "," & Trim(rs27!camount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry34)
          ws.CommitTrans
          tjrnc = tjrnc + rs27!camount
          rs27.MoveNext
         Loop
        End If
        
     ' DebitNote - credit after From date and before to date
        Sqlqry35 = "select * from debt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!Supp_no & "'"
        Set rs28 = db.OpenRecordset(Sqlqry35, dbOpenDynaset)
        If rs28.RecordCount <> 0 Then
          rs28.MoveFirst
         Do Until rs28.EOF
          Sqlqry36 = "Insert into Suppreport1 values('" & rs28!VOUC_NO & "','" & rs28!vouc_type & "','" & Trim(rs28!tDate) & "','" & Trim(rs28!Description) & "'," & 0 & "," & Trim(rs28!Amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry36)
          ws.CommitTrans
          tdbntc = tdbntc + rs28!Amount
          rs28.MoveNext
         Loop
        End If
        
        ' DebitNote - debit after From date and before to date
        Sqlqry37 = "select * from debt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Cust_no='" & rs!Supp_no & "'"
        Set rs29 = db.OpenRecordset(Sqlqry37, dbOpenDynaset)
        If rs29.RecordCount <> 0 Then
          rs29.MoveFirst
         Do Until rs29.EOF
          Sqlqry38 = "Insert into Suppreport1 values('" & rs29!VOUC_NO & "','" & rs29!vouc_type & "','" & Trim(rs29!tDate) & "','" & Trim(rs29!Description) & "'," & Trim(rs29!Amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry38)
          ws.CommitTrans
          tdbntd = tdbntd + rs29!Amount
          rs29.MoveNext
         Loop
        End If
        
        Sqlqry38 = "select * from crdt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Supp_no='" & rs!Supp_no & "'"
        Set rs30 = db.OpenRecordset(Sqlqry38, dbOpenDynaset)
         If rs30.RecordCount <> 0 Then
           rs30.MoveFirst
         Do Until rs30.EOF
           Sqlqry39 = "Insert into Suppreport1 values('" & rs30!VOUC_NO & "','" & rs30!vouc_type & "','" & Trim(rs30!tDate) & "','" & Trim(rs30!Description) & "'," & 0 & "," & Trim(rs30!Amount) & ")"
           ws.BeginTrans
           db.Execute (Sqlqry39)
           ws.CommitTrans
           tcrntc = tcrntc + rs30!Amount
           rs30.MoveNext
         Loop
        End If
        
        ' Credit Note - Debit after From date and before to date
        Sqlqry40 = "select * from crdt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & rs!Supp_no & "'"
        Set rs31 = db.OpenRecordset(Sqlqry40, dbOpenDynaset)
        If rs31.RecordCount <> 0 Then
         rs31.MoveFirst
        Do Until rs31.EOF
          Sqlqry41 = "Insert into Suppreport1 values('" & rs31!VOUC_NO & "','" & rs31!vouc_type & "','" & Trim(rs31!tDate) & "','" & Trim(rs31!Description) & "'," & Trim(rs31!Amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry41)
          ws.CommitTrans
          tcrntd = tcrntd + rs31!Amount
          rs31.MoveNext
         Loop
        End If
        
        ' Credit Sale after From date and before to date
        'Sqlqry42 = "select * from crsl_mas where tdate>= #" & DateValue(Format(txtdatefrom.textwithmask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.textwithmask, "dd/mm/yyyy")) & "# and cust_no='" & rs!Supp_no & "'"
        'Set rs32 = db.OpenRecordset(Sqlqry42, dbOpenDynaset)
        'If rs32.RecordCount <> 0 Then
        ' rs32.MoveFirst
        ' Do Until rs32.EOF
        '  Sqlqry43 = "Insert into Suppreport1 values('" & rs32!VOUC_NO & "','" & rs32!vouc_type & "','" & Trim(rs32!tDate) & "','Sales'," & Trim(rs32!namount) & "," & 0 & ")"
        '  ws.BeginTrans
        '  db.Execute (Sqlqry43)
        '  ws.CommitTrans
        '  tcrsl = tcrsl + rs32!namount
        '  rs32.MoveNext
        ' Loop
       ' End If
        
     ' Credit Purchase after From date and before to date
        Sqlqry44 = "select * from crpr_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Supp_no='" & rs!Supp_no & "'"
        Set rs33 = db.OpenRecordset(Sqlqry44, dbOpenDynaset)
        If rs33.RecordCount <> 0 Then
         rs33.MoveFirst
        Do Until rs33.EOF
          Sqlqry45 = "Insert into Suppreport1 values('" & rs33!VOUC_NO & "','" & rs33!vouc_type & "','" & Trim(rs33!tDate) & "','Purchases'," & 0 & "," & Trim(rs33!gamount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry45)
          ws.CommitTrans
          tcrpr = tcrpr + rs33!gamount
          rs33.MoveNext
         Loop
        End If
        
    ' Pdc Dated after to date
      Sqlqry47 = "Select SUM(AMOUNT) from Ppmt_tra where isnull(posting_dt) and acct_code ='" & rs!Supp_no & "'"
      Set rs35 = db.OpenRecordset(Sqlqry47, dbOpenDynaset)
      If IsNull(rs35.Fields(0)) = False Then ttlpdc = rs35.Fields(0)
           
        
      Sqlqry48 = "Select sum(camount) from Suppreport1 Where date between datevalue(now()) and datevalue(now())-30"
      Set rs36 = db.OpenRecordset(Sqlqry48, dbOpenDynaset)
      If IsNull(rs36.Fields(0)) = False Then dbb30 = rs36.Fields(0)
             
      Sqlqry49 = "Select sum(camount) from Suppreport1 where date between DateValue(Now()) - 31 and DateValue(Now()) - 60"
      Set rs37 = db.OpenRecordset(Sqlqry49, dbOpenDynaset)
      If IsNull(rs37.Fields(0)) = False Then db30 = rs37.Fields(0)
         
      Sqlqry50 = "Select sum(camount) from Suppreport1 where date between DateValue(Now()) - 61 and DateValue(Now()) - 90"
      Set rs38 = db.OpenRecordset(Sqlqry50, dbOpenDynaset)
      If IsNull(rs38.Fields(0)) = False Then db60 = rs38.Fields(0)
      
      Sqlqry51 = "Select Sum(camount) from Suppreport1 where date between DateValue(Now()) - 91 and DateValue(Now()) - 1500"
      Set rs39 = db.OpenRecordset(Sqlqry51, dbOpenDynaset)
      If IsNull(rs39.Fields(0)) = False Then db90 = rs39.Fields(0)
      
      
      ' Total Payment
      treceipts = tcrpt - tcpmt + tbrpt - tbpmt + tprpt - tppmt
      
      'Total Purchases
      tsales = -tjrnd + tjrnc - tdbntd + tdbntc - tcrntd + tcrntc - tcrsl + tcrpr
      'tgr = Ttlopbal + tsales - treceipts
      tgr = -Ttlopbal + tsales + treceipts
      tnt = tgr - ttlpdc
     
      If tnt > dbb30 + db30 + db60 Then
          A90 = tnt - dbb30 - db30 - db60
      End If
      If tnt > dbb30 + db30 Then
          A60 = tnt - dbb30 - db30 - A90
      End If
      If tnt > dbb30 Then
          A30 = tnt - dbb30 - A60 - A90
      End If
      If tnt > 0 Then
          B30 = tnt - A90 - A60 - A30
      End If
      
           
    Sqlqry52 = "Insert into Supplstrep values('" & Trim(rs!Supp_no) & "','" & findfirstfixup(Trim(rs!Supp_name)) & "'," & -Ttlopbal & "," & tsales & "," & -treceipts & "," & tgr & ", " & ttlpdc & "," & tnt & "," & B30 & "," & A30 & "," & A60 & "," & A90 & ")"
     ws.BeginTrans
     db.Execute (Sqlqry52)
     ws.CommitTrans
   
   rs.MoveNext
   Loop
   End If
     
    With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\Supplstrep.rpt"
     .Formulas(0) = "zzz='" & " From  " & Trim(txtdatefrom.TextWithMask) & "  To  " & Trim(txtdateto.TextWithMask) & "'"
     .SelectionFormula = "{Supplstrep.gr_bal}<>0"
     .WindowState = crptMaximized
     .Action = 1
    End With
   
   Else
        MsgBox "Improper Dates", vbDefaultButton1, "Invalid entry"
        Exit Sub
  End If
End Sub
Private Sub Form_Load()
 txtdatefrom.TextWithMask = Format(Now, "DD/mm/yyyy")
 txtdateto.TextWithMask = Format(Now, "DD/mm/yyyy")
End Sub
Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtdateto.SetFocus
End Sub

Private Sub txtdatefrom_LostFocus()
If IsDate(txtdatefrom.TextWithMask) = False Then
      MsgBox "Invalid Date From", vbInformation, "Invalid Entry"
      txtdatefrom.SetFocus
      SendKeys "{Home} + {End}"
    End If
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdDisplay.SetFocus
End Sub

Private Function ValidateData()
 ValidateData = False

If IsDate(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsDate(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid To Date", vbInformation, "Invalid Entry"
   txtdateto.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
ValidateData = True
End If
End Function
Private Sub textclear()
 txtdatefrom.TextWithMask = Format(Now, "DD/mm/yyyy")
 txtdateto.TextWithMask = Format(Now, "DD/mm/YYYY")
End Sub
Private Sub txtdateto_LostFocus()
If IsDate(txtdateto.TextWithMask) = False Then
      MsgBox "Invalid Date To ", vbInformation, "Invalid Entry"
      txtdateto.SetFocus
      SendKeys "{Home} + {End}"
    End If
End Sub
