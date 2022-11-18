VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "PVMASK.OCX"
Begin VB.Form frmCustomerListRepDhs 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Agencies List"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5760
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "         Statement of  Agencies Outstanding in DHS      "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4095
      Left            =   2520
      TabIndex        =   4
      Top             =   1680
      Width           =   6015
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   5535
         Begin VB.CommandButton cmdAging 
            BackColor       =   &H0080C0FF&
            Caption         =   "Aging analysis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1680
            Picture         =   "frmCustomerListRepDHS.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdDisplay 
            BackColor       =   &H0080C0FF&
            Caption         =   "&Print Preview"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   600
            Picture         =   "frmCustomerListRepDHS.frx":0442
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2760
            Picture         =   "frmCustomerListRepDHS.frx":0544
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3840
            Picture         =   "frmCustomerListRepDHS.frx":0986
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   1695
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   4335
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
            Left            =   1920
            TabIndex        =   10
            Top             =   480
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
            Left            =   1920
            TabIndex        =   11
            Top             =   1080
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
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enter Date From"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enter Date To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   240
            TabIndex        =   7
            Top             =   1200
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmCustomerListRepDhs"
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
Dim tdamount As Currency
Dim tcamount As Currency
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
Dim xxxx
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
     .ReportFileName = App.Path & "\CustAgeRep.rpt"
     .SelectionFormula = "{custlstrep.gr_bal}<>0"
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
  Dim cno As String
 
  If ValidateData = True Then
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           
   Sqlqry53 = " Delete * from Custlstrep"
           ws.BeginTrans
           db.Execute (Sqlqry53)
           ws.CommitTrans
   
   Sqlqry = " select * from agndtls order by agentname"
   Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
   If rs.RecordCount = 0 Then
     MsgBox " Agency not found in the Agency Register"
     Exit Sub
    Else
     rs.MoveFirst
      Do Until rs.EOF
            Sqlqry = " Delete * from CustReport"
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
                cno = ""
        
 ' Starting
          
          If IsNull(rs!op_DHS) = True Then
             Opbal = 0
          Else
             Opbal = rs!op_DHS
          End If
          
          If IsNull(rs!op_USD) = True Then
            Opbal = Opbal
          Else
            Opbal = Opbal + (rs!op_USD * convertion)
          End If
         cno = Trim(rs!agentname)
                
      
        Sqlqry1 = "select sum(amount) from crpt_tra where ( tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and acct_name='" & findfirstfixup(Trim(cno)) & "')"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then opcrpt = rs1.Fields(0)
        
      
       ' Cash Payment before From date
      
        Sqlqry2 = "select sum(amount) from cpmt_tra where ( tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and acct_name='" & findfirstfixup(Trim(cno)) & "')"
        Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs2.Fields(0)) = False Then opcpmt = rs2.Fields(0)
            
        '' Bank Receipt before From date
        Sqlqry3 = "select Sum(amount) from brpt_tra where (tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(Trim(cno)) & "')"
        Set rs3 = db.OpenRecordset(Sqlqry3, dbOpenDynaset)
        If IsNull(rs3.Fields(0)) = False Then Opbrpt = rs3.Fields(0)
        
        
        '' Bank Payment before From date
        Sqlqry4 = "select Sum(amount) from bpmt_tra where (tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and acct_name='" & findfirstfixup(rs!agentname) & "')"
        Set rs4 = db.OpenRecordset(Sqlqry4, dbOpenDynaset)
        If IsNull(rs4.Fields(0)) = False Then Opbpmt = rs4.Fields(0)
        
        
       ''Pdc Receipts before From date
        Sqlqry5 = "Select sum(amount) from prpt_mas1 where Cheque_dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_dt) and acct_name='" & findfirstfixup(rs!agentname) & "'"
        Set rs6 = db.OpenRecordset(Sqlqry5, dbOpenDynaset)
        If IsNull(rs6.Fields(0)) = False Then Opprpt = rs6.Fields(0)
        
        
      ''Pdc Payments before From Date
        Sqlqry7 = "select * from Ppmt_mas where Cheque_dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and not isnull(posting_Dt)"
        Set rs7 = db.OpenRecordset(Sqlqry7, dbOpenDynaset)
        If rs7.RecordCount <> 0 Then
         rs7.MoveFirst
          Do Until rs7.EOF
           Sqlqry8 = "Select sum(amount) from Ppmt_tra where Vouc_no=" & Val(rs7!vouc_no) & " and acct_name='" & findfirstfixup(rs!agentname) & "'"
           Set rs8 = db.OpenRecordset(Sqlqry8, dbOpenDynaset)
           If IsNull(rs8.Fields(0)) = False Then Opppmt = rs8.Fields(0)
           rs7.MoveNext
          Loop
        End If
        
        
            
       '' Journal Debit tra_amount before From Date
        Sqlqry9 = "select Sum(damount) from Jrnl_tra  where (tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and acct_name='" & findfirstfixup(rs!agentname) & "' and dc_code ='D' ) "
        Set rs9 = db.OpenRecordset(Sqlqry9, dbOpenDynaset)
        If IsNull(rs9.Fields(0)) = False Then opjbd = rs9.Fields(0)
        
                         
       '' Journal Credit tra_amount before From Date
        Sqlqry10 = " select sum(camount) From Jrnl_tra where (tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and acct_name='" & findfirstfixup(rs!agentname) & "' and dc_code ='C') "
        Set rs10 = db.OpenRecordset(Sqlqry10, dbOpenDynaset)
        If IsNull(rs10.Fields(0)) = False Then Opjbc = rs10.Fields(0)
        
       ' Opening balance  debit note (credit) before From Date
        SQLQRY11 = "select Sum(amount) from debt_mas where (tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs!agentname) & "')"
        Set rs11 = db.OpenRecordset(SQLQRY11, dbOpenDynaset)
        If IsNull(rs11.Fields(0)) = False Then opdbntcr = rs11.Fields(0)
        
         
       '' Opening balance debit note (debit)  before From Date
        SQLQRY12 = "select Sum(amount) from debt_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and  cust_name='" & findfirstfixup(rs!agentname) & "'"
        Set rs12 = db.OpenRecordset(SQLQRY12, dbOpenDynaset)
        If IsNull(rs12.Fields(0)) = False Then opdbntdb = rs12.Fields(0)
        
        
       '' Opening balance  Credit note (debit) before From Date
        Sqlqry13 = "select Sum(amount) from crdt_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs!agentname) & "'"
        Set rs13 = db.OpenRecordset(Sqlqry13, dbOpenDynaset)
        If IsNull(rs13.Fields(0)) = False Then Opcrntdb = rs13.Fields(0)
         
       '' Opening balance credit note (credit)  before From Date
        Sqlqry14 = "select sum(amount) from crdt_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and SUPP_NAME='" & findfirstfixup(rs!agentname) & "'"
        Set rs14 = db.OpenRecordset(Sqlqry14, dbOpenDynaset)
        If IsNull(rs14.Fields(0)) = False Then Opcrntcr = rs14.Fields(0)
        
              
        ' Opening balance credit Sales  before From Date
        Sqlqry15 = "select Sum(net_amount) from BO_mas where INVOICE_date<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#   and AGENCY='" & findfirstfixup(rs!agentname) & "' and cancell='N'"
        Set rs15 = db.OpenRecordset(Sqlqry15, dbOpenDynaset)
        If IsNull(rs15.Fields(0)) = False Then Opcrsl = rs15.Fields(0)
        
                             
        ' Opening balance credit Purchases  before From Date
        sqlqry16 = "select sum(amountdhs) from crpr_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and SUPP_NAME='" & findfirstfixup(rs!agentname) & "'"
        Set rs16 = db.OpenRecordset(sqlqry16, dbOpenDynaset)
        If IsNull(rs16.Fields(0)) = False Then Opcrpr = rs16.Fields(0)
        
        
        Ttlopbal = Opbal - opcrpt + opcpmt - Opbrpt + Opbpmt - Opprpt + Opppmt + opjbd - Opjbc + opdbntdb - opdbntcr _
                    + Opcrntdb - Opcrntcr + Opcrsl - Opcrpr
                    
        If Ttlopbal >= 0 Then
            sqlqry17 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "'," & 0 & ",'','" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','Opening Balance'," & Trim(Ttlopbal) & "," & 0 & ")"
            ws.BeginTrans
            db.Execute (sqlqry17)
            ws.CommitTrans
        Else
            sqlqry17 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "'," & 0 & ",'','" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','Opening Balance'," & 0 & "," & Trim(Ttlopbal) & ")"
            ws.BeginTrans
            db.Execute (sqlqry17)
            ws.CommitTrans
        End If
            
        ' Cash Receipt after From date and before to date
        sqlqry18 = "select * from crpt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs!agentname) & "'"
        Set rs18 = db.OpenRecordset(sqlqry18, dbOpenDynaset)
        If rs18.RecordCount <> 0 Then
         rs18.MoveFirst
         Do Until rs18.EOF
                sqlqry19 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs18!vouc_no & "','" & rs18!vouc_type & "','" & Trim(rs18!tDate) & "','" & findfirstfixup(Trim(rs18!Description)) & "'," & 0 & "," & Val(rs18!Amount) & ")"
                ws.BeginTrans
                db.Execute (sqlqry19)
                ws.CommitTrans
         
          rs18.MoveNext
         Loop
        End If
        
        ' Cash Payment after From date and before to date
        sqlqry20 = "select * from Cpmt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs!agentname) & "'"
        Set rs19 = db.OpenRecordset(sqlqry20, dbOpenDynaset)
        If rs19.RecordCount <> 0 Then
         rs19.MoveFirst
         Do Until rs19.EOF
            Sqlqry21 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs19!vouc_no & "','" & rs19!vouc_type & "','" & Trim(rs19!tDate) & "','" & findfirstfixup(Trim(rs19!Description)) & "'," & Val(rs19!Amount) & "," & 0 & ")"
            ws.BeginTrans
            db.Execute (Sqlqry21)
            ws.CommitTrans
      
          rs19.MoveNext
         Loop
        End If
                
        ' Bank Receipt after From date and before to date
        Sqlqry22 = "select * from brpt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs!agentname) & "'"
        Set rs20 = db.OpenRecordset(Sqlqry22, dbOpenDynaset)
        If rs20.RecordCount <> 0 Then
         rs20.MoveFirst
         Do Until rs20.EOF
                Sqlqry23 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs20!vouc_no & "','" & rs20!vouc_type & "','" & Trim(rs20!tDate) & "','" & findfirstfixup(Trim(rs20!Description)) & "'," & 0 & "," & Val(rs20!Amount) & ")"
                ws.BeginTrans
                db.Execute (Sqlqry23)
                ws.CommitTrans
      
          rs20.MoveNext
         Loop
        End If
        
        ' Bank Payment after From date and before to date
        Sqlqry24 = "select * from bpmt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs!agentname) & "'"
        Set rs21 = db.OpenRecordset(Sqlqry24, dbOpenDynaset)
        If rs21.RecordCount <> 0 Then
         rs21.MoveFirst
         Do Until rs21.EOF
               Sqlqry25 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs21!vouc_no & "','" & rs21!vouc_type & "','" & Trim(rs21!tDate) & "','" & findfirstfixup(Trim(rs21!Description)) & "'," & Val(rs21!Amount) & "," & 0 & ")"
               ws.BeginTrans
               db.Execute (Sqlqry25)
               ws.CommitTrans
                 
           rs21.MoveNext
         Loop
        End If
        
       ' Pdc Receipts after From date and before to date
        Sqlqry26 = "select * from prpt_mas1 where Cheque_dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Cheque_dt<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(Posting_Dt) and acct_name ='" & findfirstfixup(rs!agentname) & "'"
        Set rs23 = db.OpenRecordset(Sqlqry26, dbOpenDynaset)
        If rs23.RecordCount <> 0 Then
         rs23.MoveFirst
          Do Until rs23.EOF
                 Sqlqry28 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs23!vouc_no & "','" & rs23!vouc_type & "','" & Trim(rs23!Cheque_Dt) & "','" & findfirstfixup(Trim(rs23!Description)) & "'," & 0 & "," & Val(rs23!Amount) & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry28)
                 ws.CommitTrans
                       
                rs23.MoveNext
           Loop
         End If
          
        
        ' Pdc Payments after From date and before to date
        Sqlqry29 = "select * from Ppmt_mas where Cheque_dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#  and not Isnull(Posting_Dt) "
        Set rs24 = db.OpenRecordset(Sqlqry29, dbOpenDynaset)
        If rs24.RecordCount <> 0 Then
         rs24.MoveFirst
         Do Until rs24.EOF
            Sqlqry30 = "Select * from Ppmt_tra where Vouc_no=" & Val(rs24!vouc_no) & " and acct_name ='" & findfirstfixup(rs!agentname) & "'"
            Set rs25 = db.OpenRecordset(Sqlqry30, dbOpenDynaset)
             If rs25.RecordCount <> 0 Then
               rs25.MoveFirst
                Do Until rs25.EOF
      
                        Sqlqry31 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs25!vouc_no & "','" & rs25!vouc_type & "','" & Trim(rs24!Cheque_Dt) & "','" & findfirstfixup(Trim(rs25!Description)) & "'," & Val(rs25!Amount) & "," & 0 & ")"
                        ws.BeginTrans
                        db.Execute (Sqlqry31)
                        ws.CommitTrans
      
                 rs25.MoveNext
                Loop
              End If
          rs24.MoveNext
          Loop
         End If
         
       ' Journal Debit after From date and before to date
        Sqlqry32 = "select * from jrnl_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs!agentname) & "' and dc_code='D' "
        Set rs26 = db.OpenRecordset(Sqlqry32, dbOpenDynaset)
         If rs26.RecordCount <> 0 Then
          rs26.MoveFirst
          Do Until rs26.EOF
      
                Sqlqry33 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs26!vouc_no & "','" & rs26!vouc_type & "','" & Trim(rs26!tDate) & "','" & findfirstfixup(Trim(rs26!Description)) & "'," & Val(rs26!damount) & "," & 0 & ")"
                ws.BeginTrans
                db.Execute (Sqlqry33)
                ws.CommitTrans
      
          rs26.MoveNext
         Loop
        End If
        
    ' Journal Credit after From date and before to date
        Sqlqry33 = "select * from Jrnl_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs!agentname) & "' and dc_code='C'"
        Set rs27 = db.OpenRecordset(Sqlqry33, dbOpenDynaset)
        If rs27.RecordCount <> 0 Then
         rs27.MoveFirst
        Do Until rs27.EOF
      
            Sqlqry34 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs27!vouc_no & "','" & rs27!vouc_type & "','" & Trim(rs27!tDate) & "','" & findfirstfixup(Trim(rs27!Description)) & "'," & 0 & "," & Val(rs27!Camount) & ")"
            ws.BeginTrans
            db.Execute (Sqlqry34)
            ws.CommitTrans
                  
          rs27.MoveNext
         Loop
        End If
        
    ' DebitNote - credit after From date and before to date
        Sqlqry35 = "select * from debt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs!agentname) & "'"
        Set rs28 = db.OpenRecordset(Sqlqry35, dbOpenDynaset)
        If rs28.RecordCount <> 0 Then
          rs28.MoveFirst
         Do Until rs28.EOF
      
            Sqlqry36 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs28!vouc_no & "','" & rs28!vouc_type & "','" & Trim(rs28!tDate) & "','" & findfirstfixup(Trim(rs28!Description)) & "'," & 0 & "," & Val(rs28!Amount) & ")"
            ws.BeginTrans
            db.Execute (Sqlqry36)
            ws.CommitTrans
      
          rs28.MoveNext
         Loop
        End If
        
        ' DebitNote - debit after From date and before to date
        Sqlqry37 = "select * from debt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and cust_name='" & findfirstfixup(rs!agentname) & "'"
        Set rs29 = db.OpenRecordset(Sqlqry37, dbOpenDynaset)
        If rs29.RecordCount <> 0 Then
          rs29.MoveFirst
         Do Until rs29.EOF
      
            Sqlqry38 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs29!vouc_no & "','" & rs29!vouc_type & "','" & Trim(rs29!tDate) & "','" & findfirstfixup(Trim(rs29!Description)) & "'," & Val(rs29!Amount) & "," & 0 & ")"
            ws.BeginTrans
            db.Execute (Sqlqry38)
            ws.CommitTrans
      

          rs29.MoveNext
         Loop
        End If
        
        ' CreditNote - Credit after From date and before to date
        Sqlqry38 = "select * from crdt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and SUPP_NAME='" & findfirstfixup(rs!agentname) & "'"
        Set rs30 = db.OpenRecordset(Sqlqry38, dbOpenDynaset)
         If rs30.RecordCount <> 0 Then
           rs30.MoveFirst
         Do Until rs30.EOF
      
                Sqlqry39 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs30!vouc_no & "','" & rs30!vouc_type & "','" & Trim(rs30!tDate) & "','" & findfirstfixup(Trim(rs30!Description)) & "'," & 0 & "," & Val(rs30!Amount) & ")"
                ws.BeginTrans
                db.Execute (Sqlqry39)
                ws.CommitTrans
      
           rs30.MoveNext
         Loop
        End If
        
    ' Credit Note - Debit after From date and before to date
        Sqlqry40 = "select * from crdt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & findfirstfixup(rs!agentname) & "'"
        Set rs31 = db.OpenRecordset(Sqlqry40, dbOpenDynaset)
        If rs31.RecordCount <> 0 Then
         rs31.MoveFirst
        Do Until rs31.EOF
      
            Sqlqry41 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs31!vouc_no & "','" & rs31!vouc_type & "','" & Trim(rs31!tDate) & "','" & findfirstfixup(Trim(rs31!Description)) & "'," & Val(rs31!Amount) & "," & 0 & ")"
            ws.BeginTrans
            db.Execute (Sqlqry41)
            ws.CommitTrans
      

          rs31.MoveNext
         Loop
        End If
       
    ' Credit Sale after From date and before to date
        Sqlqry42 = "select * from BO_mas where INVOICE_date>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Invoice_date<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and AGENCY='" & findfirstfixup(rs!agentname) & "' and cancell='N'"
    
        Set rs32 = db.OpenRecordset(Sqlqry42, dbOpenDynaset)
        If rs32.RecordCount <> 0 Then
         rs32.MoveFirst
        Do Until rs32.EOF
    
            Sqlqry43 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs32!serial_no & "','INV','" & Trim(rs32!invoice_date) & "','" & findfirstfixup(Trim(rs32!bo_ref)) & " " & Trim(rs32!media) & "'," & Val(rs32!NET_Amount) & "," & 0 & ")"
            ws.BeginTrans
            db.Execute (Sqlqry43)
            ws.CommitTrans
          rs32.MoveNext
          
         Loop
         End If
        '
    ' Credit Purchase after From date and before to date
        Sqlqry44 = "select * from crpr_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and SUPP_NAME='" & findfirstfixup(rs!agentname) & "'"
        Set rs33 = db.OpenRecordset(Sqlqry44, dbOpenDynaset)
        If rs33.RecordCount <> 0 Then
         rs33.MoveFirst
         Do Until rs33.EOF
                Sqlqry45 = "Insert into CustReport values('" & findfirstfixup(rs!agentname) & "','" & rs33!vouc_no & "','" & rs33!vouc_type & "','" & Trim(rs33!tDate) & "','Purchases'," & 0 & "," & Val(rs33!amountdhs) & ")"
                ws.BeginTrans
                db.Execute (Sqlqry45)
                ws.CommitTrans
           rs33.MoveNext
         Loop
        End If
        
   ' Pdc Receipts after From date and before to date
        Sqlqry46 = "select * from prpt_mas1 where isnull(Posting_Dt) AND TCURRENCY='DHS' and acct_name ='" & findfirstfixup(rs!agentname) & "'"
        Set rs35 = db.OpenRecordset(Sqlqry46, dbOpenDynaset)
        If rs35.RecordCount <> 0 Then
         rs35.MoveFirst
            Do Until rs35.EOF
              ttlpdc = ttlpdc + rs35!tra_amount
               rs35.MoveNext
             Loop
        End If
        
        Sqlqry46 = "select * from prpt_mas1 where isnull(Posting_Dt) AND TCURRENCY='USD' and acct_name ='" & findfirstfixup(rs!agentname) & "'"
        Set rs35 = db.OpenRecordset(Sqlqry46, dbOpenDynaset)
        If rs35.RecordCount <> 0 Then
         rs35.MoveFirst
            Do Until rs35.EOF
              ttlpdc = ttlpdc + (rs35!Amount)
               rs35.MoveNext
             Loop
        End If
      
     Sqlqry48 = "Select SUM(Damount) from custreport Where date between datevalue(now()) and datevalue(now())-30"
     Set rs36 = db.OpenRecordset(Sqlqry48, dbOpenDynaset)
     If IsNull(rs36.Fields(0)) = False Then dbb30 = rs36.Fields(0)
       
      
      Sqlqry49 = "Select SUM(Damount) from custReport where date between DateValue(Now()) - 31 and DateValue(Now()) - 60"
      Set rs37 = db.OpenRecordset(Sqlqry49, dbOpenDynaset)
      If IsNull(rs37.Fields(0)) = False Then db30 = rs37.Fields(0)
      
      Sqlqry50 = "Select SUM(Damount) from custReport where date between DateValue(Now()) - 61 and DateValue(Now()) - 90"
      Set rs38 = db.OpenRecordset(Sqlqry50, dbOpenDynaset)
      If IsNull(rs38.Fields(0)) = False Then db60 = rs38.Fields(0)
            
      Sqlqry51 = "Select SUM(Damount) from custReport where date between DateValue(Now()) - 91 and DateValue(now())-366"
      Set rs39 = db.OpenRecordset(Sqlqry51, dbOpenDynaset)
      If IsNull(rs39.Fields(0)) = False Then db90 = rs39.Fields(0)
      
      tdamount = 0
      tcamount = 0
      B30 = 0
      A30 = 0
      A60 = 0
      A90 = 0
      Sqlqry52 = "Select * from custReport order by DATE"
      Set rs40 = db.OpenRecordset(Sqlqry52, dbOpenDynaset)
      If rs40.RecordCount <> 0 Then
       rs40.MoveFirst
        Do Until rs40.EOF
          tdamount = tdamount + rs40!damount
          tcamount = tcamount + rs40!Camount
          rs40.MoveNext
        Loop
      End If
      
     tgr = tdamount - tcamount
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
 
 
       Sqlqry52 = "Insert into Custlstrep values('xxxx','" & Trim(rs!agentname) & "'," & Ttlopbal & "," & tdamount & "," & tcamount & "," & tgr & ", " & ttlpdc & "," & tnt & "," & B30 & "," & A30 & "," & A60 & "," & A90 & ")"
       ws.BeginTrans
       db.Execute (Sqlqry52)
       ws.CommitTrans
   
   rs.MoveNext
   Loop
   End If
   
   
    With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\CustlstRep.rpt"
     .Formulas(0) = "zzz='" & "From" & Trim(txtdatefrom) & "  To " & Trim(txtdateto) & "'"
     .SelectionFormula = "{Custlstrep.gr_bal}<>0"
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
If IsDate(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
   
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

If IsDate(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid To Date", vbInformation, "Invalid Entry"
   txtdateto.SetFocus
   SendKeys " {Home} + {End} "
End If

End Sub
