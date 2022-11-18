VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmCashBalRep 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Cash  Report - DHS"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   3735
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   6255
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00FFFF80&
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
         Left            =   1440
         Picture         =   "frmCashBalRep.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Back <<"
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
         Left            =   3600
         Picture         =   "frmCashBalRep.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdclear 
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
         Left            =   2520
         Picture         =   "frmCashBalRep.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2640
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   5040
         Top             =   3120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   600
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
         Left            =   2880
         TabIndex        =   7
         Top             =   1200
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
         BorderColor     =   &H00404080&
         BorderWidth     =   3
         X1              =   0
         X2              =   6240
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   840
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmCashBalRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim rs As Recordset
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
Dim X
Dim frdate As Date
Dim toDate As Date
Dim Ttlopbal As Currency
Dim Opbal As Currency
Dim opsal As Currency
Dim oppur As Currency
Dim oprec As Currency
Dim oppay As Currency
Dim opjbd As Currency
Dim Opjcr As Currency
Dim Opbrpt As Currency
Dim Opbpmt As Currency

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
        
        Sqlqry = " Delete * from casreport"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        
Ttlopbal = 0
Opbal = 0
opsal = 0
oppur = 0
oprec = 0
oppay = 0
opjbd = 0
Opjcr = 0
Opbpmt = 0
Opbrpt = 0
       
       ' Op. from Account Master
        Sqlqry = " select * from acct_mas where acct_name='CASH IN HAND - DHS' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
          MsgBox " Description CASH IN HAND - DHS not found"
          Exit Sub
        Else
          rs.MoveFirst
          Opbal = rs!open_bal
        End If
        
        
        
   '* Cash sales not taken into consideration
   '''     ' Total Amount of Sales before From date
     '   Sqlqry1 = "select sum(namount) from casl_mas where Cash_code='" & rs!acct_code & "' and tdate< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "#"
     '   Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
     '   If IsNull(rs1.Fields(0)) = False Then opsal = rs1.Fields(0)
        
        ' Total Amount of Purchase before From date
        Sqlqry2 = "select sum(tra_amount) from capr_mas where Cash_Code ='" & rs!acct_code & "' and tcurrency='DHS' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs2.Fields(0)) = False Then oppur = rs2.Fields(0)
        
        
       ' Total cash Receipts before From date
        Sqlqry3 = "select sum(tra_amount) from crpt_mas where Cash_Code ='" & rs!acct_code & "' and tcurrency='DHS' and  tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs3 = db.OpenRecordset(Sqlqry3, dbOpenDynaset)
        If IsNull(rs3.Fields(0)) = False Then oprec = rs3.Fields(0)
        
        ' Total cash Payments before To date
        Sqlqry4 = "select sum(tra_amount) from cpmt_mas where Cash_Code ='" & rs!acct_code & "'and tcurrency='DHS' and  tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs4 = db.OpenRecordset(Sqlqry4, dbOpenDynaset)
        If IsNull(rs4.Fields(0)) = False Then oppay = rs4.Fields(0)
        
        ' Journal Debit Amount before From Date
        Sqlqry14 = "select sum(tra_damount) from Jrnl_tra where Acct_Code ='" & rs!acct_code & "' and tcurrency='DHS' and tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and dc_code ='D'"
        Set rs9 = db.OpenRecordset(Sqlqry14, dbOpenDynaset)
        If IsNull(rs9.Fields(0)) = False Then opjbd = rs9.Fields(0)
        
       ' Journal Credit Amount before From Date
        Sqlqry15 = "select sum(tra_camount) from Jrnl_tra where Acct_Code ='" & rs!acct_code & "' and tcurrency='DHS' and  tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and dc_code ='C'"
        Set rs10 = db.OpenRecordset(Sqlqry15, dbOpenDynaset)
        If IsNull(rs10.Fields(0)) = False Then Opjcr = rs10.Fields(0)
        
        ' Bank Payment  debit Before before From Date
        sqlqry16 = "select sum(tra_amount) from bpmt_tra where Acct_Code ='" & rs!acct_code & "' and tcurrency='DHS' and  tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs11 = db.OpenRecordset(sqlqry16, dbOpenDynaset)
        If IsNull(rs11.Fields(0)) = False Then Opbpmt = rs11.Fields(0)
         
        ' Bank Receipt Credit Before before From Date
        sqlqry17 = "select sum(tra_amount) from brpt_tra where Acct_Code ='" & rs!acct_code & "' and tcurrency='DHS' and  tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs12 = db.OpenRecordset(sqlqry17, dbOpenDynaset)
        If IsNull(rs12.Fields(0)) = False Then Opbrpt = rs12.Fields(0)
               
        Ttlopbal = Opbal + opsal + oprec + Opbpmt - Opbrpt - oppur - oppay + opjbd - Opjcr
        Sqlqry5 = "Insert into casreport values(" & 0 & ",'','" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','Opening Balance'," & Trim(Ttlopbal) & "," & 0 & ")"
        ws.BeginTrans
        db.Execute (Sqlqry5)
        ws.CommitTrans
        
        ' Cash Sales after From date and before to date
        'Sqlqry6 = "select * from casl_mas where Cash_Code ='" & rs!acct_code & "' and  tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#"
        'Set rs5 = db.OpenRecordset(Sqlqry6, dbOpenDynaset)
        'If rs5.RecordCount <> 0 Then
        ' rs5.MoveFirst
        ' Do Until rs5.EOF
        '  Sqlqry7 = "Insert into casreport values('" & rs5!VOUC_NO & "','" & rs5!vouc_type & "','" & Trim(rs5!tDate) & "','Cash Sale'," & Trim(rs5!namount) & "," & 0 & ")"
        '  ws.BeginTrans
        '  db.Execute (Sqlqry7)
        '  ws.CommitTrans
        '  rs5.MoveNext
        ' Loop
        'End If
        
        ' Cash Purchases after From date and before to date
        Sqlqry8 = "select * from capr_mas where Cash_Code ='" & rs!acct_code & "' and tcurrency='DHS' and tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs6 = db.OpenRecordset(Sqlqry8, dbOpenDynaset)
        If rs6.RecordCount <> 0 Then
         rs6.MoveFirst
         Do Until rs6.EOF
          Sqlqry9 = "Insert into casreport values('" & rs6!VOUC_NO & "','" & rs6!vouc_type & "','" & Trim(rs6!tDate) & "','Cash Purchase'," & 0 & "," & Trim(rs6!tra_amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry9)
          ws.CommitTrans
          rs6.MoveNext
         Loop
        End If
        
        ' Cash Receipts after From date and before to date
        Sqlqry10 = "select * from crpt_mas where Cash_Code ='" & rs!acct_code & "' and tcurrency='DHS' and  tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs7 = db.OpenRecordset(Sqlqry10, dbOpenDynaset)
        If rs7.RecordCount <> 0 Then
         rs7.MoveFirst
         Do Until rs7.EOF
          SQLQRY11 = "Insert into casreport values('" & rs7!VOUC_NO & "','" & rs7!vouc_type & "','" & Trim(rs7!tDate) & "','" & findfirstfixup(Trim(rs7!Description)) & "'," & Trim(rs7!tra_amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (SQLQRY11)
          ws.CommitTrans
          rs7.MoveNext
         Loop
        End If
        
        ' Cash Payments after From date and before to date
        SQLQRY12 = "select * from cpmt_mas where Cash_Code ='" & rs!acct_code & "' and tcurrency='DHS' and  tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs8 = db.OpenRecordset(SQLQRY12, dbOpenDynaset)
        If rs8.RecordCount <> 0 Then
         rs8.MoveFirst
         Do Until rs8.EOF
          Sqlqry13 = "Insert into casreport values('" & rs8!VOUC_NO & "','" & rs8!vouc_type & "','" & Trim(rs8!tDate) & "','" & findfirstfixup(Trim(rs8!Description)) & "'," & 0 & "," & Trim(rs8!tra_amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry13)
          ws.CommitTrans
          rs8.MoveNext
         Loop
        End If
        
       ' Bank Payments after From date and before to date
        sqlqry17 = "select * from bpmt_tra where acct_Code ='" & rs!acct_code & "' and tcurrency='DHS' and  tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs9 = db.OpenRecordset(sqlqry17, dbOpenDynaset)
        If rs9.RecordCount <> 0 Then
         rs9.MoveFirst
         Do Until rs9.EOF
          sqlqry18 = "Insert into casreport values('" & rs9!VOUC_NO & "','" & rs9!vouc_type & "','" & Trim(rs9!tDate) & "','" & findfirstfixup(Trim(rs9!Description)) & "'," & Trim(rs9!tra_amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (sqlqry18)
          ws.CommitTrans
          rs9.MoveNext
         Loop
        End If
        
        ' Bank Receipts after From date and before to date
        sqlqry19 = "select * from brpt_tra where acct_Code ='" & rs!acct_code & "' and tcurrency='DHS' and  tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs10 = db.OpenRecordset(sqlqry19, dbOpenDynaset)
        If rs10.RecordCount <> 0 Then
         rs10.MoveFirst
         Do Until rs10.EOF
          sqlqry20 = "Insert into casreport values('" & rs10!VOUC_NO & "','" & rs10!vouc_type & "','" & Trim(rs10!tDate) & "','" & findfirstfixup(Trim(rs10!Description)) & "'," & 0 & "," & Trim(rs10!tra_amount) & ")"
          ws.BeginTrans
          db.Execute (sqlqry20)
          ws.CommitTrans
          rs10.MoveNext
         Loop
        End If
     
     
        ' Journal Debit after From date and before to date
        sqlqry16 = "select * from jrnl_tra where Acct_Code ='" & rs!acct_code & "' and tcurrency='DHS' and  tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and dc_code='D' "
        Set rs11 = db.OpenRecordset(sqlqry16, dbOpenDynaset)
         If rs11.RecordCount <> 0 Then
          rs11.MoveFirst
          Do Until rs11.EOF
          sqlqry17 = "Insert into casreport values('" & rs11!VOUC_NO & "','" & rs11!vouc_type & "','" & Trim(rs11!tDate) & "','" & findfirstfixup(Trim(rs11!Description)) & "'," & Trim(rs11!tra_damount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (sqlqry17)
          ws.CommitTrans
          rs11.MoveNext
         Loop
        End If
        
        ' Journal Credit after From date and before to date
        sqlqry18 = "select * from Jrnl_tra where Acct_Code ='" & rs!acct_code & "' and tcurrency='DHS' and tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and dc_code='C'"
        Set rs12 = db.OpenRecordset(sqlqry18, dbOpenDynaset)
        If rs12.RecordCount <> 0 Then
         rs12.MoveFirst
        Do Until rs12.EOF
          sqlqry19 = "Insert into casreport values('" & rs12!VOUC_NO & "','" & rs12!vouc_type & "','" & Trim(rs12!tDate) & "','" & findfirstfixup(Trim(rs12!Description)) & "'," & 0 & "," & Trim(rs12!tra_Camount) & ")"
          ws.BeginTrans
          db.Execute (sqlqry19)
          ws.CommitTrans
          rs12.MoveNext
         Loop
        End If
     
     With CrystalReport1
      .DataFiles(0) = App.Path & "\misov.mdb"
      .ReportFileName = App.Path & "\CasReport.rpt"
      .Formulas(0) = "zzz='" & " From " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
      .Formulas(1) = "Curtype='DHS'"
      .WindowState = crptMaximized
     .Action = 1
     End With
     
        
    Else
        MsgBox "Improper Dates", vbDefaultButton1, "Invalid entry"
        Exit Sub
  End If
End Sub

Private Sub Form_Load()
 txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
 txtdateto.TextWithMask = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdateto.SetFocus
End Sub

Private Sub txtdatefrom_LostFocus()
If IsDate(txtdatefrom.TextWithMask) = False Then
   MsgBox "Invalid Date From", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdDisplay.SetFocus
End Sub


Private Sub textclear()
txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
txtdateto.TextWithMask = Format(Now, "dd/mm/yyyy")
End Sub

Private Function ValidateData()
ValidateData = False
If IsDate(txtdatefrom.TextWithMask) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
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

Private Sub txtdateto_LostFocus()
If IsDate(txtdateto.TextWithMask) = False Then
   MsgBox "Invalid Date To", vbInformation, "Invalid Entry"
   txtdateto.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub
