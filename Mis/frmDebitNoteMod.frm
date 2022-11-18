VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmDebitNoteMod 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Debit Note Modification"
   ClientHeight    =   8775
   ClientLeft      =   30
   ClientTop       =   345
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Debit Note  - Modification"
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
      Height          =   8055
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   11175
      Begin VB.TextBox txtConvRate 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   390
         Left            =   7365
         TabIndex        =   19
         Top             =   1800
         Width           =   1380
      End
      Begin VB.ComboBox cboCurrency 
         BackColor       =   &H80000018&
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
         Height          =   360
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   390
         Left            =   4245
         TabIndex        =   3
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ListBox lstVoucNo 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1020
         Left            =   1440
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtdesc 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   350
         Left            =   1440
         MaxLength       =   200
         TabIndex        =   7
         Top             =   5760
         Width           =   9495
      End
      Begin VB.TextBox txtRef 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   350
         Left            =   1440
         TabIndex        =   4
         Top             =   2400
         Width           =   7335
      End
      Begin VB.ListBox lstDebitedTo 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1980
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   4935
      End
      Begin VB.ListBox lstCreditedTo 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1980
         Left            =   5760
         TabIndex        =   6
         Top             =   3360
         Width           =   5175
      End
      Begin VB.TextBox txtDesc1 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   350
         Left            =   1440
         MaxLength       =   200
         TabIndex        =   8
         Top             =   6240
         Width           =   9495
      End
      Begin VB.CommandButton CmdModify 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Modify"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3600
         Picture         =   "frmDebitNoteMod.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7080
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   4800
         Picture         =   "frmDebitNoteMod.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7080
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6000
         Picture         =   "frmDebitNoteMod.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7080
         Width           =   1215
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   6960
         Top             =   7320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin PVMaskEditLib.PVMaskEdit txtdate 
         Height          =   375
         Left            =   4200
         TabIndex        =   1
         Top             =   600
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   244
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Remarks"
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
         Left            =   120
         TabIndex        =   23
         Top             =   6360
         Width           =   945
      End
      Begin VB.Label lblConvRate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Conv. Rate"
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
         Height          =   240
         Left            =   6105
         TabIndex        =   22
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Currency"
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
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1170
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Total Amount"
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
         Height          =   240
         Left            =   2805
         TabIndex        =   20
         Top             =   1920
         Width           =   1380
      End
      Begin VB.Line Line4 
         X1              =   11160
         X2              =   -120
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Line Line3 
         X1              =   5520
         X2              =   5520
         Y1              =   3000
         Y2              =   5520
      End
      Begin VB.Line Line2 
         X1              =   11160
         X2              =   -120
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   11160
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Voucher No."
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
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Debited To"
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
         Left            =   1200
         TabIndex        =   17
         Top             =   3120
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Date"
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
         Left            =   3360
         TabIndex        =   16
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Reference"
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
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Credited To"
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
         Left            =   8040
         TabIndex        =   14
         Top             =   3120
         Width           =   1245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Description "
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
         Left            =   120
         TabIndex        =   13
         Top             =   5880
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmDebitNoteMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim rs3 As Recordset
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim Sqlqry3 As String
Dim X
Dim y
Dim Z
Dim i
Dim j

Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtAmount.SetFocus
End Sub

Private Sub cboCurrency_LostFocus()
 If cboCurrency.Text = "USD" Then
     lblConvRate.Visible = True
     txtConvRate.Visible = True
     txtConvRate.Text = ""
     txtConvRate.TabIndex = 4
  Else
     lblConvRate.Visible = False
     txtConvRate.Visible = False
     txtConvRate.TabIndex = 12
     txtConvRate.Text = 1
 End If
End Sub
Private Sub cmdBack_Click()
 Unload Me
End Sub
Private Sub cmdClear_Click()
 textclear
End Sub

Private Sub Cmdmodify_Click()
Dim ctype As String
 If ValidateData = True Then
  
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       Sqlqry = " Update debt_mas set tdate=#" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "#," & _
                       " acct_code = '" & Mid(lstCreditedTo, 1, 6) & "'," & _
                       " acct_name ='" & Mid(lstCreditedTo, 12, 35) & "'," & _
                       " cust_no ='" & Mid(lstDebitedTo, 1, 4) & "'," & _
                       " Cust_name ='" & Trim(Mid(lstDebitedTo, 9, 35)) & "'," & _
                       " Ref_no ='" & findfirstfixup(UCase(Trim(txtRef))) & "'," & _
                       " Description ='" & findfirstfixup(UCase(Trim(txtdesc))) & "'," & _
                       " Description1='" & findfirstfixup(UCase(Trim(txtDesc1))) & "'," & _
                       " Tcurrency='" & Trim(cboCurrency) & "'," & _
                       " Tconvertion=" & Val(txtConvRate) & "," & _
                       " Tra_Amount=" & Val(txtAmount) & "," & _
                       " Amount =" & Val(txtAmount.Text) * Val(txtConvRate.Text) & " where vouc_no=" & Val(lstVoucNo) & ""
                 
       ws.BeginTrans
       db.Execute (Sqlqry)
       ws.CommitTrans
      ctype = cboCurrency.Text
  MsgBox " Record is Modified", vbInformation, "Status"
  Dim X
  X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
  If X = vbYes Then
  If ctype = "DHS" Then
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\DbntVou.rpt"
        CrystalReport1.SelectionFormula = "{Debt_Mas.Vouc_no}=" & Val(lstVoucNo.Text) & ""
        CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtAmount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
   Else
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\DbntVou.rpt"
        CrystalReport1.SelectionFormula = "{Debt_Mas.Vouc_no}=" & Val(lstVoucNo.Text) & ""
        CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(txtAmount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
  End If
  End If
  
  textclear
  End If
  
End Sub

Private Sub Form_Load()
 txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
 cboCurrency.AddItem "DHS"
 cboCurrency.AddItem "USD"
 lblConvRate.Visible = False
 txtConvRate.Visible = False
 txtConvRate.Text = 1
 PopulateVoucher
 PopulateAcctSuppCust
 PopulateAcctSuppCust1
 End Sub
Private Sub PopulateVoucher()
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Select Vouc_No from DEBT_MAS where status='N' order by vouc_no"
  Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
  lstVoucNo.Clear
  
 If rs.RecordCount <> 0 Then
    rs.MoveFirst
    Do Until rs.EOF
     lstVoucNo.AddItem rs!vouc_no
     rs.MoveNext
    Loop
 End If
End Sub

Private Sub PopulateAcctSuppCust()
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Select * from agndtls order by agentname"
 Sqlqry1 = "Select * from Supp_fin order by Supp_name"
 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
 Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
 

 lstDebitedTo.Clear

 If rs.RecordCount = 0 Then
    MsgBox "No Records found in the Agency Register"
 Else
    rs.MoveFirst
    Do Until rs.EOF
      lstDebitedTo.AddItem "AGNC" & "  :  " & rs!agentname
      rs.MoveNext
    Loop
 End If

If rs1.RecordCount = 0 Then
    MsgBox "No Records found in the Supplier Master"
Else
   rs1.MoveFirst
   Do Until rs1.EOF
      lstDebitedTo.AddItem rs1!Supp_no & "  :  " & rs1!Supp_name
      rs1.MoveNext
   Loop
End If

End Sub

Private Function ValidateData()

ValidateData = False
If IsDate(txtdate.TextWithMask) = False Then
  MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
  txtdate.SetFocus
  SendKeys "{Home} + {End}"
  Exit Function
ElseIf txtConvRate.Text = "" Then
  MsgBox "Enter Convertion Rate - - cannot be zero", vbInformation, "Invalid Entry"
  txtConvRate.SetFocus
  Exit Function
ElseIf txtAmount.Text = "" Or IsNumeric(txtAmount) = False Then
  MsgBox "Invalid Amount", vbInformation, "Invalid Entry"
  txtAmount.SetFocus
  Exit Function
ElseIf lstDebitedTo.SelCount = 0 Then
  MsgBox "Select Code to be Debited", vbInformation, "Invalid Entry"
  lstDebitedTo.SetFocus
  Exit Function
ElseIf lstCreditedTo.SelCount = 0 Then
  MsgBox "Select Code to be Credited", vbInformation, "Invalid Entry"
  lstCreditedTo.SetFocus
  Exit Function
ElseIf txtdesc.Text = "" Or IsNumeric(txtdesc) = True Then
  MsgBox "Invalid Description", vbInformation, "Invalid Entry"
  txtdesc.SetFocus
  Exit Function

Else
  ValidateData = True
End If
End Function

Private Function textclear()
     txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
     cboCurrency.ListIndex = -1
     txtConvRate.Visible = False
     lblConvRate.Visible = False
     lstDebitedTo.ListIndex = 0
     lstCreditedTo.ListIndex = 0
     txtRef.Text = ""
     txtdesc.Text = ""
     txtDesc1.Text = ""
     txtAmount.Text = ""
     txtdate.SetFocus
     
End Function

Private Sub PopulateAcctSuppCust1()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Acct_mas order by acct_code"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    
    lstCreditedTo.Clear
    
    If rs.RecordCount = 0 Then
        MsgBox "No Records found in the Account Register"
     Else
       rs.MoveFirst
       Do Until rs.EOF
          lstCreditedTo.AddItem rs!acct_code & "  :  " & rs!acct_name
          rs.MoveNext
       Loop
    End If
    
End Sub
Private Sub lstCreditedTo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdesc.SetFocus
End Sub
Private Sub lstDebitedTo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then lstCreditedTo.SetFocus
End Sub
Private Sub lstVoucNo_Click()
 txtdate.SetFocus
End Sub
Private Sub lstVoucNo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdate.SetFocus
End Sub

Private Sub lstVoucNo_LostFocus()
    Dim i
    Dim X
    Dim y
    Dim Z
    Dim U
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Val(lstVoucNo.Text)
        
        Sqlqry = " Select * from debt_mas Where Vouc_no= " & i
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
         If rs.RecordCount <> 0 Then
          
           txtdate.TextWithMask = Format(rs!tDate, "dd/mm/yyyy")
           If rs!tcurrency = "DHS" Then
             cboCurrency.ListIndex = 0
             txtConvRate.Text = rs!tconvertion
           Else
             cboCurrency.ListIndex = 1
             txtConvRate.Text = rs!tconvertion
           End If
                                
           txtAmount = rs!tra_amount
           txtRef = rs!Ref_no
           txtdesc = rs!Description
           txtDesc1 = rs!Description1
         
         Sqlqry3 = "Select acct_Code,acct_name from Acct_mas where acct_code='" & rs!acct_code & "' order by acct_code"
         Set rs3 = db.OpenRecordset(Sqlqry3, dbOpenDynaset)
           If rs3.RecordCount <> 0 Then
                  lstCreditedTo.Text = rs3!acct_code & "  :  " & rs3!acct_name
           End If
           
           
         Sqlqry = "Select * from Agndtls where agentname='" & Trim(rs!cust_name) & "' order by agentname"
         Set rs1 = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs1.RecordCount <> 0 Then
                  lstDebitedTo.Text = "AGNC" & "  :  " & rs1!agentname
           Else
             Sqlqry1 = "Select Supp_no,Supp_name from supp_Fin where supp_no='" & Trim(rs!cust_no) & "' order by supp_no"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                lstDebitedTo.Text = Trim(rs1!Supp_no) & "  :  " & rs1!Supp_name
               End If
           End If
           
           txtdate.SetFocus
         End If
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
   If cboCurrency.Text = "USD" Then
     If KeyAscii = 13 Then txtConvRate.SetFocus
   Else
     If KeyAscii = 13 Then txtRef.SetFocus
     txtConvRate.Text = 1
   End If
End Sub

Private Sub txtConvRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRef.SetFocus
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboCurrency.SetFocus
End Sub

Private Sub txtdate_LostFocus()
If IsDate(txtdate.TextWithMask) = False Then
      MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
      txtdate.SetFocus
      SendKeys "{Home} + {End}"
End If
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtDesc1.SetFocus
End Sub
Private Sub txtDesc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdModify.SetFocus
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstDebitedTo.SetFocus
End Sub
