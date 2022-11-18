VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "PVMASK.OCX"
Begin VB.Form frmCreditNoteAddOrg 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Credit Note Addition"
   ClientHeight    =   8775
   ClientLeft      =   -60
   ClientTop       =   285
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "  Credit Note - New Entry"
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
      Height          =   7935
      Left            =   600
      TabIndex        =   12
      Top             =   360
      Width           =   10095
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
         Height          =   315
         Left            =   4365
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
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
         Height          =   315
         Left            =   7485
         TabIndex        =   11
         Top             =   1320
         Width           =   1380
      End
      Begin VB.TextBox txtdesc 
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
         ForeColor       =   &H00404040&
         Height          =   350
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   6
         Top             =   5400
         Width           =   8175
      End
      Begin VB.TextBox txtRef 
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
         ForeColor       =   &H00404040&
         Height          =   350
         Left            =   1560
         TabIndex        =   3
         Top             =   1920
         Width           =   7335
      End
      Begin VB.ListBox lstDebitedTo 
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
         ForeColor       =   &H00404040&
         Height          =   1860
         Left            =   5280
         TabIndex        =   4
         Top             =   2880
         Width           =   4695
      End
      Begin VB.ListBox lstCreditedTo 
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
         ForeColor       =   &H00404040&
         Height          =   1860
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   4575
      End
      Begin VB.TextBox txtDesc1 
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
         ForeColor       =   &H00404040&
         Height          =   350
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   7
         Top             =   5880
         Width           =   8175
      End
      Begin VB.CommandButton CmdAdd 
         BackColor       =   &H00FFFF00&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   2520
         Picture         =   "frmCreditNoteAddOrg.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFF00&
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   3840
         Picture         =   "frmCreditNoteAddOrg.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6720
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFF00&
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   5040
         Picture         =   "frmCreditNoteAddOrg.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6720
         Width           =   1215
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   480
         Top             =   6600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin PVMaskEditLib.PVMaskEdit txtdate 
         Height          =   375
         Left            =   4320
         TabIndex        =   0
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
         Left            =   240
         TabIndex        =   23
         Top             =   6000
         Width           =   945
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
         Left            =   2925
         TabIndex        =   22
         Top             =   1440
         Width           =   1380
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
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   1170
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
         Left            =   6225
         TabIndex        =   20
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Line Line4 
         X1              =   10080
         X2              =   0
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line Line3 
         X1              =   5040
         X2              =   5040
         Y1              =   2520
         Y2              =   5040
      End
      Begin VB.Line Line2 
         X1              =   10080
         X2              =   0
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10080
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   19
         Top             =   720
         Width           =   1365
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
         Left            =   6240
         TabIndex        =   18
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   3720
         TabIndex        =   17
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
         Left            =   240
         TabIndex        =   16
         Top             =   2040
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
         Left            =   840
         TabIndex        =   15
         Top             =   2640
         Width           =   1245
      End
      Begin VB.Label lblVoucNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   600
         Width           =   1035
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
         Left            =   240
         TabIndex        =   13
         Top             =   5400
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmCreditNoteAddOrg"
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
     txtConvRate.TabIndex = 3
    Else
     lblConvRate.Visible = False
     txtConvRate.Visible = False
     txtConvRate.Text = 1
     txtConvRate.TabIndex = 11
    End If
End Sub
Private Sub cmdBack_Click()
 Unload Me
End Sub
Private Sub cmdClear_Click()
  textclear
End Sub

Private Sub cmdadd_Click()
 If ValidateData = True Then
  
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       Sqlqry = " Insert into Crdt_MasOrg values('" & lblVoucNo & "','CNT','" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & Mid(lstDebitedTo, 1, 6) & "','" _
                                     & Mid(lstDebitedTo, 12, 35) & "','" _
                                     & Mid(lstCreditedTo, 1, 4) & "','" _
                                     & Trim(Mid(lstCreditedTo, 9, 35)) & "','" _
                                     & findfirstfixup(UCase(Trim(txtRef))) & "','" _
                                     & findfirstfixup(UCase(Trim(txtdesc))) & "','" _
                                     & findfirstfixup(UCase(Trim(txtDesc1))) & "','" _
                                     & Trim(cboCurrency) & "'," _
                                     & Val(txtConvRate) & "," _
                                     & Val(txtAmount) & "," _
                                    & Val(txtAmount) * Val(txtConvRate) & ",'N')"
       ws.BeginTrans
       db.Execute (Sqlqry)
       ws.CommitTrans
        
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Update docu_mas set doc_no='" & lblVoucNo & "' where doc_type='CNT'"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
  lblVoucNo = lblVoucNo + 1
  MsgBox " Record is inserted", vbInformation, "Status"
  Dim X As Integer
   X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
  If X = vbYes Then
   CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
   CrystalReport1.ReportFileName = App.Path & "\crntvou.rpt"
   CrystalReport1.SelectionFormula = "{Crdt_MasOrg.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
   CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtAmount)) & " Only" & "'"
   CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
   CrystalReport1.WindowState = crptMaximized
   CrystalReport1.Action = 1
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
     AutoIncrementVoucher
     PopulateAcctSuppCust
     PopulateAcctSuppCust1
 End Sub

Private Sub AutoIncrementVoucher()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select DOC_no from DOCU_MAS WHERE DOC_TYPE='CNT'"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
If rs.RecordCount = 0 Then
   MsgBox "Document type 'CNT' not found"
   Exit Sub
Else
   rs.MoveLast
   lblVoucNo = Val(rs!doc_no) + 1
End If
End Sub

Private Sub PopulateAcctSuppCust()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select * from agndtls order by agentname"
Sqlqry1 = "Select * from Supp_fin order by Supp_name"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)

lstCreditedTo.Clear

If rs.RecordCount = 0 Then
    MsgBox "No Records found in the Agency Register"
Else
   rs.MoveFirst
   Do Until rs.EOF
      lstCreditedTo.AddItem "AGNC" & "  :  " & rs!agentname
      rs.MoveNext
   Loop
End If

If rs1.RecordCount = 0 Then
    MsgBox "No Records found in the Supplier Master"
Else
   rs1.MoveFirst
   Do Until rs1.EOF
      lstCreditedTo.AddItem rs1!Supp_no & "  :  " & rs1!Supp_name
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
ElseIf lstCreditedTo.SelCount = 0 Then
  MsgBox "Select Code to be Debited", vbInformation, "Invalid Entry"
  lstCreditedTo.SetFocus
  Exit Function
ElseIf lstDebitedTo.SelCount = 0 Then
  MsgBox "Select Code to be Credited", vbInformation, "Invalid Entry"
  lstDebitedTo.SetFocus
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
     lstCreditedTo.ListIndex = 0
     lstDebitedTo.ListIndex = 0
     cboCurrency.ListIndex = -1
     txtConvRate.Text = ""
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

lstDebitedTo.Clear

If rs.RecordCount = 0 Then
    MsgBox "No Records found in the Account Register"
 Else
   rs.MoveFirst
   Do Until rs.EOF
      lstDebitedTo.AddItem rs!acct_code & "  :  " & rs!acct_name
      rs.MoveNext
   Loop
End If
  
End Sub

Private Sub lstDebitedTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdesc.SetFocus
End Sub

Private Sub lstCreditedTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstDebitedTo.SetFocus
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
   MsgBox "Invalid Date", vbInformation, "Invalid Entry"
   txtdate.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtDesc1.SetFocus
End Sub
Private Sub txtDesc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdAdd.SetFocus
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstCreditedTo.SetFocus
End Sub

Private Sub txtRef_LostFocus()
 Dim crdcur
 
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Select * from bo_mas where serial_no='" & Mid(txtRef, 1, 7) & "'"
 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
 If rs.RecordCount <> 0 Then
   rs.MoveFirst
   crdcur = rs!tcurrency
 End If
 
 If cboCurrency <> crdcur Then
   MsgBox " Reference Booking order booked in different currency"
   cboCurrency.SetFocus
   Exit Sub
 End If


  
End Sub
