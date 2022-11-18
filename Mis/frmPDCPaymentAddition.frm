VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPdcPaymentAddition 
   BackColor       =   &H00C0C0FF&
   Caption         =   "PDC Payment Addition"
   ClientHeight    =   8775
   ClientLeft      =   90
   ClientTop       =   315
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "PDC Issue - New Entry"
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
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   10935
      Begin VB.TextBox txtChequeDt 
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
         Height          =   315
         Left            =   4320
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtChequeNO 
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
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboBankCode 
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
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtTtlAmount 
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
         Height          =   315
         Left            =   7320
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtDate 
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
         Height          =   315
         Left            =   4320
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtPaidTo 
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
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   1320
         Width           =   4695
      End
      Begin VB.TextBox txtAmount 
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
         Height          =   315
         Left            =   8160
         TabIndex        =   9
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtDesc 
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
         Height          =   315
         Left            =   4560
         TabIndex        =   8
         Top             =   3360
         Width           =   3495
      End
      Begin VB.ListBox lstAcctCode 
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
         Height          =   1260
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   3360
         Width           =   4335
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
         Left            =   4680
         Picture         =   "frmPDCPaymentAddition.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H0080C0FF&
         Caption         =   "<<&Back"
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
         Left            =   5640
         Picture         =   "frmPDCPaymentAddition.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Save"
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
         Left            =   3720
         Picture         =   "frmPDCPaymentAddition.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7080
         Width           =   975
      End
      Begin VB.TextBox txtTtlDesc 
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
         Height          =   315
         Left            =   1560
         MaxLength       =   150
         TabIndex        =   6
         Top             =   1800
         Width           =   4695
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4680
         Top             =   5280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1575
         Left            =   240
         TabIndex        =   14
         Top             =   4920
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   5
         Cols            =   4
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
         BackColorFixed  =   8388608
         BackColorBkg    =   8421376
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   10920
         X2              =   0
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cheque Dt."
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
         Left            =   3120
         TabIndex        =   28
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cheque No."
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
         TabIndex        =   27
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Bank Name"
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
         Left            =   5880
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   25
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Payee Name"
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
         TabIndex        =   24
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
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
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   5880
         TabIndex        =   23
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
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
         Left            =   3120
         TabIndex        =   22
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Account / Supplier / Customer Code"
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
         TabIndex        =   21
         Top             =   3120
         Width           =   3690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Description"
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
         Left            =   5640
         TabIndex        =   20
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Amount"
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
         Left            =   8520
         TabIndex        =   19
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label lblVoucNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label LblTtlAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   9120
         TabIndex        =   17
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label lblTtlAmt 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
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
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   7560
         TabIndex        =   16
         Top             =   6600
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Description"
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
         Top             =   1920
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmPdcPaymentAddition"
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
Private Sub cboBankCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPaidTo.SetFocus
End Sub
Private Sub cmdBack_Click()
 Unload Me
End Sub

Private Sub cmdClear_Click()
     textclear
     LblTtlAmount.Caption = ""
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from DUMPPMT1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     MSFlexGrid1.Clear
     txtTtlAmount.SetFocus
End Sub
Private Sub cmdSave_Click()
 If ValidateData = True Then
   If Val(txtTtlAmount.Text) = LblTtlAmount Then
         
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       Sqlqry = " Insert into PPMT_MAS values('" & lblVoucNo & "','PPT','" _
                                     & Format(txtDate, "dd/mm/yyyy") & "','" _
                                     & UCase(Trim(txtPaidTo)) & "','" _
                                     & Mid(cboBankCode, 1, 6) & "','" _
                                     & Mid(cboBankCode, 10, 25) & "','" _
                                     & UCase(Trim(txtChequeNO)) & "','" _
                                     & Format(txtChequeDt, "dd/mm/yyyy") & "',' ','" _
                                     & UCase(Trim(txtTtlDesc)) & "','" _
                                     & Trim(txtTtlAmount) & "','N')"
       ws.BeginTrans
       db.Execute (Sqlqry)
       ws.CommitTrans
        
    Sqlqry1 = "Select * from DUMPPMT1"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount = 0 Then
         MsgBox " Transactions are not recorded"
         Exit Sub
      Else
         rs.MoveFirst
         Do Until rs.EOF
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry3 = "Insert into PPMT_TRA values('" & rs!VOUC_NO & "','" & rs!vouc_type & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & rs!acct_code & "','" _
                                     & rs!acct_name & "','" _
                                     & rs!Description & "','" _
                                     & Format(txtChequeDt, "dd/mm/yyyy") & "',' ','" _
                                     & rs!amount & "')"

            ws.BeginTrans
            db.Execute (Sqlqry3)
            ws.CommitTrans
          rs.MoveNext
         Loop
       End If
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Update docu_mas set doc_no='" & lblVoucNo & "' where doc_type='PPT'"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     
     textclear
     lblVoucNo = lblVoucNo + 1
     
   Else
   MsgBox "Total amount is not equal to entered amount"
   Exit Sub
   End If
  End If
  MsgBox " Record is inserted", vbInformation, "Status"
  Dim X As Integer
  X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
  If X = vbYes Then
   CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
   CrystalReport1.ReportFileName = App.Path & "\ppmtVou.rpt"
   CrystalReport1.SelectionFormula = "{ppmt_tra.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
   CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtTtlAmount)) & " Only" & "'"
   CrystalReport1.WindowState = crptMaximized
   CrystalReport1.Action = 1
  End If
End Sub
Private Sub Form_Load()
 txtDate.Text = Format(Now, "dd/mm/yyyy")
 txtChequeDt.Text = Format(Now, "DD/MM/YYYY")
 AutoIncrementVoucher
 PopulateAcctSuppCust
 PopulateBankCodes
 Flexitems
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
  Sqlqry = "delete * from DUMPPMT1"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
 If Forms.Count >= 1 Then
  MainForm.Picture2.Visible = False
  MainForm.Picture1.Visible = False
 End If
End Sub

Private Sub AutoIncrementVoucher()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select DOC_no from DOCU_MAS WHERE DOC_TYPE='PPT'"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
If rs.RecordCount = 0 Then
   MsgBox "Document type 'PPT' not found"
   Exit Sub
Else
   lblVoucNo = Val(rs!doc_no) + 1
End If
End Sub

Private Sub PopulateAcctSuppCust()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry1 = "Select * from Supp_fin order by Supp_no"
Sqlqry2 = "Select * from AGNDTLS order by agentname"
Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)

lstAcctCode.Clear


If rs1.RecordCount = 0 Then
    MsgBox "No Records found in the Supplier Register"
Else
   rs1.MoveFirst
   Do Until rs1.EOF
      lstAcctCode.AddItem rs1!Supp_no & "    :  " & rs1!Supp_name
      rs1.MoveNext
   Loop
End If

If rs2.RecordCount = 0 Then
    MsgBox "No Records found in the Agency Register"
Else
   rs2.MoveFirst
   Do Until rs2.EOF
      lstAcctCode.AddItem "AGNC" & "    :  " & rs2!Agentname
      rs2.MoveNext
   Loop
End If
    
End Sub

Private Function ValidateData()

ValidateData = False
If txtDate = "" Or IsDate(txtDate) = False Then
  MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
  txtDate.SetFocus
  SendKeys "{Home} + {End}"
  Exit Function
ElseIf txtTtlAmount.Text = "" Or IsNumeric(txtTtlAmount) = False Then
  MsgBox "Invalid Total Amount", vbInformation, "Invalid Entry"
  txtTtlAmount.SetFocus
  Exit Function
ElseIf txtPaidTo.Text = "" Or IsNumeric(txtTtlAmount) = False Then
  MsgBox "Invalid Name of Payee", vbInformation, "Invalid Entry"
  txtPaidTo.SetFocus
  Exit Function
ElseIf lstAcctCode.SelCount = 0 Then
  MsgBox "Select Supplier/Customer Code from list box", vbInformation, "Invalid Entry"
  lstAcctCode.SetFocus
  Exit Function
ElseIf txtDesc.Text = "" Or IsNumeric(txtDesc) = True Then
  MsgBox "Invalid Description", vbInformation, "Invalid Entry"
  txtDesc.SetFocus
  Exit Function
ElseIf txtAmount.Text = "" Or IsNumeric(txtAmount) = False Then
  MsgBox "Invalid Amount", vbInformation, "Invalid Entry"
  txtAmount.SetFocus
  Exit Function
ElseIf txtChequeNO.Text = "" Then
  MsgBox "Invalid Cheque No.", vbInformation, "Invalid Entry"
  txtChequeNO.SetFocus
  Exit Function
ElseIf txtChequeDt.Text = "" Then
  MsgBox "Invalid Cheque Date", vbInformation, "Invalid Entry"
  txtChequeDt.SetFocus
  Exit Function
ElseIf cboBankCode.Text = "" Or IsNumeric(txtAmount) = False Then
  MsgBox "Invalid Amount", vbInformation, "Invalid Entry"
  txtAmount.SetFocus
  Exit Function
Else
  ValidateData = True
End If
End Function

Private Sub Flexitems()
With MSFlexGrid1

    .Clear
    .AllowUserResizing = flexResizeColumns
    .Rows = 1
    .Cols = 4
    .Col = 0
    .Text = "Account Code"
    .ColAlignment(0) = 0
    .ColWidth(0) = 1500
    .ColWidth(1) = 4000
    .ColWidth(2) = 4500
    .ColWidth(3) = 1500
    .Col = 1
    .Text = "Account Name"
    .Col = 2
    .Text = "Description"
    .Col = 3
    .Text = "Amount"
    .Row = 0
    .Col = 1
  
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub lstAcctCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtDesc.SetFocus
End Sub

Private Sub Msflexgrid1_dblclick()
 Dim i
 Dim j
 Dim X
 Dim y, Z, U
 Dim txtaccode, txtacname
 
 X = MSFlexGrid1.Rows
 If X > 1 Then
   i = MsgBox(" Are you sure .. ! You want to Remove this transaction", vbInformation + vbYesNo)
    If i = vbYes Then
     With MSFlexGrid1
        j = .Row
        .Col = 0
        txtaccode = .Text
        .Col = 1
        txtacname = .Text
        .Col = 2
        txtDesc = .Text
        .Col = 3
        txtAmount = .Text
                   
             Sqlqry1 = "select Supp_no,Supp_name from supp_Fin where supp_no='" & Trim(txtaccode) & "' order by supp_no"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                lstAcctCode.Text = rs1!Supp_no & "    :  " & rs1!Supp_name
               Else
                 Sqlqry2 = "Select * from Agndtls where Agentname='" & Trim(txtacname) & "' order by Agentname"
                 Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                 If rs2.RecordCount <> 0 Then
                  lstAcctCode.Text = "C " & rs2!cust_no & "    :  " & rs2!cust_name
                 Else
                  MsgBox "Selected Code not found in Account/Supplier/Customer list"
                 End If
                End If
            End If
                    
          LblTtlAmount.Caption = Val(LblTtlAmount.Caption) - Val(txtAmount)
        
        .RemoveItem (j)
        
        Sqlqry1 = "Delete * from dumppmt1 where Acct_Code='" & txtaccode & "' and description ='" & txtDesc & "' and amount =" & Val(txtAmount) & ""
        ws.BeginTrans
        db.Execute Sqlqry1
        ws.CommitTrans
        
        
     End With
    End If
   Else
    MsgBox " You cannot delete all the Transactions"
   End If
End Sub

Private Sub txtAmount_LostFocus()
  If ValidateData = True Then
    If Val(txtAmount.Text) > Val(txtTtlAmount.Text) Then
      MsgBox " Entered Amount Greater than Total Amount"
      txtAmount.SetFocus
    Exit Sub
     End If
     
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = " select * from DUMPPMT1"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If txtAmount.Text = 0 Then
      Exit Sub
      txtAmount.SetFocus
    End If
    If rs.RecordCount = 0 Then
       Sqlqry = " Insert into DUMPPMT1 values('" & lblVoucNo & "','PPT','" _
                                     & Format(txtDate, "dd/mm/yyyy") & "','" _
                                     & Mid(lstAcctCode, 3, 6) & "','" _
                                     & Mid(lstAcctCode, 14, 35) & "','" _
                                     & UCase(Trim(txtDesc)) & "','" _
                                     & Trim(txtAmount) & "')"

        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        Sqlqry1 = "select * from DUMPPMT1"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlexGrid1.Clear
            Exit Sub
        Else
            Flexitems
            rs.MoveFirst
            Do Until rs.EOF
              MSFlexGrid1.AddItem rs!acct_code & Chr(9) & rs!acct_name & Chr(9) & rs!Description & Chr(9) & rs!amount
              rs.MoveNext
            Loop
        End If
            LblTtlAmount = Val(txtAmount.Text)
            LblTtlAmount.Alignment = 1
            If Val(txtTtlAmount) = Val(txtAmount.Text) Then
            cmdSave.SetFocus
            Else
            lstAcctCode.SetFocus
            End If
      Else
        rs.MoveFirst
        X = 0
         Do Until rs.EOF
          X = X + rs!amount
          rs.MoveNext
         Loop
      
       If Val(txtTtlAmount.Text) >= X + Val(txtAmount.Text) Then
        Sqlqry = " Insert into DUMPPMT1 values('" & lblVoucNo & "','PPT','" _
                                    & Format(txtDate, "dd/mm/yyyy") & "','" _
                                    & Mid(lstAcctCode, 3, 6) & "','" _
                                    & Mid(lstAcctCode, 14, 35) & "','" _
                                    & UCase(Trim(txtDesc)) & "','" _
                                    & Trim(txtAmount) & "')"

          ws.BeginTrans
          db.Execute (Sqlqry)
          ws.CommitTrans
          Sqlqry1 = "Select * from DUMPPMT1"
          Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
           If rs.RecordCount = 0 Then
             MSFlexGrid1.Clear
             Exit Sub
           Else
             Flexitems
             y = 0
             rs.MoveFirst
             Do Until rs.EOF
               MSFlexGrid1.AddItem rs!acct_code & Chr(9) & rs!acct_name & Chr(9) & rs!Description & Chr(9) & rs!amount
               y = y + rs!amount
               rs.MoveNext
             Loop
           End If
             LblTtlAmount = y
             LblTtlAmount.Alignment = 1
             If Val(txtTtlAmount.Text) = y Then
               cmdSave.SetFocus
             Else
               lstAcctCode.SetFocus
             End If
         Else
             MsgBox "Entered Amount is more than Total Amount"
             txtAmount.SetFocus
             Exit Sub
         End If
       End If
 End If
 txtDesc.Text = Trim(txtTtlDesc.Text)
End Sub
Private Sub txtChequeDt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboBankCode.SetFocus
End Sub
Private Sub txtChequeNO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtChequeDt.SetFocus
End Sub
Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtTtlAmount.SetFocus
End Sub
Private Sub txtdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAmount.SetFocus
End Sub
Private Sub txtpaidto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtTtlDesc.SetFocus
End Sub
Private Sub txtTtlAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtChequeNO.SetFocus
End Sub
Private Sub txtTtlAmount_LostFocus()
txtAmount.Text = Val(txtTtlAmount.Text)
End Sub
Private Sub txtTtlDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstAcctCode.SetFocus
End Sub

Private Function textclear()
    
     txtChequeNO.Text = ""
     txtChequeDt.Text = Format(Now, "dd/mm/yyyy")
     cboBankCode.Clear
     txtPaidTo.Text = ""
     txtTtlDesc.Text = ""
     txtDesc.Text = ""
     txtAmount.Text = ""
     txtDate.Text = Format(Now, "dd/mm/yyyy")
     LblTtlAmount.Caption = ""
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from DUMPPMT1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     Flexitems
     txtTtlAmount.SetFocus
     PopulateBankCodes
End Function

Public Sub PopulateBankCodes()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select * from Bank_mas order by bank_code"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

cboBankCode.Clear

 If rs.RecordCount = 0 Then
      MsgBox "No Records found in the Bank Master"
 Else
      rs.MoveFirst
   Do Until rs.EOF
      cboBankCode.AddItem rs!bank_code & " : " & rs!BANK_NAME
      rs.MoveNext
   Loop
 End If

End Sub

Private Sub txtTtlDesc_LostFocus()
txtDesc.Text = Trim(txtTtlDesc.Text)
End Sub
