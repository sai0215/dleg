VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPdcPaymentModification 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Pdc Payment Modification"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "PDC Payment - Modification "
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
      Height          =   6495
      Left            =   840
      TabIndex        =   16
      Top             =   480
      Width           =   9135
      Begin VB.TextBox txtAmount 
         BackColor       =   &H80000018&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   7680
         TabIndex        =   10
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H80000018&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   4560
         TabIndex        =   9
         Top             =   2760
         Width           =   2895
      End
      Begin VB.ListBox lstAcctCode 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1035
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   2760
         Width           =   4335
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1200
         Picture         =   "frmPdcPaymentModification.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H0080C0FF&
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
         Height          =   855
         Left            =   3360
         Picture         =   "frmPdcPaymentModification.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H0080C0FF&
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
         Height          =   855
         Left            =   2280
         Picture         =   "frmPdcPaymentModification.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0080C0FF&
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
         Height          =   855
         Left            =   120
         Picture         =   "frmPdcPaymentModification.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5520
         Width           =   1095
      End
      Begin VB.TextBox txtTtlDesc 
         BackColor       =   &H80000018&
         ForeColor       =   &H00404040&
         Height          =   350
         Left            =   1200
         TabIndex        =   7
         Top             =   2040
         Width           =   7815
      End
      Begin VB.TextBox txtPaidTo 
         BackColor       =   &H80000018&
         ForeColor       =   &H00404040&
         Height          =   350
         Left            =   6240
         TabIndex        =   6
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtTtlAmount 
         BackColor       =   &H80000018&
         ForeColor       =   &H00404040&
         Height          =   350
         Left            =   6240
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H80000018&
         ForeColor       =   &H00404040&
         Height          =   350
         Left            =   3480
         TabIndex        =   1
         Top             =   360
         Width           =   1335
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
         ForeColor       =   &H00000080&
         Height          =   1020
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtChequeDt 
         BackColor       =   &H80000018&
         ForeColor       =   &H00404040&
         Height          =   350
         Left            =   6240
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtChequeNO 
         BackColor       =   &H80000018&
         ForeColor       =   &H00404040&
         Height          =   350
         Left            =   3720
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.ListBox lstBankCode 
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
         Height          =   300
         Left            =   1200
         TabIndex        =   5
         Top             =   1560
         Width           =   3735
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4440
         Picture         =   "frmPdcPaymentModification.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5520
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   5880
         Top             =   5400
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1215
         Left            =   120
         TabIndex        =   17
         Top             =   3840
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
         BackColorFixed  =   8388608
         BackColorBkg    =   8421376
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Account / Supplier / Customer Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Width           =   3105
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   5520
         TabIndex        =   29
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   8040
         TabIndex        =   28
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6480
         TabIndex        =   27
         Top             =   5160
         Width           =   1140
      End
      Begin VB.Label lblTtlAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   7800
         TabIndex        =   26
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Payee Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   5040
         TabIndex        =   25
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   5040
         TabIndex        =   24
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2640
         TabIndex        =   23
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Voucher No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cheque Dt."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   5040
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cheque No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2640
         TabIndex        =   19
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmPdcPaymentModification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim db As Database
Dim db1 As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim j
Dim i

Private Sub cmdBack_Click()
Unload Me

End Sub
Private Sub Form_Unload(Cancel As Integer)
If Forms.Count >= 2 Then
  MainForm.Picture2.Visible = True
  MainForm.Picture1.Visible = True
End If
End Sub
Private Sub cmdClear_Click()
textclear
End Sub

Private Sub cmdDelete_Click()
MsgBox "You are not allowed to delete Consult System Administrator"
Exit Sub
End Sub

Private Sub cmdModify_Click()
 Dim a As Integer
 Dim B As Integer
 Dim c As Integer
 Dim X
 Dim accode, acdesc, acname
 Dim acamt As Currency
 
 If txtTtlAmount <> Val(lblTtlAmount) Then
  MsgBox "Total Amount is Not tallying with Transaction Amount"
  Exit Sub
 End If
 
X = MsgBox("Do You Want to Modify Pdc Payment Voucher No." & Val(lstVoucNo), vbInformation + vbYesNo, "Confirm")
 
If X = vbNo Then Exit Sub
If ValidateData = True Then
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\UDEUFIN.mdb")
    
   Sqlqry = "Update ppmt_mas set TDATE=#" & Format(txtdate, "dd/mm/yyyy") & "#," & _
                            " Paid_to ='" & UCase(Trim(txtPaidTo)) & "'," & _
                            " Bank_Code ='" & Mid(lstBankCode, 1, 6) & "'," & _
                            " Bank_name ='" & Mid(lstBankCode, 10, 25) & "'," & _
                            " Cheque_No ='" & UCase(Trim(txtChequeNO)) & "'," & _
                            " Cheque_Dt =#" & Format(txtchequedt, "dd/mm/yyyy") & "#," & _
                            " description='" & UCase(Trim(txtTtlDesc)) & "'," & _
                            " TTl_Amount =" & Val(txtTtlAmount) & " Where VOUC_NO = " & Val(lstVoucNo.Text) & " ;"
   ws.BeginTrans
   db.Execute Sqlqry
   ws.CommitTrans
 
   Sqlqry1 = "Delete * from ppmt_tra where vouc_no=" & Val(lstVoucNo) & " "
   ws.BeginTrans
   db.Execute Sqlqry1
   ws.CommitTrans
 
   With MSFlexGrid1
      a = .Rows
     For B = 1 To a - 1
      .Row = B
      .Col = 0
        accode = .Text
      .Col = 1
        acname = .Text
      .Col = 2
        acdesc = .Text
      .Col = 3
        acamt = .Text
      
       Sqlqry2 = "Insert into ppmt_tra values(" & Val(lstVoucNo.Text) & ",'PPT','" _
                                     & Format(txtdate, "DD/MM/YYYY") & "','" _
                                     & accode & "','" _
                                     & acname & "','" _
                                     & acdesc & "',#" _
                                     & Format(txtchequedt, "dd/mm/yyyy") & "#,''," _
                                     & acamt & ")"

        ws.BeginTrans
        db.Execute Sqlqry2
        ws.CommitTrans
     Next
   End With
      
  MsgBox " Pdc Payment Voucher is modified"
  textclear
  Flexitems
  lstVoucNo.SetFocus
 End If

End Sub

Private Sub CmdPrint_Click()
   If lstVoucNo.SelCount = 0 Then
    MsgBox "Select Voucher Number"
    lstVoucNo.SetFocus
   End If
   CrystalReport1.DataFiles(0) = App.Path & "\UDEUFIN.mdb"
   CrystalReport1.ReportFileName = App.Path & "\PpmtVou.rpt"
   CrystalReport1.SelectionFormula = "{ppmt_tra.Vouc_no}=" & Val(lstVoucNo.Text) & ""
   CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtTtlAmount)) & " Only" & "'"
   CrystalReport1.WindowState = crptMaximized
   CrystalReport1.Action = 1
End Sub

Private Sub Form_Load()
 txtdate.Text = Format(Now, "dd/mm/yyyy")
 PopulateVoucher
 PopulateBankCodes
 PopulateAcctSuppCust
 Flexitems
 If Forms.Count >= 1 Then
  MainForm.Picture2.Visible = False
  MainForm.Picture1.Visible = False
 End If
End Sub

Private Sub PopulateAcctSuppCust()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\UDEUFIN.mdb")
Sqlqry = "Select * from Acct_mas order by acct_code"
Sqlqry1 = "Select * from Supp_fin order by Supp_no"
Sqlqry2 = "Select * from Cust_fin order by Cust_no"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)

lstAcctCode.Clear


If rs.RecordCount = 0 Then
    MsgBox "No Records found in the Account Master"
Else
   rs.MoveFirst
   Do Until rs.EOF
      lstAcctCode.AddItem "A " & rs!acct_code & "  :  " & rs!acct_name
      rs.MoveNext
   Loop
End If

If rs1.RecordCount = 0 Then
    MsgBox "No Records found in the Supplier Master"
Else
   rs1.MoveFirst
   Do Until rs1.EOF
      lstAcctCode.AddItem "S " & rs1!Supp_no & "    :  " & rs1!Supp_name
      rs1.MoveNext
   Loop
End If

If rs2.RecordCount = 0 Then
    MsgBox "No Records found in the Customer Master"
Else
   rs2.MoveFirst
   Do Until rs2.EOF
      lstAcctCode.AddItem "C " & rs2!cust_no & "    :  " & rs2!cust_name
      rs2.MoveNext
   Loop
End If
    
End Sub

Private Sub PopulateVoucher()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\UDEUFIN.mdb")
Sqlqry = "select * from ppmt_mas where status='N' ORDER BY VOUC_NO"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
lstVoucNo.Clear
If rs.RecordCount <> 0 Then
    
    rs.MoveFirst
    Do Until rs.EOF
        lstVoucNo.AddItem rs!VOUC_NO
        rs.MoveNext
    Loop
End If
    
End Sub

Private Function ValidateData()

ValidateData = False
If txtdate = "" Or IsDate(txtdate) = False Then
  MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
  txtdate.SetFocus
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
ElseIf txtTtlDesc.Text = "" Or IsNumeric(txtTtlDesc) = True Then
  MsgBox "Invalid Main Description", vbInformation, "Invalid Entry"
  txtTtlDesc.SetFocus
  Exit Function
ElseIf txtChequeNO.Text = "" Then
  MsgBox "Invalid Cheque No.", vbInformation, "Invalid Entry"
  txtChequeNO.SetFocus
  Exit Function
ElseIf txtchequedt.Text = "" Then
  MsgBox "Invalid Cheque Date", vbInformation, "Invalid Entry"
  txtchequedt.SetFocus
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
    .ColWidth(0) = 1100
    .ColWidth(1) = 3000
    .ColWidth(2) = 3500
    .ColWidth(3) = 1100
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

Private Sub lstAcctCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtDesc.SetFocus
End Sub

Private Sub lstBankCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtPaidTo.SetFocus
End Sub

Private Sub lstVoucNo_Click()
Dim i
Dim X
Dim Y
Dim Z
Dim U
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\UDEUFIN.mdb")
    i = Val(lstVoucNo.Text)
        
        Sqlqry = " Select * from ppmt_mas Where Vouc_no= " & i
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
         If rs.RecordCount <> 0 Then
          
           txtdate = Format(rs!tDate, "dd/mm/yyyy")
           txtTtlAmount = rs!ttl_amount
           txtChequeNO = rs!CHEQUE_NO
           txtchequedt = Format(rs!Cheque_Dt, "dd/mm/yyyy")
          End If
          
          Sqlqry1 = "Select BANK_CODE,BANK_NAME from Bank_mas WHERE BANK_CODE='" & Trim(rs!bank_code) & "' ORDER by bank_code"
          Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
          If rs1.RecordCount = 0 Then
             MsgBox "Bank Code Not Found"
          Else
            lstBankCode.Text = rs1!bank_code & " : " & rs1!BANK_NAME
          End If
                                        
          
           txtPaidTo = rs!Paid_To
           txtTtlDesc = rs!Description
           
           lblTtlAmount.Caption = 0
         
         Sqlqry1 = "Select * from ppmt_tra where Vouc_no= " & i
         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
         If rs1.RecordCount <> 0 Then
           MSFlexGrid1.Clear
           Flexitems
           j = 0
           rs1.MoveFirst
           Do Until rs1.EOF
           MSFlexGrid1.AddItem rs1!acct_code & Chr(9) & rs1!acct_name & Chr(9) & rs1!Description & Chr(9) & rs1!amount
           j = j + rs1!amount
           rs1.MoveNext
           Loop
           lblTtlAmount.Caption = j
           lblTtlAmount.Alignment = 1
           txtdate.SetFocus
         End If
    End Sub

Private Function textclear()
     txtTtlAmount.Text = ""
     txtChequeNO.Text = ""
     txtchequedt.Text = ""
     lstBankCode.ListIndex = 0
     lstAcctCode.ListIndex = 0
     txtDesc = ""
     txtAmount = ""
     txtPaidTo.Text = ""
     txtTtlDesc.Text = ""
     lblTtlAmount.Caption = ""
     MSFlexGrid1.Clear
     lstVoucNo.ListIndex = 0
     lstVoucNo.SetFocus
     
End Function

Public Sub PopulateBankCodes()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\UDEUFIN.mdb")
Sqlqry = "Select * from Bank_mas order by bank_code"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

lstBankCode.Clear

 If rs.RecordCount = 0 Then
      MsgBox "No Records found in the Bank Master"
 Else
      rs.MoveFirst
   Do Until rs.EOF
      lstBankCode.AddItem rs!bank_code & " : " & rs!BANK_NAME
      rs.MoveNext
   Loop
   i = lstBankCode.ListIndex
   If i = -1 Then
   lstBankCode.Enabled = True
   End If
   
 End If

End Sub

Private Sub Msflexgrid1_dblclick()
 Dim i
 Dim j
 Dim X
 Dim Y, Z, U
 Dim txtaccode, txtacname
 
 X = MSFlexGrid1.Rows
 If MSFlexGrid1.Row = 1 Then
  Exit Sub
 End If
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
                   
         Sqlqry = "Select acct_Code,acct_name from Acct_mas where acct_code='" & txtaccode & "' order by acct_code"
         Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
                  lstAcctCode.Text = "A " & rs!acct_code & "  :  " & rs!acct_name
           Else
             Sqlqry1 = "Select Supp_no,Supp_name from supp_Fin where supp_no='" & Trim(txtaccode) & "' order by supp_no"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                lstAcctCode.Text = "S " & rs1!Supp_no & "    :  " & rs1!Supp_name
               Else
                 Sqlqry2 = "Select cust_no,cust_name from Cust_fin where cust_no='" & Trim(txtaccode) & "' order by Cust_no"
                 Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                 If rs2.RecordCount <> 0 Then
                  lstAcctCode.Text = "C " & rs2!cust_no & "    :  " & rs2!cust_name
                 Else
                  MsgBox "Selected Code not found in Account/Supplier/Customer list"
                 End If
                End If
            End If
                    
          lblTtlAmount.Caption = Val(lblTtlAmount.Caption) - Val(txtAmount)
        
        .RemoveItem (j)
                       
     End With
    End If
   Else
    MsgBox " You cannot delete all the Transactions"
   End If
End Sub

Private Sub txtAmount_LostFocus()
   Dim X
   Dim Y
   
   If ValidateData = True Then
    If Val(txtAmount.Text) > Val(txtTtlAmount.Text) Then
      MsgBox " Entered Amount Greater than Total Amount"
      txtAmount.SetFocus
    Exit Sub
    End If
      
   X = lblTtlAmount + Val(txtAmount)
   If X > Val(txtTtlAmount) Then
    MsgBox "Entered Amount Greater Than Total Amount"
    Exit Sub
   Else
    MSFlexGrid1.AddItem Mid(lstAcctCode, 3, 6) & Chr(9) & Mid(lstAcctCode, 12, 25) & Chr(9) & txtDesc & Chr(9) & txtAmount
    lblTtlAmount = lblTtlAmount + Val(txtAmount)
   End If
  End If
  txtDesc.Text = Trim(txtTtlDesc.Text)
End Sub

Private Sub txtChequeDt_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then lstBankCode.SetFocus
End Sub

Private Sub txtChequeNO_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtchequedt.SetFocus
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

Private Sub txtTtlDesc_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then lstAcctCode.SetFocus
End Sub

Private Sub txtTtlDesc_LostFocus()
txtDesc.Text = Trim(txtTtlDesc.Text)
End Sub
