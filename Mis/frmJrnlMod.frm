VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmJrnlMod 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Journal Modification"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Journal - Modification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7095
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   9375
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Print"
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
         Left            =   3480
         Picture         =   "frmJrnlMod.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6120
         Width           =   975
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
         Left            =   1200
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtDate 
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
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Modify"
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
         Left            =   2520
         Picture         =   "frmJrnlMod.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6120
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
         Left            =   5400
         Picture         =   "frmJrnlMod.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6120
         Width           =   975
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
         Left            =   4440
         Picture         =   "frmJrnlMod.frx":0986
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6120
         Width           =   975
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
         Height          =   1020
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1920
         Width           =   4215
      End
      Begin VB.TextBox txtDesc 
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
         Height          =   315
         Left            =   5040
         TabIndex        =   4
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtAmount 
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
         Height          =   315
         Left            =   8040
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.ListBox lstDCcode 
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
         Height          =   300
         Left            =   4440
         TabIndex        =   2
         Top             =   1920
         Width           =   495
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4440
         Top             =   5160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4260
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
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   9360
         X2              =   0
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Label lblvno 
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
         TabIndex        =   19
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label1 
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
         Left            =   2760
         TabIndex        =   18
         Top             =   600
         Width           =   420
      End
      Begin VB.Label lblTtlAmt 
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
         Left            =   5400
         TabIndex        =   17
         Top             =   5520
         Width           =   1140
      End
      Begin VB.Label LblCAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   8040
         TabIndex        =   16
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label6 
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
         Left            =   8280
         TabIndex        =   15
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label4 
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
         Left            =   5880
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   3105
      End
      Begin VB.Label lblDamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   6840
         TabIndex        =   12
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "D/C"
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
         Left            =   4560
         TabIndex        =   11
         Top             =   1680
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmJrnlMod"
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

Private Sub cmdback_Click()
Unload Me

End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub cmdClear_Click()
     txtDate.Text = Format(Now, "dd/mm/yyyy")
     txtDesc.Text = ""
     LblCAmount.Caption = ""
     lblDamount.Caption = ""
     txtAmount.Text = ""
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Flexitems
     txtDate.SetFocus
End Sub

Private Sub cmdModify_Click()
 Dim a, B
 Dim iCode, iAcName, iDC, iDesc
 Dim iDAmt, iCAmt As Currency

 If ValidateData = True Then
   If Val(lblDamount.Caption) = Val(LblCAmount.Caption) Then
         
   Sqlqry = " Delete * from jrnl_tra where vouc_no=" & Val(lstVoucNo) & ""
   ws.BeginTrans
   db.Execute Sqlqry
   ws.CommitTrans
   
   With MSFlexGrid1
     a = .Rows
    For B = 1 To a - 1
     .Row = B
     .Col = 0
     iCode = .Text
     .Col = 1
     iAcName = .Text
     .Col = 2
     iDC = .Text
     .Col = 3
     iDesc = .Text
     .Col = 4
     iDAmt = .Text
     .Col = 5
     iCAmt = .Text
     
        Sqlqry = "Insert into jrnl_tra values('" & Val(lstVoucNo) & "','JNL','" _
                                & Format(txtDate, "DD/MM/YYYY") & "','" _
                                & Trim(iCode) & "','" _
                                & Trim(iAcName) & "','" _
                                & Trim(iDC) & "','" _
                                & Trim(iDesc) & "','" _
                                & Val(iDAmt) & "','" _
                                & Val(iCAmt) & "','N')"

     
     
          ws.BeginTrans
          db.Execute (Sqlqry)
          ws.CommitTrans
    Next
    End With
  Else
     MsgBox "Total Debit is not equal to Total Credit"
     Exit Sub
  End If
 End If
  MsgBox " Record is Modified", vbInformation, "Status"
  Dim X As Integer
  X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
  If X = vbYes Then
   CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
   CrystalReport1.ReportFileName = App.Path & "\JrnlVou.rpt"
   CrystalReport1.SelectionFormula = "{Jrnl_Tra.Vouc_no}=" & Val(lstVoucNo.Text) & ""
   CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(LblCAmount)) & " Only" & "'"
   CrystalReport1.WindowState = crptMaximized
   CrystalReport1.Action = 1
  End If
  textclear
  lstVoucNo.SetFocus
End Sub

Private Sub cmdPrint_Click()
Dim ttlamount
ttlamount = 0
If lstVoucNo.SelCount = 0 Then
  MsgBox "Select Vouc_no from List Box "
  lstVoucNo.SetFocus
  Exit Sub
 Else
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = " Select * from Jrnl_Tra where Vouc_No=" & Val(Mid(lstVoucNo, 1, 6)) & ""
   Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
   If rs.RecordCount <> 0 Then
    rs.MoveFirst
    Do Until rs.EOF
    ttlamount = ttlamount + rs!damount
    rs.MoveNext
    Loop
   End If
   
   CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
   CrystalReport1.ReportFileName = App.Path & "\JrnlVou.rpt"
   CrystalReport1.SelectionFormula = "{Jrnl_tra.Vouc_no}=" & Val(lstVoucNo) & ""
   CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(ttlamount)) & " Only" & "'"
   CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
   CrystalReport1.WindowState = crptMaximized
   CrystalReport1.Action = 1
     
   End If
End Sub

Private Sub Form_Load()
 txtDate.Text = Format(Now, "dd/mm/yyyy")
 PopulateVoucher
 PopulateAcctSuppCust
 lstDCcode.AddItem "D"
 lstDCcode.AddItem "C"
 lblDamount.Caption = 0
 LblCAmount.Caption = 0
 Flexitems
Unload Me
End Sub

Private Sub PopulateVoucher()
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
  Sqlqry = "Select Distinct(Vouc_No) from JRNL_TRA where status='N' order by vouc_no"
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

Private Sub PopulateAcctSuppCust()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
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

Private Function ValidateData()

ValidateData = False
If txtDate = "" Or IsDate(txtDate) = False Then
  MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
  txtDate.SetFocus
  SendKeys "{Home} + {End}"
  Exit Function
ElseIf lstDCcode.SelCount = 0 Then
  MsgBox "Select Debit or Credit", vbInformation, "Invalid Entry"
  lstDCcode.SetFocus
  Exit Function
ElseIf lstAcctCode.SelCount = 0 Then
  MsgBox "Select Account/Supplier/Customer Code from list box", vbInformation, "Invalid Entry"
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
Else
  ValidateData = True
End If
End Function

Private Sub Flexitems()
With MSFlexGrid1

    .Clear
    .AllowUserResizing = flexResizeColumns
    .Rows = 1
    .Cols = 6
    .Col = 0
    .Text = " Code"
    .ColAlignment(0) = 0
    .ColWidth(0) = 700
    .ColWidth(1) = 2500
    .ColWidth(2) = 500
    .ColWidth(3) = 3200
    .ColWidth(4) = 900
    .ColWidth(5) = 900
    .Col = 1
    .Text = "Account Name"
    .Col = 2
    .Text = "D/C"
    .Col = 3
    .Text = "Description"
    .Col = 4
    .Text = "D_Amount"
    .Col = 5
    .Text = "C_Amount"
    .Row = 0
    .Col = 1
  
  End With
End Sub

Private Sub lstAcctCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then lstDCcode.SetFocus
End Sub

Private Sub lstDCcode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtDesc.SetFocus
End Sub

Private Sub lstVoucNo_Click()
txtDate.SetFocus
End Sub

Private Sub lstVoucNo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtDate.SetFocus
End Sub

Private Sub lstVoucNo_LostFocus()
Dim i

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
      i = Val(lstVoucNo.Text)
        MSFlexGrid1.Clear
        lblDamount.Caption = 0
        LblCAmount.Caption = 0
        
        Sqlqry = " Select * from jrnl_tra Where Vouc_no= " & i
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
         If rs.RecordCount <> 0 Then
           txtDate = Format(rs!tdate, "dd/mm/yyyy")
           rs.MoveFirst
           Do Until rs.EOF
            MSFlexGrid1.AddItem rs!acct_code & Chr(9) & rs!acct_name & Chr(9) & rs!DC_CODE & Chr(9) & rs!Description & Chr(9) & rs!damount & Chr(9) & rs!camount
            lblDamount.Caption = Val(lblDamount) + rs!damount
            LblCAmount.Caption = Val(LblCAmount) + rs!camount
            rs.MoveNext
           Loop
           txtDate.SetFocus
         End If
    End Sub


Private Sub Msflexgrid1_dblclick()
 Dim i As Integer
 Dim j As Integer
 Dim X As Integer
 Dim y, Z, U As Integer
 Dim txtaccode, txtacname
 Dim DAMT, CAMT
 
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
        lstDCcode.Text = .Text
        .Col = 3
        txtDesc = .Text
        .Col = 4
        DAMT = .Text
        .Col = 5
        CAMT = .Text
        
        If DAMT = 0 Then
          txtAmount.Text = Val(CAMT)
          LblCAmount.Caption = Val(LblCAmount.Caption) - Val(txtAmount)
        Else
          txtAmount.Text = Val(DAMT)
          lblDamount.Caption = Val(lblDamount.Caption) - Val(txtAmount)
        End If
                         
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
                    
      .RemoveItem (j)
                       
     End With
    End If
   Else
    MsgBox " You cannot delete all the Transactions"
   End If
 
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstAcctCode.SetFocus
End Sub

Private Sub txtAmount_LostFocus()
 If ValidateData = True Then
    If Val(txtAmount) <= 0 Then
        MsgBox "InValid Amount"
        Exit Sub
        txtAmount.SetFocus
    End If
     
    If lstDCcode.Text = "D" Then
      MSFlexGrid1.AddItem Mid(lstAcctCode, 3, 6) & Chr(9) & Mid(lstAcctCode, 14, 35) & Chr(9) & UCase(lstDCcode) & Chr(9) & Trim(txtDesc) & Chr(9) & Val(txtAmount) & Chr(9) & 0
      lblDamount.Caption = Val(lblDamount) + Val(txtAmount.Text)
      lstAcctCode.SetFocus
    Else
      MSFlexGrid1.AddItem Mid(lstAcctCode, 3, 6) & Chr(9) & Mid(lstAcctCode, 14, 35) & Chr(9) & UCase(lstDCcode) & Chr(9) & Trim(txtDesc) & Chr(9) & 0 & Chr(9) & Val(txtAmount)
      LblCAmount.Caption = Val(LblCAmount) + Val(txtAmount.Text)
      lstAcctCode.SetFocus
    End If
    
 End If
    
End Sub
Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstAcctCode.SetFocus
End Sub


Private Sub txtdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAmount.SetFocus
End Sub


Private Function textclear()
     
     txtDate.Text = Format(Now, "dd/mm/yyyy")
     txtDesc.Text = ""
     txtAmount.Text = ""
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from dumjrnl1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     MSFlexGrid1.Clear
     lstAcctCode.SetFocus
End Function

