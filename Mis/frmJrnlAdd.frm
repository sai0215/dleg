VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmJrnlAdd 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Journal Addition"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
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
   ScaleHeight     =   8175
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Journal  - New Entry"
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
      Height          =   7575
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   11295
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
         Height          =   345
         Left            =   9600
         TabIndex        =   24
         Top             =   570
         Width           =   1140
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
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   3855
         Begin VB.CommandButton cmdAcCode 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Gen. A/C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmdSupplier 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Supplier"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdAgency 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Agency"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFF80&
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
         Left            =   3480
         Picture         =   "frmJrnlAdd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6600
         Width           =   975
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFF80&
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
         Picture         =   "frmJrnlAdd.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6600
         Width           =   975
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFF80&
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
         Picture         =   "frmJrnlAdd.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6600
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
         Height          =   1020
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1800
         Width           =   4095
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
         Height          =   315
         Left            =   5520
         TabIndex        =   4
         Top             =   1800
         Width           =   3975
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
         Height          =   315
         Left            =   9720
         TabIndex        =   5
         Top             =   1800
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
         Height          =   300
         Left            =   4320
         TabIndex        =   3
         Top             =   1800
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   3360
         Top             =   5160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2775
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   5
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
         BackColorFixed  =   9613530
         BackColorBkg    =   8421376
      End
      Begin PVMaskEditLib.PVMaskEdit txtdate 
         Height          =   375
         Left            =   3720
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
      Begin VB.Label lblcurtype 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   8040
         TabIndex        =   27
         Top             =   6000
         Width           =   555
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
         Left            =   8385
         TabIndex        =   26
         Top             =   720
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
         Left            =   5760
         TabIndex        =   25
         Top             =   720
         Width           =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         BorderWidth     =   2
         X1              =   11280
         X2              =   0
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Label lblvno 
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
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label lblVoucNo 
         Alignment       =   2  'Center
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
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   1560
         TabIndex        =   18
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label Label1 
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
         Left            =   3120
         TabIndex        =   17
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblTtlAmt 
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
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   6360
         TabIndex        =   16
         Top             =   6000
         Width           =   1380
      End
      Begin VB.Label LblCAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   9840
         TabIndex        =   15
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   9960
         TabIndex        =   14
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   5880
         TabIndex        =   13
         Top             =   1560
         Width           =   1200
      End
      Begin VB.Label lblDamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8640
         TabIndex        =   12
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Db / Cr"
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
         Left            =   4560
         TabIndex        =   11
         Top             =   1560
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmJrnlAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim rs1 As Recordset
Dim rs As Recordset
Dim rs2 As Recordset
Dim Sqlqry1 As String
Dim Sqlqry As String
Dim Sqlqry2 As String
Dim Sqlqry3 As String
Dim X
Dim y
Dim Z
Dim i
Dim j

Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
    If cboCurrency.Text = "USD" Then
       lblConvRate.Visible = True
       txtConvRate.Visible = True
      If KeyAscii = 13 Then txtConvRate.SetFocus
    Else
       If KeyAscii = 13 Then lstAcctCode.SetFocus
    End If

End Sub

Private Sub cboCurrency_LostFocus()
  If cboCurrency.Text = "USD" Then
     lblConvRate.Visible = True
     txtConvRate.Visible = True
     lblcurtype.Caption = "USD"
     txtConvRate.Text = ""
     txtConvRate.TabIndex = 2
     
    Else
     lblConvRate.Visible = False
     txtConvRate.Visible = False
     lblcurtype.Caption = "DHS"
     txtConvRate.Text = 1
     txtConvRate.TabIndex = 10
    End If
End Sub

Private Sub cmdAcCode_Click()
PopulateAcctCodes
End Sub
Private Sub PopulateAcctCodes()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select * from Acct_mas order by acct_code"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

lstAcctCode.Clear

 If rs.RecordCount = 0 Then
      MsgBox "No Records found in the Account Register"
 Else
      rs.MoveFirst
      Do Until rs.EOF
       lstAcctCode.AddItem rs!acct_code & " : " & rs!acct_name
      rs.MoveNext
   Loop
 End If
 
   LSTSUP = 0
   LSTAC = 1
   LSTAGN = 0
 
End Sub
Private Sub cmdAgency_Click()

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Agndtls order by agentname"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    
    lstAcctCode.Clear
    
    If rs.RecordCount = 0 Then
        MsgBox "No Records found in the Agency Register"
    Else
       rs.MoveFirst
       Do Until rs.EOF
          lstAcctCode.AddItem "  AGNC" & "  :  " & rs!agentname
          rs.MoveNext
       Loop
    End If
    
    LSTSUP = 0
    LSTAC = 0
    LSTAGN = 1
End Sub
Private Sub cmdBack_Click()
 Unload Me
End Sub

Private Sub cmdSupplier_Click()
 Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Supp_fin order by Supp_no"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    
    lstAcctCode.Clear
    
    If rs.RecordCount = 0 Then
        MsgBox "No Records found in the Suppliers Register"
    Else
       rs.MoveFirst
       Do Until rs.EOF
          lstAcctCode.AddItem "  " & rs!Supp_no & "  :  " & rs!Supp_name
          rs.MoveNext
       Loop
    End If
    LSTSUP = 1
    LSTAC = 0
    LSTAGN = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub
Private Sub cmdClear_Click()
     txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
     txtdesc.Text = ""
     LblCAmount.Caption = ""
     lblDamount.Caption = ""
     txtAmount.Text = ""
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from dumjrnl1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     MSFlexGrid1.Clear
     txtdate.SetFocus
     cboCurrency.ListIndex = -1
     txtConvRate.Text = 1
     txtConvRate.Visible = False
     lblConvRate.Visible = False
     
End Sub
Private Sub CmdSave_Click()
Dim ctype As String
 If ValidateData = True Then
   If Val(lblDamount.Caption) = Val(LblCAmount.Caption) Then
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         Sqlqry1 = "Select * from dumjrnl1"
         Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
          If rs.RecordCount = 0 Then
           MsgBox " Transactions are not recorded"
          Exit Sub
   Else
         rs.MoveFirst
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry3 = "Insert into jrnl_tra values('" & rs!vouc_no & "','" & rs!vouc_type & "','" _
                                & Trim(rs!tDate) & "','" _
                                & rs!acct_code & "','" _
                                & findfirstfixup(rs!acct_name) & "','" _
                                & rs!DC_CODE & "','" _
                                & findfirstfixup(rs!Description) & "','" _
                                & Trim(cboCurrency.Text) & "'," _
                                & Val(txtConvRate.Text) & "," _
                                & rs!damount & "," _
                                & rs!camount & "," _
                                & rs!damount * Val(txtConvRate.Text) & "," _
                                & rs!camount * Val(txtConvRate.Text) & ",'N')"

          ws.BeginTrans
          db.Execute (Sqlqry3)
          ws.CommitTrans
          rs.MoveNext
          Loop
     End If
     lblDamount.Caption = ""
     LblCAmount.Caption = ""
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Update docu_mas set doc_no='" & lblVoucNo & "' where doc_type='JNL'"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     
     lblVoucNo = lblVoucNo + 1
     
     
   Else
     MsgBox "Total Debit is not equal to Total Credit"
     Exit Sub
   End If
    MsgBox " Record is inserted", vbInformation, "Status"
  Dim X As Integer
  ctype = cboCurrency.Text
  X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
  If X = vbYes Then
   If ctype = "DHS" Then
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\JrnlVou.rpt"
        CrystalReport1.SelectionFormula = "{Jrnl_Tra.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
        CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(LblCAmount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    Else
                CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\JrnlVou.rpt"
        CrystalReport1.SelectionFormula = "{Jrnl_Tra.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
        CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(LblCAmount)) & " Only" & "'"
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
 AutoIncrementVoucher
 PopulateAcctSuppCust
 lstDCcode.AddItem "Debit"
 lstDCcode.AddItem "Credit"
 lblDamount.Caption = 0
 LblCAmount.Caption = 0
 
 cboCurrency.AddItem "DHS"
 cboCurrency.AddItem "USD"
 lblConvRate.Visible = False
 txtConvRate.Visible = False
 txtConvRate.Text = 1
 
 LSTSUP = 0
 LSTAC = 1
 LSTAGN = 0
 
 Flexitems
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
  Sqlqry = "delete * from dumjrnl1"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
End Sub

Private Sub AutoIncrementVoucher()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select DOC_no from DOCU_MAS WHERE DOC_TYPE='JNL'"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
If rs.RecordCount = 0 Then
   MsgBox "Document type 'JNL' not found"
   Exit Sub
Else
   lblVoucNo = Val(rs!doc_no) + 1
End If
End Sub

Private Sub PopulateAcctSuppCust()

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select * from Acct_mas order by acct_code"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

lstAcctCode.Clear

If rs.RecordCount = 0 Then
    MsgBox "No Records found in the Account Register"
Else
   rs.MoveFirst
   Do Until rs.EOF
      lstAcctCode.AddItem rs!acct_code & "  :  " & rs!acct_name
      rs.MoveNext
   Loop
End If

 LSTSUP = 0
 LSTAC = 1
 LSTAGN = 0
    
End Sub

Private Function ValidateData()

ValidateData = False
If IsDate(txtdate.TextWithMask) = False Then
  MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
  txtdate.SetFocus
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
ElseIf txtdesc.Text = "" Or IsNumeric(txtdesc) = True Then
  MsgBox "Invalid Description", vbInformation, "Invalid Entry"
  txtdesc.SetFocus
  Exit Function
ElseIf txtConvRate.Text = "" Then
  MsgBox "Enter Convertion Rate - - cannot be zero", vbInformation, "Invalid Entry"
  txtConvRate.SetFocus
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
    .CellBackColor = RGB(180, 170, 160)
    .Text = " Code"
    .ColAlignment(0) = 0
    .ColWidth(0) = 700
    .ColWidth(1) = 2500
    .ColWidth(2) = 500
    .ColWidth(3) = 5300
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .Col = 1
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Account Name"
    .Col = 2
    .CellBackColor = RGB(180, 170, 160)
    .Text = "D/C"
    .Col = 3
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Description"
    .Col = 4
    .CellBackColor = RGB(180, 170, 160)
    .Text = "D_Amount"
    .Col = 5
    .CellBackColor = RGB(180, 170, 160)
    .Text = "C_Amount"
    .Row = 0
    .Col = 1
  
  End With
End Sub

Private Sub lstAcctCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then lstDCcode.SetFocus
End Sub
Private Sub lstDCcode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdesc.SetFocus
End Sub

Private Sub MSFlexGrid1_Click()
 Dim i
 Dim j
 Dim X
 Dim y, Z, U
 Dim txtaccode, txtacname
 Dim DAMT, CAMT
 
 X = MSFlexGrid1.Rows
If X > 1 Then
 If MSFlexGrid1.Row = MSFlexGrid1.TopRow Then
  Exit Sub
 Else
   i = MsgBox(" Are you sure .. ! You want to Remove this transaction", vbInformation + vbYesNo)
    If i = vbYes Then
    
     With MSFlexGrid1
        j = .Row
        .Col = 0
        txtaccode = .Text
         If Len(.Text) = 6 Then
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = "Select * from Acct_mas order by acct_code"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            
            lstAcctCode.Clear
            
            If rs.RecordCount = 0 Then
                MsgBox "No Records found in the Accounts Register"
            Else
               rs.MoveFirst
               Do Until rs.EOF
                  lstAcctCode.AddItem rs!acct_code & "  :  " & rs!acct_name
                  rs.MoveNext
               Loop
            End If
            LSTSUP = 0
            LSTAC = 1
            LSTAGN = 0
            
          ElseIf .Text = "AGNC" Then
     
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = "Select * from Agndtls order by agentname"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            
            lstAcctCode.Clear
            
            If rs.RecordCount = 0 Then
                MsgBox "No Records found in the Agency Register"
            Else
               rs.MoveFirst
               Do Until rs.EOF
                  lstAcctCode.AddItem rs!agentname
                  rs.MoveNext
               Loop
            End If
            
            LSTSUP = 0
            LSTAC = 0
            LSTAGN = 1

        Else
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = "Select * from Supp_fin order by Supp_no"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            
            lstAcctCode.Clear
            
            If rs.RecordCount = 0 Then
                MsgBox "No Records found in the Suppliers Register"
            Else
               rs.MoveFirst
               Do Until rs.EOF
                  lstAcctCode.AddItem rs!Supp_no & "  :  " & rs!Supp_name
                  rs.MoveNext
               Loop
            End If
            LSTSUP = 1
            LSTAC = 0
            LSTAGN = 0
       End If
       
        .Col = 1
        txtacname = .Text
        .Col = 2
        lstDCcode.Text = .Text
        .Col = 3
        txtdesc = .Text
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
                  lstAcctCode.Text = rs!acct_code & "  :  " & rs!acct_name
           Else
             Sqlqry1 = "Select Supp_no,Supp_name from supp_Fin where supp_no='" & Trim(txtaccode) & "' order by supp_no"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                lstAcctCode.Text = "  " & rs1!Supp_no & "  :  " & rs1!Supp_name
               Else
                 Sqlqry2 = "Select * from agndtls where agentname='" & Trim(txtacname) & "' order by agentname"
                 Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                 If rs2.RecordCount <> 0 Then
                  lstAcctCode.Text = "  AGNC" & "  :  " & rs2!agentname
                 Else
                  MsgBox "Selected Code not found in Account/Supplier/Customer Register"
                 End If
                End If
            End If
                    
      .RemoveItem (j)
                       
     End With
    End If
   End If
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
     
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = " select * from dumjrnl1"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
       If rs.RecordCount = 0 Then
        
        If lstDCcode.Text = "Debit" Then
           Sqlqry = " Insert into dumjrnl1 values('" & lblVoucNo & "','JNL','" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & Trim(Mid(lstAcctCode, 1, 6)) & "','" _
                                     & Trim(Mid(lstAcctCode, 12, 35)) & "','" _
                                     & UCase(lstDCcode) & "','" _
                                     & findfirstfixup(Trim(txtdesc)) & "'," _
                                     & Val(Trim(txtAmount)) & ", " & 0 & ")"

            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
            
            Sqlqry1 = "select * from dumjrnl1"
            Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
            If rs.RecordCount > 0 Then
              
              Flexitems
              rs.MoveFirst
               Do Until rs.EOF
                MSFlexGrid1.AddItem rs!acct_code & Chr(9) & rs!acct_name & Chr(9) & rs!DC_CODE & Chr(9) & rs!Description & Chr(9) & rs!damount & Chr(9) & rs!camount
                rs.MoveNext
               Loop
            End If
            lblDamount.Caption = Val(txtAmount.Text)
            lstAcctCode.SetFocus
        Else
            Sqlqry = " Insert into dumjrnl1 values('" & lblVoucNo & "','JNL','" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & Trim(Mid(lstAcctCode, 1, 6)) & "','" _
                                     & Trim(Mid(lstAcctCode, 12, 35)) & "','" _
                                     & UCase(lstDCcode) & "','" _
                                     & findfirstfixup(Trim(txtdesc)) & "'," _
                                     & 0 & ", " & Val(Trim(txtAmount)) & ")"

            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
            
            Sqlqry1 = "select * from dumjrnl1"
            Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
            If rs.RecordCount = 0 Then
              MSFlexGrid1.Clear
              Exit Sub
            Else
              Flexitems
              rs.MoveFirst
               Do Until rs.EOF
                MSFlexGrid1.AddItem rs!acct_code & Chr(9) & rs!acct_name & Chr(9) & rs!DC_CODE & Chr(9) & rs!Description & Chr(9) & rs!damount & Chr(9) & rs!camount
                rs.MoveNext
               Loop
            End If
            LblCAmount.Caption = Val(txtAmount.Text)
            lstAcctCode.SetFocus
          End If
       
     Else
       X = 0
       y = 0
       rs.MoveFirst
       Do Until rs.EOF
        If rs!DC_CODE = "D" Then
          X = X + rs!damount
        End If
        If rs!DC_CODE = "C" Then
          y = y + rs!camount
        End If
        rs.MoveNext
       Loop
       
       If lstDCcode = "Debit" Then
        lblDamount.Caption = Val(txtAmount) + X
       Else
        LblCAmount.Caption = Val(txtAmount) + y
       End If
       
       If lstDCcode.Text = "Debit" Then
          Sqlqry = " Insert into dumjrnl1 values('" & lblVoucNo & "','JNL','" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & Trim(Mid(lstAcctCode, 1, 6)) & "','" _
                                     & Trim(Mid(lstAcctCode, 12, 35)) & "','" _
                                     & UCase(lstDCcode) & "','" _
                                     & findfirstfixup(Trim(txtdesc)) & "'," _
                                     & Val(Trim(txtAmount)) & ", " & 0 & ")"

            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
            
            Sqlqry1 = "select * from dumjrnl1"
            Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
            If rs.RecordCount = 0 Then
              MSFlexGrid1.Clear
              Exit Sub
            Else
              Flexitems
              rs.MoveFirst
               Do Until rs.EOF
                MSFlexGrid1.AddItem rs!acct_code & Chr(9) & rs!acct_name & Chr(9) & rs!DC_CODE & Chr(9) & rs!Description & Chr(9) & rs!damount & Chr(9) & rs!camount
                rs.MoveNext
               Loop
            End If
            lstAcctCode.SetFocus
        Else
            Sqlqry = " Insert into dumjrnl1 values('" & lblVoucNo & "','JNL','" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & Trim(Mid(lstAcctCode, 1, 6)) & "','" _
                                     & Trim(Mid(lstAcctCode, 12, 35)) & "','" _
                                     & UCase(lstDCcode) & "','" _
                                     & findfirstfixup(Trim(txtdesc)) & "'," _
                                     & 0 & ", " & Val(Trim(txtAmount)) & ")"

            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
            
            Sqlqry1 = "select * from dumjrnl1"
            Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
            If rs.RecordCount = 0 Then
              MSFlexGrid1.Clear
              Exit Sub
            Else
              Flexitems
              rs.MoveFirst
               Do Until rs.EOF
                MSFlexGrid1.AddItem rs!acct_code & Chr(9) & rs!acct_name & Chr(9) & rs!DC_CODE & Chr(9) & rs!Description & Chr(9) & rs!damount & Chr(9) & rs!camount
                rs.MoveNext
               Loop
            End If
            lstAcctCode.SetFocus
         End If
       End If
     End If
 
End Sub

Private Sub txtConvRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstAcctCode.SetFocus
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
If KeyAscii = 13 Then txtAmount.SetFocus
End Sub

Private Function textclear()
     
     txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
     txtdesc.Text = ""
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

