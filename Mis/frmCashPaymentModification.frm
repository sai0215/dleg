VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmCashPaymentModification 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Cash Payment Modification"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   11850
   FillColor       =   &H00C0C0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cash Payment - Modification"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8295
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   11655
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   3360
         Width           =   4335
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
            TabIndex        =   31
            Top             =   0
            Width           =   1335
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
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   0
            Width           =   1575
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
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.TextBox txtTtlAmount 
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
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   7560
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
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
         Height          =   375
         Left            =   10605
         TabIndex        =   24
         Top             =   1200
         Width           =   780
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
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
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
         Height          =   315
         Left            =   10080
         TabIndex        =   8
         Top             =   3720
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
         Height          =   315
         Left            =   5160
         TabIndex        =   7
         Top             =   3720
         Width           =   4695
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
         Left            =   360
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   3720
         Width           =   4335
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4080
         MaskColor       =   &H0080C0FF&
         Picture         =   "frmCashPaymentModification.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFF80&
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
         Height          =   735
         Left            =   6240
         MaskColor       =   &H0080C0FF&
         Picture         =   "frmCashPaymentModification.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFF80&
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
         Height          =   735
         Left            =   5160
         MaskColor       =   &H0080C0FF&
         Picture         =   "frmCashPaymentModification.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Modify"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3000
         MaskColor       =   &H0080C0FF&
         Picture         =   "frmCashPaymentModification.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7320
         Width           =   1095
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
         Height          =   350
         Left            =   1800
         TabIndex        =   5
         Top             =   2760
         Width           =   9615
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
         Height          =   350
         Left            =   1800
         TabIndex        =   4
         Top             =   2160
         Width           =   4095
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
         Height          =   1500
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7320
         MaskColor       =   &H0080C0FF&
         Picture         =   "frmCashPaymentModification.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7320
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   6720
         Top             =   7560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1575
         Left            =   240
         TabIndex        =   15
         Top             =   5040
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   4
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
         BackColorFixed  =   11258084
         BackColorBkg    =   8421376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PVMaskEditLib.PVMaskEdit txtdate 
         Height          =   375
         Left            =   4320
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
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
      Begin VB.Line Line2 
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         X1              =   11640
         X2              =   0
         Y1              =   7200
         Y2              =   7200
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
         Left            =   9360
         TabIndex        =   27
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Currency "
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
         Left            =   3300
         TabIndex        =   26
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Label5 
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
         Left            =   6165
         TabIndex        =   25
         Top             =   1320
         Width           =   1380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         X1              =   11640
         X2              =   0
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Description"
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
         Left            =   5520
         TabIndex        =   23
         Top             =   3360
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Amount"
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
         Left            =   10320
         TabIndex        =   22
         Top             =   3360
         Width           =   945
      End
      Begin VB.Label Label7 
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
         Left            =   8520
         TabIndex        =   21
         Top             =   6720
         Width           =   1380
      End
      Begin VB.Label lblTtlAmount 
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
         ForeColor       =   &H00404080&
         Height          =   375
         Left            =   10080
         TabIndex        =   20
         Top             =   6600
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Label Label4 
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
         TabIndex        =   18
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   480
         TabIndex        =   17
         Top             =   2880
         Width           =   1200
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
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmCashPaymentModification"
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
Dim LSTAC As Integer
Dim LSTSUP As Integer
Dim LSTAGN As Integer
Dim cur As String
Dim con As Currency
Dim cod As String
Dim j
Dim i

Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtTtlAmount.SetFocus
End Sub

Private Sub cboCurrency_LostFocus()
  If cboCurrency.Text = "USD" Then
     txtConvRate.Text = ""
     lblConvRate.Visible = True
     txtConvRate.Visible = True
   '  lblcurtype.Caption = "USD"
     txtConvRate.TabIndex = 3
     
    Else
     txtConvRate.Text = 1
     lblConvRate.Visible = False
     txtConvRate.Visible = False
   '  lblcurtype.Caption = "DHS"
     txtConvRate.TabIndex = 16
    End If
End Sub

Private Sub cmdAcCode_Click()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Acct_mas order by acct_code"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    
    lstAcctCode.Clear
    
    If rs.RecordCount = 0 Then
        MsgBox "No Records found in the Accounts Master"
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
          lstAcctCode.AddItem rs!agentname
          rs.MoveNext
       Loop
    End If
    
    LSTSUP = 0
    LSTAC = 0
    LSTAGN = 1
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
          lstAcctCode.AddItem rs!Supp_no & "  :  " & rs!Supp_name
          rs.MoveNext
       Loop
    End If
    LSTSUP = 1
    LSTAC = 0
    LSTAGN = 0
End Sub
Private Sub cmdBack_Click()
 Unload Me
End Sub

Private Sub cmdClear_Click()
 textclear
End Sub

Private Sub cmdDelete_Click()
 MsgBox "You are not allowed to delete, Consult Your System Administrator"
 Exit Sub
End Sub

Private Sub Cmdmodify_Click()
 Dim a As Integer
 Dim B As Integer
 Dim C As Integer
 Dim X
 Dim accode, acdesc, acname
 Dim acamt As Currency


cur = ""
cod = ""
con = 1

 If Val(txtTtlAmount.Text) = lblTtlAmount Then
    If cboCurrency.Text = "USD" Then
      cur = "USD"
      cod = "103002"
      con = Val(Trim(txtConvRate.Text))
       
    Else
      cur = "DHS"
      cod = "103001"
      con = 1
    End If
    
 
 If txtTtlAmount <> Val(lblTtlAmount) Then
  MsgBox "Total Amount is Not tallying with Transaction Amount"
  Exit Sub
 End If
 
X = MsgBox("Do You Want to Modify cash Payment Voucher No." & Val(lstVoucNo), vbInformation + vbYesNo, "Confirm")
 
If X = vbNo Then Exit Sub
If ValidateData = True Then
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    
   Sqlqry = "Update cpmt_mas set TDATE=#" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "#," & _
                            " Paid_to ='" & UCase(Trim(txtPaidTo)) & "'," & _
                            " description='" & findfirstfixup(UCase(Trim(txtTtlDesc))) & "'," & _
                            " Tcurrency='" & Trim(cboCurrency) & "'," & _
                            " Tconvertion=" & Val(Trim(txtConvRate)) & "," & _
                            " Tra_amount=" & Val(Trim(txtTtlAmount)) & "," & _
                            " TTl_Amount =" & Val(txtTtlAmount) * Val(Trim(txtConvRate)) & "," & _
                            " Cash_code ='" & Val(cod) & "' Where VOUC_NO = " & Val(lstVoucNo.Text) & " ;"
   ws.BeginTrans
   db.Execute Sqlqry
   ws.CommitTrans
 
   Sqlqry1 = "Delete * from cpmt_tra where vouc_no=" & Val(lstVoucNo) & " "
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
      
       Sqlqry2 = "Insert into cpmt_tra values(" & Val(lstVoucNo.Text) & ",'CPT','" _
                                     & Format(txtdate.TextWithMask, "DD/MM/YYYY") & "','" _
                                     & accode & "','" _
                                     & findfirstfixup(acname) & "','" _
                                     & findfirstfixup(acdesc) & "','" _
                                     & Trim(cboCurrency) & "'," _
                                     & Val(con) & "," _
                                     & acamt & "," _
                                     & Val(acamt) * con & ")"

        ws.BeginTrans
        db.Execute Sqlqry2
        ws.CommitTrans
     Next
   End With
      
  MsgBox " Cash Payment Voucher is modified"
  textclear
  Flexitems
  lstVoucNo.ListIndex = 0
  lstVoucNo.SetFocus
 End If
  
 End If
 

End Sub

Private Sub CmdPrint_Click()
Dim ctype

   If lstVoucNo.SelCount = 0 Then
    MsgBox "Select Voucher Number"
    lstVoucNo.SetFocus
   End If
   ctype = cboCurrency.Text
   If ctype = "DHS" Then
   CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
   CrystalReport1.ReportFileName = App.Path & "\cpmtVou.rpt"
   CrystalReport1.SelectionFormula = "{cpmt_tra.Vouc_no}=" & Val(lstVoucNo.Text) & ""
   CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtTtlAmount)) & " Only" & "'"
   CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
   CrystalReport1.WindowState = crptMaximized
   CrystalReport1.Action = 1
   Else
    CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
    CrystalReport1.ReportFileName = App.Path & "\cpmtVou.rpt"
    CrystalReport1.SelectionFormula = "{cpmt_tra.Vouc_no}=" & Val(lstVoucNo.Text) & ""
    CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(txtTtlAmount)) & " Only" & "'"
    CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
   End If
End Sub

Private Sub Form_Load()
 txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
 
 cboCurrency.AddItem "DHS"
 cboCurrency.AddItem "USD"
 
 lblConvRate.Visible = False
 txtConvRate.Visible = False
 
 LSTSUP = 0
 LSTAC = 1
 LSTAGN = 0
  
 
 PopulateVoucher

 Flexitems
 
End Sub


Private Sub PopulateVoucher()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "select * from cpmt_mas where status='N' ORDER BY VOUC_NO"
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
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Account Code"
    .ColAlignment(0) = 0
    .ColWidth(0) = 1100
    .ColWidth(1) = 3250
    .ColWidth(2) = 6000
    .ColWidth(3) = 900
    .Col = 1
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Account Name"
    .Col = 2
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Description"
    .Col = 3
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Amount"
    .Row = 0
    .Col = 1
 End With
End Sub

Private Sub lstAcctCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdesc.SetFocus
End Sub

Private Sub lstBankCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtPaidTo.SetFocus
End Sub

Private Sub lstVoucNo_Click()
Dim i
Dim X
Dim y
Dim Z
Dim U
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Val(lstVoucNo.Text)
        
        Sqlqry = " Select * from cpmt_mas Where Vouc_no= " & i
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
         If rs.RecordCount <> 0 Then
          
           txtdate.TextWithMask = Format(rs!tDate, "dd/mm/yyyy")
           If rs!tcurrency = "USD" Then
            cboCurrency.ListIndex = 1
           Else
            cboCurrency.ListIndex = 0
           End If
           txtConvRate.Text = rs!tconvertion
           txtTtlAmount = rs!tra_amount
           txtPaidTo = rs!PAID_TO
           txtTtlDesc = rs!Description
           lblTtlAmount.Caption = 0
           
          End If
         
         Sqlqry1 = "Select * from cpmt_tra where Vouc_no= " & i
         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
         If rs1.RecordCount <> 0 Then
           MSFlexGrid1.Clear
           Flexitems
           j = 0
           rs1.MoveFirst
           Do Until rs1.EOF
           MSFlexGrid1.AddItem rs1!acct_code & Chr(9) & rs1!acct_name & Chr(9) & rs1!Description & Chr(9) & rs1!tra_amount
           j = j + rs1!tra_amount
           rs1.MoveNext
           Loop
           lblTtlAmount.Caption = j
           lblTtlAmount.Alignment = 1
           txtdate.SetFocus
         End If
    End Sub

Private Function textclear()
     txtTtlAmount.Text = ""
     lstAcctCode.ListIndex = 0
     txtdesc = ""
     txtAmount = ""
     txtConvRate.Text = ""
     txtPaidTo.Text = ""
     txtTtlDesc.Text = ""
     lblTtlAmount.Caption = ""
     MSFlexGrid1.Clear
     lstVoucNo.ListIndex = 0
     lstVoucNo.SetFocus
     
End Function
Private Sub Msflexgrid1_dblclick()
 Dim i
 Dim j
 Dim X
 Dim y, Z, U
 Dim txtaccode, txtacname
 
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
                MsgBox "No Records found in the Accounts Master"
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
        txtdesc = .Text
        .Col = 3
        txtAmount = .Text
                   
         Sqlqry = "Select acct_Code,acct_name from Acct_mas where acct_code='" & txtaccode & "' order by acct_code"
         Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
                  lstAcctCode.Text = rs!acct_code & "  :  " & rs!acct_name
           Else
             Sqlqry1 = "Select Supp_no,Supp_name from supp_Fin where supp_no='" & Trim(txtaccode) & "' order by supp_no"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                lstAcctCode.Text = rs1!Supp_no & "  :  " & rs1!Supp_name
               Else
                 Sqlqry2 = "Select AGENTNAME from AGNDTLS where AGENTNAME='" & Trim(txtacname) & "' order by aGENTNAME"
                 Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                 If rs2.RecordCount <> 0 Then
                  lstAcctCode.Text = rs2!agentname
                 Else
                  MsgBox "Selected Code not found in Account/Supplier/Agency list"
                 End If
                End If
            End If
                    
          lblTtlAmount.Caption = Val(lblTtlAmount.Caption) - Val(txtAmount)
        
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
   Dim X
   Dim y
   
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
        If LSTAC = 1 Then
         MSFlexGrid1.AddItem Mid(lstAcctCode, 1, 6) & Chr(9) & Mid(lstAcctCode, 12, 25) & Chr(9) & txtdesc & Chr(9) & txtAmount
        ElseIf LSTSUP = 1 Then
         MSFlexGrid1.AddItem Mid(lstAcctCode, 1, 4) & Chr(9) & Mid(lstAcctCode, 10, 25) & Chr(9) & txtdesc & Chr(9) & txtAmount
        Else
          MSFlexGrid1.AddItem "AGNC" & Chr(9) & Trim(lstAcctCode) & Chr(9) & txtdesc & Chr(9) & txtAmount
        End If
    lblTtlAmount = lblTtlAmount + Val(txtAmount)
   End If
  End If
  txtdesc.Text = Trim(txtTtlDesc.Text)
End Sub

Private Sub txtConvRate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then txtPaidTo.SetFocus
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
 If KeyAscii = 13 Then txtAmount.SetFocus
End Sub

Private Sub txtpaidto_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtTtlDesc.SetFocus
End Sub

Private Sub txtTtlAmount_KeyPress(KeyAscii As Integer)
 If cboCurrency.Text = "DHS" Then
    If KeyAscii = 13 Then txtPaidTo.SetFocus
 Else
    If KeyAscii = 13 Then txtConvRate.SetFocus
 End If
 
End Sub

Private Sub txtTtlDesc_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then lstAcctCode.SetFocus
End Sub

Private Sub txtTtlDesc_LostFocus()
txtdesc.Text = Trim(txtTtlDesc.Text)
End Sub
