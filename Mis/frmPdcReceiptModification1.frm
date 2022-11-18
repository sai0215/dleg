VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "PVMASK.OCX"
Begin VB.Form frmPdcReceiptModification1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pdc Receipt Modification"
   ClientHeight    =   8685
   ClientLeft      =   30
   ClientTop       =   345
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PDC Receipt Modification"
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
      TabIndex        =   17
      Top             =   120
      Width           =   11535
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
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
         Height          =   735
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   7440
         Width           =   975
      End
      Begin VB.TextBox txtpendingamount 
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   5520
         TabIndex        =   4
         Top             =   1560
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
         Height          =   1500
         Left            =   1560
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7440
         Width           =   975
      End
      Begin VB.TextBox txtChequeNO 
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
         Left            =   360
         TabIndex        =   8
         Top             =   4320
         Width           =   855
      End
      Begin VB.ComboBox lstBankCode 
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
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4320
         Width           =   3135
      End
      Begin VB.TextBox txtRecdFrom 
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
         Left            =   1560
         TabIndex        =   5
         Top             =   2280
         Width           =   4215
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
         Left            =   9600
         TabIndex        =   12
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFFFF&
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
         Height          =   735
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7440
         Width           =   975
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFFFF&
         Caption         =   "<<&Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7440
         Width           =   975
      End
      Begin VB.CommandButton cmdmodify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7440
         Width           =   975
      End
      Begin VB.TextBox txtTtlDesc 
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
         Left            =   1560
         MaxLength       =   150
         TabIndex        =   7
         Top             =   3480
         Width           =   7575
      End
      Begin VB.ComboBox lstAcctCode 
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
         TabIndex        =   6
         Top             =   2880
         Width           =   4215
      End
      Begin VB.TextBox txtChequeBank 
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
         Left            =   3000
         TabIndex        =   10
         Top             =   4320
         Width           =   2895
      End
      Begin VB.TextBox txtConvRate 
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   8640
         TabIndex        =   18
         Top             =   960
         Width           =   1215
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
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtTtlAmount 
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   5520
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   8400
         Top             =   1920
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2055
         Left            =   240
         TabIndex        =   19
         Top             =   4800
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   5
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
         BackColorFixed  =   12111599
         BackColorBkg    =   8421376
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PVMaskEditLib.PVMaskEdit txtdate 
         Height          =   375
         Left            =   5520
         TabIndex        =   1
         Top             =   360
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
      Begin PVMaskEditLib.PVMaskEdit txtchequedt 
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   4320
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
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pending Amount"
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
         Left            =   3480
         TabIndex        =   35
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Chq Date"
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
         Left            =   1560
         TabIndex        =   34
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Chq #"
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
         TabIndex        =   33
         Top             =   4080
         Width           =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Depositing Bank"
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
         TabIndex        =   32
         Top             =   4080
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   31
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Payer Name"
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
         TabIndex        =   30
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Left            =   3600
         TabIndex        =   29
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   240
         TabIndex        =   28
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   27
         Top             =   4080
         Width           =   780
      End
      Begin VB.Label LblTtlAmount 
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   9960
         TabIndex        =   26
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Label lblTtlAmt 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total "
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
         Left            =   9120
         TabIndex        =   25
         Top             =   6960
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   240
         TabIndex        =   24
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   11520
         X2              =   0
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cheque Bank"
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
         TabIndex        =   23
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   11520
         X2              =   0
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Label lblConvRate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Left            =   7320
         TabIndex        =   22
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Left            =   7440
         TabIndex        =   21
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Left            =   3480
         TabIndex        =   20
         Top             =   1080
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmPdcReceiptModification1"
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
Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtTtlAmount.SetFocus
End Sub

Private Sub cboCurrency_LostFocus()
If CboCurrency.Text = "DHS" Then
   txtConvRate.Text = 1
   txtConvRate.Visible = False
   lblConvRate.Visible = False
   txtConvRate.TabIndex = 20
Else
  txtConvRate.Visible = True
  lblConvRate.Visible = True
  txtConvRate.Text = ""
  txtConvRate.TabIndex = 4
End If
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

Private Sub cmdDelete_Click()
 MsgBox "You are not allowed to delete, Consult System Administrator"
 Exit Sub
End Sub

Private Sub Cmdmodify_Click()
 Dim a
 Dim B
 Dim C
 Dim X
 Dim MchqNo, MchqBank, MDepBank, MBankCode
 Dim MchqDt As Date
 Dim mamount As Currency
 
 If txtpendingamount <> Val(lblTtlAmount) Then
  MsgBox "Total amount is not tallying with transactions total"
  Exit Sub
 End If
 
 X = MsgBox("Do You Want to Modify Pdc Receipt Voucher No." & Val(lstVoucNo), vbInformation + vbYesNo, "Confirm")
 
If X = vbNo Then Exit Sub
 If ValidateData = True Then
  If Val(txtpendingamount.Text) = Val(lblTtlAmount.Caption) Then
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    
   
   Sqlqry1 = "Delete * from Prpt_mas1 where vouc_no=" & Val(lstVoucNo) & " and status='N'"
   ws.BeginTrans
   db.Execute Sqlqry1
   ws.CommitTrans
 
   With MSFlexGrid1
      a = .Rows
     For B = 1 To a - 1
      .Row = B
      .Col = 0
        MchqNo = .Text
      .Col = 1
        MchqDt = .Text
      .Col = 2
        MchqBank = .Text
      .Col = 3
        MBankCode = .Text
      .Col = 4
        MDepBank = .Text
      .Col = 5
         mamount = .Text
         
         Sqlqry2 = "Insert into Prpt_Mas1 values('" & Val(lstVoucNo.Text) & "','PRT',#" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "#,'" _
                                     & Trim(txtRecdFrom) & "','" _
                                     & findfirstfixup(Trim(txtTtlDesc)) & "','" _
                                     & Trim(CboCurrency) & "'," _
                                     & Trim(Val(txtConvRate)) & "," _
                                     & Val(Trim(mamount)) & ",'" _
                                     & Mid(lstAcctCode, 1, 4) & "','" _
                                     & findfirstfixup(Mid(lstAcctCode, 12, 35)) & "','" _
                                     & Trim(MBankCode) & "','" _
                                     & Trim(MDepBank) & "','" _
                                     & MchqNo & "','" _
                                     & Trim(MchqBank) & "','" _
                                     & Trim(MchqDt) & "',''," _
                                     & Val(Trim(mamount)) * Val(txtConvRate) & ",'N')"

            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
     Next
   End With
      
  MsgBox " Pdc Receipt Voucher is modified"
  textclear
  Flexitems
  lstVoucNo.ListIndex = 0
 Else
  MsgBox " Total amount is not tallying with transactions total"
  Exit Sub
 End If
 End If

End Sub

Private Sub CmdPrint_Click()
Dim ctype As String
   If lstVoucNo.SelCount = 0 Then
    MsgBox "Select Voucher Number"
    lstVoucNo.SetFocus
   End If
   ctype = CboCurrency.Text
   If ctype = "DHS" Then
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\PrptVou1.rpt"
        CrystalReport1.SelectionFormula = "{Prpt_Mas1.Vouc_no}=" & Val(lstVoucNo.Text) & ""
        CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtpendingamount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & CboCurrency.Text & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
   Else
       CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\PrptVou1.rpt"
        CrystalReport1.SelectionFormula = "{Prpt_Mas1.Vouc_no}=" & Val(lstVoucNo.Text) & ""
        CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(txtpendingamount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & CboCurrency.Text & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
   End If
End Sub

Private Sub Form_Load()
 txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
 txtchequedt.TextWithMask = Format(Now, "dd/mm/yyyy")
 CboCurrency.AddItem "DHS"
 CboCurrency.AddItem "USD"
 lblConvRate.Visible = False
 txtConvRate.Visible = False
 txtConvRate.Text = 1
  
 PopulateVoucher
 PopulateBankCodes
 PopulateAcctSuppCust
 Flexitems
 
End Sub

Private Sub PopulateAcctSuppCust()
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry1 = "Select * from Supp_fin order by Supp_no"
 Sqlqry2 = "Select * from Agndtls order by agentname"

 Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
 Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)

 lstAcctCode.Clear

 If rs2.RecordCount = 0 Then
    MsgBox "No Records found in the Agency Register"
 Else
    rs2.MoveFirst
   Do Until rs2.EOF
      lstAcctCode.AddItem "AGNC" & "    :  " & rs2!agentname
      rs2.MoveNext
   Loop
 End If

  If rs1.RecordCount = 0 Then
    MsgBox "No Records found in the Supplier Register"
  Else
   rs1.MoveFirst
   Do Until rs1.EOF
      lstAcctCode.AddItem rs1!Supp_no & "    :  " & rs1!SUPP_NAME
      rs1.MoveNext
   Loop
  End If
   
End Sub
 
Private Sub PopulateVoucher()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "select distinct(Vouc_no) from Prpt_mas1 where status='N' ORDER BY VOUC_NO"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
lstVoucNo.Clear
 If rs.RecordCount = 0 Then
     MsgBox "No records found in the PDC Receipt Register"
 Else
     rs.MoveFirst
     Do Until rs.EOF
         lstVoucNo.AddItem rs!vouc_no
         rs.MoveNext
     Loop
 End If
    
End Sub

Private Function ValidateData()

ValidateData = False
If txtdate = "" Or IsDate(txtdate.TextWithMask) = False Then
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
ElseIf txtTtlDesc.Text = "" Or IsNumeric(txtTtlDesc) = True Then
  MsgBox "Invalid Description", vbInformation, "Invalid Entry"
  txtTtlDesc.SetFocus
  Exit Function
ElseIf txtChequeNO.Text = "" Then
  MsgBox "Invalid Cheque No.", vbInformation, "Invalid Entry"
  txtChequeNO.SetFocus
  Exit Function
ElseIf IsDate(txtchequedt.TextWithMask) = False Then
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
    .Cols = 6
    .Col = 0
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Cheque_No"
    .ColAlignment(0) = 0
    .ColWidth(0) = 1000
    .ColWidth(1) = 1225
    .ColWidth(2) = 2200
    .ColWidth(3) = 1000
    .ColWidth(4) = 4150
    .ColWidth(5) = 900
    .Col = 1
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Chq_Date"
    .Col = 2
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Chq_Bank"
    .Col = 3
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Code"
    .Col = 4
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Dep_Bank"
    .Col = 5
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Amount"
    .Row = 0
    .Col = 1
  
  End With
End Sub

Private Sub lstAcctCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtTtlDesc.SetFocus
End Sub

Private Sub lstBankCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtamount.SetFocus
End Sub

Private Sub lstVoucNo_Click()
Dim i
Dim X
Dim Y
Dim Z
Dim U
Dim j
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Val(lstVoucNo.Text)
    lblTtlAmount.Caption = 0
        Sqlqry = " Select * from Prpt_mas1 Where Vouc_no= " & i & " and status='N'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
         If rs.RecordCount = 0 Then
          MsgBox " Record not found", vbInformation, "Deleted Status"
          Exit Sub
         Else
           Sqlqry1 = " Select sum(Tra_amount) from prpt_mas1 where vouc_no=" & i
           Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
           If IsNull(rs1.Fields(0)) = False Then txtTtlAmount.Text = Val(rs1.Fields(0))
           
           Sqlqry1 = " Select sum(Tra_amount) from prpt_mas1 where vouc_no=" & i & " and status ='N'"
           Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
           If IsNull(rs1.Fields(0)) = False Then txtpendingamount.Text = Val(rs1.Fields(0))
           
           txtdate.TextWithMask = Format(rs!tDate, "dd/mm/yyyy")
           txtRecdFrom.Text = Trim(rs!Recd_From)
           CboCurrency.Text = Trim(rs!tcurrency)
           txtTtlDesc.Text = Trim(UCase(rs!Description))
             Sqlqry1 = "Select Supp_no,Supp_name from supp_Fin where supp_no='" & Trim(rs!acct_code) & "' order by supp_no"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                 lstAcctCode.Text = rs1!Supp_no & "    :  " & rs1!SUPP_NAME
               Else
                 Sqlqry2 = "Select * from AGNDTLS where Agentname='" & Trim(rs!acct_name) & "' order by agentname"
                 Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                  If rs2.RecordCount <> 0 Then
                   lstAcctCode.Text = "AGNC" & "    :  " & rs2!agentname
                  Else
                   MsgBox "Selected Code not found in Customer/Supplier Register"
                  End If
                End If
           
                     
                
           MSFlexGrid1.Clear
           Flexitems
           j = 0
           rs.MoveFirst
            Do Until rs.EOF
             MSFlexGrid1.AddItem rs!CHEQUE_NO & Chr(9) & Format(rs!Cheque_Dt, "dd/mm/yyyy") & Chr(9) & rs!cheque_Bank & Chr(9) & rs!bank_code & Chr(9) & rs!BANK_NAME & Chr(9) & rs!tra_amount
             j = j + rs!tra_amount
             rs.MoveNext
            Loop
           lblTtlAmount.Caption = j
           lblTtlAmount.Alignment = 1
           txtdate.SetFocus
         End If
    End Sub

Private Function textclear()
     txtTtlAmount.Text = ""
     txtChequeNO.Text = ""
     txtChequeBank.Text = " "
     txtchequedt.TextWithMask = ""
     txtdate.TextWithMask = ""
     lstBankCode.ListIndex = 0
     lstAcctCode.ListIndex = 0
     txtamount = ""
     txtRecdFrom.Text = ""
     txtTtlDesc.Text = ""
     lblTtlAmount.Caption = ""
     Flexitems
     lstVoucNo.ListIndex = 0
     txtdate.SetFocus
End Function

Public Sub PopulateBankCodes()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select * from Bank_mas order by bank_code"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

lstBankCode.Clear

 If rs.RecordCount = 0 Then
      MsgBox "No Records found in the Bank Register"
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
 Dim Y, Z, U As Long
 Dim MchequeNo, MChequeBank, MBankCode, MDepBank
 Dim MchequeDt As Date
 Dim mamount As Currency
 
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
        MchequeNo = .Text
        .Col = 1
        MchequeDt = .Text
        .Col = 5
        mamount = .Text
       
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " Select * from Prpt_mas1 Where Vouc_no= " & Val(lstVoucNo) & "  AND Cheque_no='" & MchequeNo & "' AND CHEQUE_DT =#" & DateValue(Format(MchequeDt, "DD/MM/YYYY")) & "# AND STATUS ='Y' AND ISNULL(POSTING_DT)<> TRUE"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
           MsgBox "This Transaction already posted this Cannot be removed"
           Exit Sub
        End If
       End With
     With MSFlexGrid1
        j = .Row
        .Col = 0
        txtChequeNO = .Text
        .Col = 1
        txtchequedt = .Text
        .Col = 2
        txtChequeBank = .Text
        .Col = 3
        MBankCode = .Text
        .Col = 4
        MDepBank = .Text
        .Col = 5
        txtamount = .Text
        
        Sqlqry = "Select Bank_code,Bank_name from Bank_mas Where Bank_Code='" & MBankCode & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

        If rs.RecordCount = 0 Then
            MsgBox "No Matching Record found in the Bank Register"
        Else
            lstBankCode.Text = rs!bank_code & " : " & rs!BANK_NAME
        End If
        
                                   
        lblTtlAmount.Caption = Val(lblTtlAmount.Caption) - Val(txtamount)
        Y = Val(lblTtlAmount.Caption)
        .RemoveItem (j)
        
     End With
    End If
   End If
  End If
End Sub

Private Sub txtAmount_LostFocus()
   Dim X As Long
   Dim Y As Long
   
   If ValidateData = True Then
    If Val(txtamount.Text) > Val(txtpendingamount.Text) Then
      MsgBox " Entered Amount Greater than Total Pending Amount"
      txtamount.SetFocus
    Exit Sub
    End If
      
   X = Val(lblTtlAmount.Caption) + Val(txtamount)
   If X > Val(txtpendingamount) Then
    MsgBox "Entered Amount Greater Than Total Pending Amount"
    Exit Sub
   Else
    MSFlexGrid1.AddItem txtChequeNO & Chr(9) & Format(txtchequedt.TextWithMask, "dd/mm/yyyy") & Chr(9) & txtChequeBank & Chr(9) & Mid(lstBankCode, 1, 6) & Chr(9) & Mid(lstBankCode, 10, 35) & Chr(9) & txtamount
    lblTtlAmount = lblTtlAmount + Val(txtamount)
   End If
  End If
End Sub

Private Sub txtChequeBank_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstBankCode.SetFocus
End Sub

Private Sub txtChequeDt_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtChequeBank.SetFocus
End Sub

Private Sub txtchequedt_LostFocus()
Dim a As Date
Dim B As Date

If Mid(txtchequedt.TextWithMask, 4, 2) > 12 Then
      MsgBox "Invalid  cheque Date ", vbInformation, "Invalid Entry"
      txtchequedt.SetFocus
      SendKeys "{Home} + {End}"
End If

If IsDate(txtchequedt.TextWithMask) = False Then
      MsgBox "Invalid  cheque Date ", vbInformation, "Invalid Entry"
      txtchequedt.SetFocus
      SendKeys "{Home} + {End}"
End If


a = Format(txtchequedt.TextWithMask, "dd/mm/yyyy")
B = Format(Now(), "dd/mm/yyyy")

If DateValue(a) <= DateValue(B) Then
  MsgBox "Cheque date cannot be lesser than or equal to current date"
  txtchequedt.SetFocus
  Exit Sub
End If
End Sub

Private Sub txtChequeNO_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtchequedt.SetFocus
End Sub


Private Sub txtConvRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtpendingamount.SetFocus
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboCurrency.SetFocus
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtamount.SetFocus
End Sub

Private Sub txtdate_LostFocus()
If IsDate(txtdate.TextWithMask) = False Then
      MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
      txtdate.SetFocus
      SendKeys "{Home} + {End}"
 End If
End Sub

Private Sub txtrecdfrom_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then lstAcctCode.SetFocus
End Sub

Private Sub txtpendingamount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRecdFrom.SetFocus
End Sub

Private Sub txtTtlAmount_KeyPress(KeyAscii As Integer)
 If CboCurrency.Text = "USD" Then
  If KeyAscii = 13 Then txtConvRate.SetFocus
 Else
  If KeyAscii = 13 Then txtpendingamount.SetFocus
 End If
End Sub

Private Sub txtTtlDesc_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtChequeNO.SetFocus
End Sub
