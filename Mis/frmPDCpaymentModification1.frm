VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmPDCPaymentModification1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "PDC Payment Modification"
   ClientHeight    =   8775
   ClientLeft      =   -90
   ClientTop       =   315
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "PDC Payment Modification"
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
      Height          =   8415
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   11655
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   4200
         Width           =   3615
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
            TabIndex        =   11
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
            TabIndex        =   10
            Top             =   0
            Width           =   1215
         End
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
            TabIndex        =   9
            Top             =   0
            Width           =   1095
         End
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
         Left            =   9645
         TabIndex        =   3
         Top             =   480
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
         Left            =   6885
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1095
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
         Left            =   4320
         TabIndex        =   33
         Top             =   1080
         Width           =   1455
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
         Left            =   9960
         TabIndex        =   14
         Top             =   4560
         Width           =   1335
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
         Left            =   4560
         TabIndex        =   13
         Top             =   4560
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
         Height          =   1020
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   4560
         Width           =   4335
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFF80&
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
         Left            =   3720
         Picture         =   "frmPDCpaymentModification1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7440
         Width           =   1095
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
         Height          =   855
         Left            =   5880
         Picture         =   "frmPDCpaymentModification1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   7440
         Width           =   1095
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
         Height          =   855
         Left            =   4800
         Picture         =   "frmPDCpaymentModification1.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   7440
         Width           =   1095
      End
      Begin VB.CommandButton cmdModify 
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
         Height          =   855
         Left            =   2640
         Picture         =   "frmPDCpaymentModification1.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7440
         Width           =   1095
      End
      Begin VB.TextBox txtTtlDesc 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1680
         TabIndex        =   8
         Top             =   3600
         Width           =   9615
      End
      Begin VB.TextBox txtPaidTo 
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
         Height          =   350
         Left            =   1680
         TabIndex        =   7
         Top             =   3000
         Width           =   4335
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
         Height          =   1260
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   1335
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
         Height          =   350
         Left            =   4320
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
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
         Height          =   780
         Left            =   1680
         TabIndex        =   6
         Top             =   2040
         Width           =   4335
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Preview"
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
         Left            =   6960
         Picture         =   "frmPDCpaymentModification1.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   7440
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   6720
         Top             =   4560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1215
         Left            =   120
         TabIndex        =   21
         Top             =   5640
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   4
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   -2147483647
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
      Begin PVMaskEditLib.PVMaskEdit txtchequedt 
         Height          =   375
         Left            =   6960
         TabIndex        =   5
         Top             =   1560
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
         BorderColor     =   &H00004000&
         BorderWidth     =   2
         X1              =   11640
         X2              =   0
         Y1              =   7320
         Y2              =   7320
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
         Left            =   8205
         TabIndex        =   36
         Top             =   600
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
         Left            =   5880
         TabIndex        =   35
         Top             =   600
         Width           =   930
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
         Left            =   3120
         TabIndex        =   34
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404080&
         BorderWidth     =   2
         X1              =   11640
         X2              =   0
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label8 
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
         Left            =   4920
         TabIndex        =   32
         Top             =   4320
         Width           =   1200
      End
      Begin VB.Label Label3 
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
         Left            =   10200
         TabIndex        =   31
         Top             =   4320
         Width           =   780
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
         Left            =   8640
         TabIndex        =   30
         Top             =   6960
         Width           =   1380
      End
      Begin VB.Label lblTtlAmount 
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   10080
         TabIndex        =   29
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Payee Name "
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
         Left            =   180
         TabIndex        =   28
         Top             =   3120
         Width           =   1425
      End
      Begin VB.Label Label4 
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
         Left            =   3120
         TabIndex        =   27
         Top             =   600
         Width           =   975
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
         Left            =   120
         TabIndex        =   26
         Top             =   3720
         Width           =   1440
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
         TabIndex        =   25
         Top             =   600
         Width           =   1410
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cheque Dt"
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
         Left            =   5820
         TabIndex        =   24
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cheque #"
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
         TabIndex        =   23
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Bank Name "
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
         TabIndex        =   22
         Top             =   2160
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmPDCPaymentModification1"
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
If cboCurrency.Text = "USD" Then
  txtConvRate.Text = ""
  lblConvRate.Visible = True
  txtConvRate.Visible = True
  txtConvRate.TabIndex = 4
Else
  txtConvRate.Text = 1
  lblConvRate.Visible = False
  txtConvRate.Visible = False
  txtConvRate.TabIndex = 23
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
 MsgBox "You are not allowed to delete Consult System Administrator"
 Exit Sub
End Sub

Private Sub Cmdmodify_Click()
 Dim a
 Dim B
 Dim C
 Dim X
 Dim accode, acdesc, acname
 Dim acamt As Currency
 
 If txtTtlAmount <> Val(lblTtlAmount) Then
  MsgBox "Total Amount is Not tallying with Transaction Amount"
  Exit Sub
 End If
 
X = MsgBox("Do You Want to Modify PDC Payment Voucher No." & Val(lstVoucNo), vbInformation + vbYesNo, "Confirm")
 
If X = vbNo Then Exit Sub
If ValidateData = True Then
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    
   Sqlqry = "Update PPMT_MAS set TDATE=#" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "#," & _
                            " Paid_To ='" & UCase(Trim(txtPaidTo)) & "'," & _
                            " Bank_Code ='" & Mid(lstBankCode, 1, 6) & "'," & _
                            " Bank_name ='" & Mid(lstBankCode, 10, 30) & "'," & _
                            " Cheque_No ='" & UCase(Trim(txtChequeNO)) & "'," & _
                            " Cheque_Dt =#" & Format(txtchequedt.TextWithMask, "dd/mm/yyyy") & "#," & _
                            " description='" & findfirstfixup(UCase(Trim(txtTtlDesc))) & "'," & _
                            " Tcurrency='" & Trim(cboCurrency) & "'," & _
                            " Tconvertion=" & Val(Trim(txtConvRate)) & "," & _
                            " Tra_amount=" & Val(Trim(txtTtlAmount)) & "," & _
                            " TTl_Amount =" & Val(txtTtlAmount) * Val(txtConvRate) & " Where VOUC_NO = " & Val(lstVoucNo.Text) & " ;"
   ws.BeginTrans
   db.Execute Sqlqry
   ws.CommitTrans
 
   Sqlqry1 = "Delete * from PPMT_TRA where vouc_no=" & Val(lstVoucNo) & " "
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
      
       Sqlqry2 = "Insert into PPMT_TRA values(" & Val(lstVoucNo.Text) & ",'PPT','" _
                                     & Format(txtdate.TextWithMask, "DD/MM/YYYY") & "','" _
                                     & accode & "','" _
                                     & findfirstfixup(acname) & "','" _
                                     & findfirstfixup(acdesc) & "','" _
                                     & Format(txtchequedt.TextWithMask, "DD/MM/YYYY") & "','','" _
                                     & Trim(cboCurrency) & "'," _
                                     & Val(txtConvRate) & ", " _
                                     & Val(acamt) & ", " _
                                     & Val(acamt) * Val(txtConvRate) & ")"

        ws.BeginTrans
        db.Execute Sqlqry2
        ws.CommitTrans
     Next
   End With
      
  MsgBox " PDC Payment Voucher is modified"
  textclear
  Flexitems
  lstVoucNo.SetFocus
 End If

End Sub

Private Sub CmdPrint_Click()
Dim ctype As String
   If lstVoucNo.SelCount = 0 Then
    MsgBox "Select Voucher Number"
    lstVoucNo.SetFocus
   End If
   ctype = cboCurrency.Text
   If ctype = "DHS" Then
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\ppmtVou.rpt"
        CrystalReport1.SelectionFormula = "{PPMT_TRA.Vouc_no}=" & Val(lstVoucNo.Text) & ""
        CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtTtlAmount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    Else
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\ppmtVou.rpt"
        CrystalReport1.SelectionFormula = "{PPMT_TRA.Vouc_no}=" & Val(lstVoucNo.Text) & ""
        CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(txtTtlAmount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1

    End If
End Sub

Private Sub Form_Load()
 
 txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
 txtchequedt.TextWithMask = Format(Now, "dd/mm/yyyy")
 cboCurrency.AddItem "DHS"
 cboCurrency.AddItem "USD"
 
 lblConvRate.Visible = False
 txtConvRate.Visible = False
  
txtConvRate.Text = 1

 LSTSUP = 0
 LSTAC = 0
 LSTAGN = 1
 
 PopulateVoucher
 PopulateBankCodes
 PopulateAcctSuppCust
 Flexitems
 
End Sub

Private Sub PopulateAcctSuppCust()
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

Private Sub PopulateVoucher()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select * from PPMT_MAS where status='N' ORDER BY VOUC_NO"
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
ElseIf txtConvRate.Text = "" Then
  MsgBox "Enter Convertion Rate - - cannot be zero", vbInformation, "Invalid Entry"
  txtConvRate.SetFocus
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
    .Cols = 4
    .Col = 0
    .CellBackColor = RGB(180, 170, 160)
    .CellAlignment = 2
    .Text = "Ac. Code"
    .ColAlignment(0) = 0
    .ColWidth(0) = 1100
    .ColWidth(1) = 3250
    .ColWidth(2) = 6000
    .ColWidth(3) = 900
    .Col = 1
    .CellBackColor = RGB(180, 170, 160)
    .CellAlignment = 2
    .Text = "Account Name"
    .Col = 2
    .CellBackColor = RGB(180, 170, 160)
    .CellAlignment = 2
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
    cmdModify.Enabled = True
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Val(lstVoucNo.Text)
        
        Sqlqry = " Select * from PPMT_MAS Where Vouc_no= " & i
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
           txtdate.TextWithMask = Format(rs!tDate, "dd/mm/yyyy")
           txtTtlAmount = rs!tra_amount
           txtChequeNO = rs!CHEQUE_NO
           txtchequedt.TextWithMask = Format(rs!Cheque_Dt, "dd/mm/yyyy")
        End If
          Sqlqry1 = "Select BANK_CODE,BANK_NAME from Bank_mas WHERE BANK_CODE='" & Trim(rs!bank_code) & "' ORDER by bank_code"
          Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
          If rs1.RecordCount = 0 Then
             MsgBox "Bank Code Not Found"
          Else
            lstBankCode.Text = rs1!bank_code & " : " & rs1!BANK_NAME
          End If
                                        
           If rs!tcurrency = "USD" Then
             cboCurrency.ListIndex = 1
             lblConvRate.Visible = True
             txtConvRate.Visible = True
             txtConvRate.TabIndex = 4
           Else
             cboCurrency.ListIndex = 0
           End If
             
           txtPaidTo = rs!PAID_TO
           txtConvRate.Text = rs!tconvertion
           txtTtlDesc = rs!Description
           
           lblTtlAmount.Caption = 0
         
         Sqlqry1 = "Select * from PPMT_TRA where Vouc_no= " & i
         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
         If rs1.RecordCount = 0 Then
           MsgBox " Particular record was deleted.", vbInformation, "Deleted Status"
           Exit Sub
         Else
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
         If IsNull(rs!posting_Dt) <> True Then
           MsgBox "Transaction is posted, cannot modify it"
           lstVoucNo.SetFocus
           cmdModify.Enabled = False
         End If
         
    End Sub

Private Function textclear()
     txtTtlAmount.Text = ""
     txtChequeNO.Text = ""
     txtchequedt.TextWithMask = ""
     txtConvRate.Text = ""
     cboCurrency.ListIndex = -1
     lstBankCode.ListIndex = 0
     lstAcctCode.ListIndex = 0
     txtdesc = ""
     txtAmount = ""
     txtPaidTo.Text = ""
     txtTtlDesc.Text = ""
     lblTtlAmount.Caption = ""
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from dumbpmt1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     MSFlexGrid1.Clear
     lstVoucNo.ListIndex = 0
     lstVoucNo.SetFocus
     
End Function

Public Sub PopulateBankCodes()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
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

Private Sub txtAmount_LostFocus()
   Dim X
   Dim y
   Dim accd As String
   Dim acname As String
   
  If LSTAC = 1 Then
     accd = Mid(lstAcctCode, 1, 6)
     acname = Mid(lstAcctCode, 12, 35)
  ElseIf LSTSUP = 1 Then
     accd = Mid(lstAcctCode, 1, 4)
     acname = Mid(lstAcctCode, 10, 35)
  Else
     accd = "AGNC"
     acname = Trim(lstAcctCode.Text)
  End If
  
 
    cur = ""
    con = 1
 
  If cboCurrency.Text = "USD" Then
      cur = "USD"
      con = Val(Trim(txtConvRate.Text))
       
  Else
      cur = "DHS"
      con = 1
  End If
   
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
    MSFlexGrid1.AddItem accd & Chr(9) & acname & Chr(9) & txtdesc & Chr(9) & txtAmount
    lblTtlAmount = Val(lblTtlAmount) + Val(txtAmount)
   End If
  End If
  txtdesc.Text = Trim(txtTtlDesc.Text)
End Sub
Private Sub txtChequeDt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstBankCode.SetFocus
End Sub

Private Sub txtchequedt_LostFocus()
If IsDate(txtchequedt.TextWithMask) = False Then
      MsgBox "Invalid Cheque Date ", vbInformation, "Invalid Entry"
      txtchequedt.SetFocus
      SendKeys "{Home} + {End}"
End If
End Sub

Private Sub txtChequeNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtchequedt.SetFocus
End Sub
Private Sub txtConvRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtChequeNO.SetFocus
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
Private Sub txtpaidto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTtlDesc.SetFocus
End Sub
Private Sub txtTtlAmount_KeyPress(KeyAscii As Integer)
    If cboCurrency = "USD" Then
     If KeyAscii = 13 Then txtConvRate.SetFocus
    Else
     If KeyAscii = 13 Then txtChequeNO.SetFocus
    End If
End Sub

Private Sub txtTtlDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstAcctCode.SetFocus
End Sub

Private Sub txtTtlDesc_LostFocus()
   txtdesc.Text = Trim(txtTtlDesc.Text)
End Sub
