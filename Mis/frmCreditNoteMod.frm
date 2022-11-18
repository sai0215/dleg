VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmCreditNoteMod 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CreditNoteModification"
   ClientHeight    =   8505
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Credit Note  - Modification"
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
      Height          =   8175
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   11175
      Begin VB.TextBox txtProduct 
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
         Left            =   7965
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtAgency 
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
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txtCrntPer 
         BackColor       =   &H80000014&
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
         Left            =   10065
         TabIndex        =   34
         Top             =   2400
         Width           =   780
      End
      Begin VB.TextBox txtGross 
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TxtAgCom 
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
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtAddDiscount 
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
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtAdddiscountper 
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtNet 
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
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtConvRate 
         BackColor       =   &H80000014&
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
         Left            =   10080
         TabIndex        =   19
         Top             =   3000
         Width           =   780
      End
      Begin VB.ComboBox cboCurrency 
         BackColor       =   &H80000014&
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
         Height          =   360
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H80000014&
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
         Left            =   6360
         TabIndex        =   3
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ListBox lstVoucNo 
         BackColor       =   &H80000014&
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
         Height          =   2700
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtdesc 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   350
         Left            =   1440
         MaxLength       =   200
         TabIndex        =   7
         Top             =   6240
         Width           =   9495
      End
      Begin VB.TextBox txtRef 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   350
         Left            =   6360
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.ListBox lstDebitedTo 
         BackColor       =   &H80000014&
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
         Height          =   1980
         Left            =   5880
         TabIndex        =   5
         Top             =   3960
         Width           =   4935
      End
      Begin VB.ListBox lstCreditedTo 
         BackColor       =   &H80000014&
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
         Height          =   1980
         Left            =   120
         TabIndex        =   6
         Top             =   3960
         Width           =   5175
      End
      Begin VB.TextBox txtDesc1 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   350
         Left            =   1440
         MaxLength       =   200
         TabIndex        =   8
         Top             =   6720
         Width           =   9495
      End
      Begin VB.CommandButton CmdModify 
         BackColor       =   &H00E0E0E0&
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
         Height          =   705
         Left            =   4080
         Picture         =   "frmCreditNoteMod.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7320
         Width           =   975
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
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
         Height          =   705
         Left            =   5040
         Picture         =   "frmCreditNoteMod.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7320
         Width           =   975
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00E0E0E0&
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
         Height          =   705
         Left            =   6000
         Picture         =   "frmCreditNoteMod.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7320
         Width           =   975
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
         Left            =   3240
         TabIndex        =   1
         Top             =   480
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
         ForeColor       =   0
         Alignment       =   1
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   7065
         TabIndex        =   39
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label lblAgency 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   2280
         TabIndex        =   37
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CRNT %"
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
         Height          =   240
         Left            =   9120
         TabIndex        =   35
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gross"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   1680
         TabIndex        =   33
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ag Comm (%)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   4920
         TabIndex        =   32
         Top             =   1320
         Width           =   1410
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Additional  Discount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   7920
         TabIndex        =   31
         Top             =   1320
         Width           =   2085
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Disc. (%)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   1680
         TabIndex        =   30
         Top             =   2400
         Width           =   1425
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Net"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   5880
         TabIndex        =   29
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   120
         TabIndex        =   23
         Top             =   6840
         Width           =   945
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   8880
         TabIndex        =   22
         Top             =   3120
         Width           =   1155
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   1920
         TabIndex        =   21
         Top             =   3120
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   4920
         TabIndex        =   20
         Top             =   3120
         Width           =   1380
      End
      Begin VB.Line Line4 
         X1              =   11160
         X2              =   -120
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Line Line3 
         X1              =   5520
         X2              =   5520
         Y1              =   3600
         Y2              =   6120
      End
      Begin VB.Line Line2 
         X1              =   11160
         X2              =   -120
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   11160
         Y1              =   3600
         Y2              =   3600
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Debit To"
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
         Height          =   240
         Left            =   5880
         TabIndex        =   17
         Top             =   3600
         Width           =   915
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   1800
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   5040
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Credit To"
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
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   6360
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmCreditNoteMod"
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
Dim sqlqry3 As String
Dim X
Dim Y
Dim Z
Dim i
Dim CrdCur
Dim j
Public MTYPE
Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtamount.SetFocus
End Sub
Private Sub cboCurrency_LostFocus()
    If CboCurrency.Text = "USD" Then
        lblConvRate.Visible = True
        txtConvRate.Visible = True
        txtConvRate.Text = ""
        txtConvRate.TabIndex = 4
     Else
        lblConvRate.Visible = False
        txtConvRate.Visible = False
        txtConvRate.Text = 1
        txtConvRate.TabIndex = 12
    End If
End Sub
Private Sub cmdBack_Click()
 Unload Me
End Sub
Private Sub cmdClear_Click()
 textclear
End Sub

Private Sub Cmdmodify_Click()
  If ValidateData = True Then
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       Sqlqry = " Update crdt_mas set tdate=#" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "#," & _
                       " acct_code = '" & Mid(lstCreditedTo, 1, 6) & "'," & _
                       " acct_name ='" & Mid(lstCreditedTo, 12, 35) & "'," & _
                       " Supp_no ='" & Mid(lstDebitedTo, 1, 4) & "'," & _
                       " Supp_name ='" & Trim(Mid(lstDebitedTo, 9, 35)) & "'," & _
                       " Ref_no ='" & findfirstfixup(UCase(Trim(txtRef))) & "'," & _
                       " Description ='" & findfirstfixup(UCase(Trim(txtdesc))) & "'," & _
                       " Description1='" & findfirstfixup(UCase(Trim(txtDesc1))) & "'," & _
                       " CRNTPer=" & Round(Val(txtCrntPer), 2) & "," & _
                       " Tcurrency='" & Trim(CboCurrency) & "'," & _
                       " Tconvertion=" & Val(txtConvRate) & "," & _
                       " Tra_Amount=" & Val(txtamount) & "," & _
                       " Amount =" & Val(txtamount.Text) * Val(txtConvRate.Text) & " where vouc_no=" & Val(lstVoucNo) & ""
                 
       ws.BeginTrans
       db.Execute (Sqlqry)
       ws.CommitTrans
      
  MsgBox " Record is Modified", vbInformation, "Status"
  Dim X
  X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
  If X = vbYes Then
    MTYPE = Trim(Mid(MTYPE, 1, 11))
    If MTYPE = "ZEINA" Or MTYPE = "ALAM ASSAYA" Then
   
            If CboCurrency = "DHS" Then
            
                 CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
                 CrystalReport1.ReportFileName = App.Path & "\CRNTVOUMPS.rpt"
                 CrystalReport1.SelectionFormula = "{crdt_mas.Vouc_no}=" & Val(lstVoucNo.Text) & ""
                 CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtamount)) & " Only" & "'"
                 CrystalReport1.Formulas(1) = "curtype='" & CboCurrency.Text & "'"
                 CrystalReport1.WindowState = crptMaximized
                 CrystalReport1.Action = 1
              Else
                 CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
                 CrystalReport1.ReportFileName = App.Path & "\CRNTVOUMPS.rpt"
                 CrystalReport1.SelectionFormula = "{crdt_mas.Vouc_no}=" & Val(lstVoucNo.Text) & ""
                 CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(txtamount)) & " Only" & "'"
                 CrystalReport1.Formulas(1) = "curtype='" & CboCurrency.Text & "'"
                 CrystalReport1.WindowState = crptMaximized
                 CrystalReport1.Action = 1
             End If
       Else
           If CboCurrency = "DHS" Then
   
            CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
            CrystalReport1.ReportFileName = App.Path & "\CRNTVOU.rpt"
            CrystalReport1.SelectionFormula = "{crdt_mas.Vouc_no}=" & Val(lstVoucNo.Text) & ""
            CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtamount)) & " Only" & "'"
            CrystalReport1.Formulas(1) = "curtype='" & CboCurrency.Text & "'"
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Action = 1
            Else
            CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
            CrystalReport1.ReportFileName = App.Path & "\CRNTVOU.rpt"
            CrystalReport1.SelectionFormula = "{crdt_mas.Vouc_no}=" & Val(lstVoucNo.Text) & ""
            CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(txtamount)) & " Only" & "'"
            CrystalReport1.Formulas(1) = "curtype='" & CboCurrency.Text & "'"
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Action = 1
          End If
    End If
  End If
  
  textclear
  End If
  
End Sub

Private Sub Form_Load()
 txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
 CboCurrency.AddItem "DHS"
 CboCurrency.AddItem "USD"
 lblConvRate.Visible = False
 txtConvRate.Visible = False
 txtCrntPer.Text = 0
 txtConvRate.Text = 1
 PopulateVoucher
 PopulateAcctSuppCust
 PopulateAcctSuppCust1
 End Sub
Private Sub PopulateVoucher()
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Select Vouc_No from crdt_mas where status='N' order by vouc_no"
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
      lstDebitedTo.AddItem rs1!Supp_no & "  :  " & rs1!supp_name
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
ElseIf txtamount.Text = "" Or IsNumeric(txtamount) = False Then
  MsgBox "Invalid Amount", vbInformation, "Invalid Entry"
  txtamount.SetFocus
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
     CboCurrency.ListIndex = -1
     txtConvRate.Visible = False
     lblConvRate.Visible = False
     lstDebitedTo.ListIndex = 0
     lstCreditedTo.ListIndex = 0
     txtRef.Text = ""
     txtdesc.Text = ""
     txtDesc1.Text = ""
     txtamount.Text = ""
     txtGross.Text = ""
     TxtAgCom.Text = ""
     txtAddDiscount.Text = ""
     txtAdddiscountper.Text = ""
     txtNet.Text = ""
     txtCrntPer.Text = 0
     txtAgency.Text = ""
     txtproduct.Text = ""
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
    Dim Y
    Dim Z
    Dim U
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Val(lstVoucNo.Text)
        
        Sqlqry = " Select * from crdt_mas Where Vouc_no= " & i
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
         If rs.RecordCount <> 0 Then
          
           txtdate.TextWithMask = Format(rs!tDate, "dd/mm/yyyy")
           If rs!tcurrency = "DHS" Then
             CboCurrency.ListIndex = 0
             txtConvRate.Text = rs!tconvertion
           Else
             CboCurrency.ListIndex = 1
             txtConvRate.Text = rs!tconvertion
           End If
                                
           If IsNull(rs!tra_amount) = True Then
            txtamount = 0
           Else
            txtamount = rs!tra_amount
           End If
           
           If IsNull(rs!ref_no) = True Then
             txtRef = 0
           Else
             txtRef = rs!ref_no
           End If
                         Set ws = DBEngine.Workspaces(0)
                         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         Sqlqry1 = "Select * from bo_mas where serial_no='" & Mid(txtRef, 1, 7) & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                           TxtAgCom.Text = ""
                           txtGross.Text = ""
                           txtAddDiscount.Text = ""
                           txtAdddiscountper.Text = ""
                           txtNet.Text = ""
                           
                           
                        If rs1.RecordCount <> 0 Then
                          rs1.MoveFirst
                          CrdCur = rs1!tcurrency
                          txtGross = rs1!tra_gamount
                          TxtAgCom = rs1!disc_rate
                          txtAdddiscountper.Text = rs1!disc_percentage
                          txtAddDiscount.Text = rs1!add_discount
                          txtNet = rs1!tra_namount
                        End If
                        
           If IsNull(rs!Description) = True Then
            txtdesc = ""
           Else
            txtdesc = rs!Description
           End If
           If IsNull(rs!DESCRIPTION1) = True Then
            txtDesc1 = ""
           Else
            txtDesc1 = rs!DESCRIPTION1
           End If
           
           If IsNull(rs!CRNTPER) = True Then
            txtCrntPer = ""
           Else
            txtCrntPer = rs!CRNTPER
           End If
           
         
         sqlqry3 = "Select acct_Code,acct_name from Acct_mas where acct_code='" & rs!acct_code & "' order by acct_code"
         Set rs3 = db.OpenRecordset(sqlqry3, dbOpenDynaset)
           If rs3.RecordCount <> 0 Then
                  lstCreditedTo.Text = rs3!acct_code & "  :  " & rs3!acct_name
           End If
           
           
         Sqlqry = "Select * from Agndtls where agentname='" & Trim(rs!supp_name) & "' order by agentname"
         Set rs1 = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs1.RecordCount <> 0 Then
                  lstDebitedTo.Text = "AGNC" & "  :  " & rs1!agentname
           Else
             Sqlqry1 = "Select Supp_no,Supp_name from supp_Fin where supp_no='" & Trim(rs!Supp_no) & "' order by supp_no"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                lstDebitedTo.Text = Trim(rs1!Supp_no) & "  :  " & rs1!supp_name
               End If
           End If
           txtdate.SetFocus
         End If
End Sub
Private Sub txtamount_KeyPress(KeyAscii As Integer)
   If CboCurrency.Text = "USD" Then
     If KeyAscii = 13 Then txtConvRate.SetFocus
   Else
     If KeyAscii = 13 Then lstCreditedTo.SetFocus
     txtConvRate.Text = 1
   End If
End Sub
Private Sub txtConvRate_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then lstCreditedTo.SetFocus
End Sub
Private Sub txtCrntPer_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboCurrency.SetFocus
End Sub
Private Sub txtCrntPer_LostFocus()
If txtCrntPer.Text = "" Then txtCrntPer = 0
If txtCrntPer <> 0 Then
 txtamount.Text = Round(Val(txtGross - (txtGross * TxtAgCom / 100)) * txtCrntPer / 100, 2)
End If
End Sub
Private Sub txtDate_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtRef.SetFocus
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
If KeyAscii = 13 Then cmdModify.SetFocus
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCrntPer.SetFocus
End Sub

Private Sub txtRef_LostFocus()
 If txtRef.Text = "" Or Val(txtRef.Text) = 0 Then
  txtCrntPer.SetFocus
  Exit Sub
Else

If Len(txtRef.Text) > 7 Then
 MsgBox "Invalid Reference number"
 Exit Sub
End If

 CrdCur = ""
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Select * from bo_mas where serial_no='" & Mid(txtRef, 1, 7) & "'"
 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    TxtAgCom.Text = ""
    txtGross.Text = ""
    txtAddDiscount.Text = ""
    txtAdddiscountper.Text = ""
    txtNet.Text = ""
    txtAgency.Text = ""
    txtproduct.Text = ""
    
    
    
 If rs.RecordCount = 0 Then
     MsgBox "Reference No. is not matching with the Booking Order No."
     txtRef.SetFocus
     Exit Sub
 Else
   rs.MoveFirst
   CrdCur = rs!tcurrency
   txtGross = rs!tra_gamount
   TxtAgCom = rs!disc_rate
   txtAgency = rs!Agency
   txtproduct = rs!Product
   txtAdddiscountper.Text = rs!disc_percentage
   txtAddDiscount.Text = rs!add_discount
   txtNet = rs!tra_namount
   MTYPE = rs!sub_Media
   lstDebitedTo.Text = "AGNC" & "  :  " & rs!Agency
   lstCreditedTo.Text = "301000 : Sales"
 End If
 
 
End If
 End Sub
