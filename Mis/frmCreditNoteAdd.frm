VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmCreditNoteAdd 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CreditNoteAddition"
   ClientHeight    =   8595
   ClientLeft      =   -60
   ClientTop       =   285
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   8175
      Left            =   840
      TabIndex        =   12
      Top             =   120
      Width           =   10095
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1680
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1680
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
         Left            =   8400
         TabIndex        =   34
         Top             =   2280
         Width           =   900
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2280
         Width           =   1335
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2280
         Width           =   1095
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1080
         Width           =   855
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1080
         Width           =   855
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1080
         Width           =   1335
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
         Height          =   315
         Left            =   4560
         TabIndex        =   2
         Top             =   3000
         Width           =   1575
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   3000
         Width           =   1095
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
         Height          =   315
         Left            =   8400
         TabIndex        =   11
         Top             =   3000
         Width           =   1020
      End
      Begin VB.TextBox txtdesc 
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
         Height          =   350
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   6
         Top             =   6120
         Width           =   8175
      End
      Begin VB.TextBox txtRef 
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
         Height          =   350
         Left            =   8400
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.ListBox lstDebitedTo 
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
         Height          =   1860
         Left            =   5280
         TabIndex        =   4
         Top             =   3960
         Width           =   4695
      End
      Begin VB.ListBox lstCreditedTo 
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
         Height          =   1860
         Left            =   240
         TabIndex        =   5
         Top             =   3960
         Width           =   4575
      End
      Begin VB.TextBox txtDesc1 
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
         Height          =   350
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   7
         Top             =   6720
         Width           =   8175
      End
      Begin VB.CommandButton CmdAdd 
         BackColor       =   &H00E0E0E0&
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
         Height          =   705
         Left            =   2880
         Picture         =   "frmCreditNoteAdd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7440
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
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
         Height          =   705
         Left            =   3960
         Picture         =   "frmCreditNoteAdd.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7440
         Width           =   975
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00E0E0E0&
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
         Height          =   705
         Left            =   4920
         Picture         =   "frmCreditNoteAdd.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7440
         Width           =   975
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   480
         Top             =   7440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin PVMaskEditLib.PVMaskEdit txtdate 
         Height          =   375
         Left            =   4560
         TabIndex        =   0
         Top             =   360
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
         Left            =   5340
         TabIndex        =   39
         Top             =   1800
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
         Left            =   555
         TabIndex        =   37
         Top             =   1800
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
         Left            =   7455
         TabIndex        =   35
         Top             =   2400
         Width           =   900
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
         Left            =   4080
         TabIndex        =   33
         Top             =   2400
         Width           =   375
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
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   1425
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
         Left            =   6240
         TabIndex        =   29
         Top             =   1200
         Width           =   2085
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
         Left            =   3120
         TabIndex        =   27
         Top             =   1200
         Width           =   1410
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
         Left            =   600
         TabIndex        =   25
         Top             =   1200
         Width           =   750
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
         Height          =   240
         Left            =   240
         TabIndex        =   23
         Top             =   6840
         Width           =   945
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
         Left            =   3120
         TabIndex        =   22
         Top             =   3120
         Width           =   1380
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
         Left            =   240
         TabIndex        =   21
         Top             =   3120
         Width           =   1170
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
         Left            =   7200
         TabIndex        =   20
         Top             =   3120
         Width           =   1155
      End
      Begin VB.Line Line4 
         X1              =   10080
         X2              =   0
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line3 
         X1              =   5040
         X2              =   5040
         Y1              =   3600
         Y2              =   6000
      End
      Begin VB.Line Line2 
         X1              =   10080
         X2              =   0
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10080
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1365
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
         Left            =   5280
         TabIndex        =   18
         Top             =   3720
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
         Left            =   3960
         TabIndex        =   17
         Top             =   480
         Width           =   510
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
         Left            =   7080
         TabIndex        =   16
         Top             =   480
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
         Left            =   240
         TabIndex        =   15
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lblVoucNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
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
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   360
         Width           =   1035
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
         Left            =   240
         TabIndex        =   13
         Top             =   6240
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmCreditNoteAdd"
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
Dim sqlqry3 As String
Dim X
Dim Y
Dim Z
Dim i
Dim j
Dim CrdCur
Dim MTYPE

Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtamount.SetFocus
End Sub

Private Sub cboCurrency_LostFocus()
If txtRef.Text <> "" Then
    If CboCurrency <> CrdCur Then
       MsgBox " Reference Booking order booked in different currency"
       CboCurrency.SetFocus
       Exit Sub
    End If
End If
 If CboCurrency.Text = "USD" Then
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
      Sqlqry = " Insert into crdt_mas values('" & lblVoucNo & "','CNT','" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & Mid(lstDebitedTo, 1, 6) & "','" _
                                     & Mid(lstDebitedTo, 12, 35) & "','" _
                                     & Mid(lstCreditedTo, 1, 4) & "','" _
                                     & Trim(Mid(lstCreditedTo, 9, 35)) & "','" _
                                     & findfirstfixup(UCase(Trim(txtRef))) & "','" _
                                     & findfirstfixup(UCase(Trim(txtdesc))) & "','" _
                                     & findfirstfixup(UCase(Trim(txtDesc1))) & "'," _
                                     & Round(Val(txtCrntPer), 2) & ",'" _
                                     & Trim(CboCurrency) & "'," _
                                     & Val(txtConvRate) & "," _
                                     & Val(txtamount) & "," _
                                     & Val(txtamount) * Val(txtConvRate) & ",'N')"
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
    MTYPE = Trim(Mid(MTYPE, 1, 11))
   If MTYPE = "ZEINA" Or MTYPE = "ALAM ASSAYA" Then
   
        If CboCurrency.Text = "DHS" Then
         CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
         CrystalReport1.ReportFileName = App.Path & "\crntvouMPS.rpt"
         CrystalReport1.SelectionFormula = "{crdt_mas.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
         CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtamount)) & " Only" & "'"
         CrystalReport1.Formulas(1) = "curtype='" & CboCurrency.Text & "'"
         CrystalReport1.WindowState = crptMaximized
         CrystalReport1.Action = 1
        Else
         CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
         CrystalReport1.ReportFileName = App.Path & "\crntvouMPS.rpt"
         CrystalReport1.SelectionFormula = "{crdt_mas.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
         CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(txtamount)) & " Only" & "'"
         CrystalReport1.Formulas(1) = "curtype='" & CboCurrency.Text & "'"
         CrystalReport1.WindowState = crptMaximized
         CrystalReport1.Action = 1
        End If
        
    Else
         If CboCurrency.Text = "DHS" Then
         CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
         CrystalReport1.ReportFileName = App.Path & "\crntvou.rpt"
         CrystalReport1.SelectionFormula = "{crdt_mas.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
         CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtamount)) & " Only" & "'"
         CrystalReport1.Formulas(1) = "curtype='" & CboCurrency.Text & "'"
         CrystalReport1.WindowState = crptMaximized
         CrystalReport1.Action = 1
        Else
         CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
         CrystalReport1.ReportFileName = App.Path & "\crntvou.rpt"
         CrystalReport1.SelectionFormula = "{crdt_mas.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
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
    txtCrntPer.Text = 0
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
      lstCreditedTo.AddItem rs1!Supp_no & "  :  " & rs1!supp_name
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
     CboCurrency.ListIndex = -1
     txtConvRate.Text = ""
     txtRef.Text = ""
     txtdesc.Text = ""
     txtDesc1.Text = ""
     txtamount.Text = ""
     txtGross.Text = ""
     TxtAgCom.Text = ""
     txtAddDiscount.Text = ""
     txtAdddiscountper.Text = ""
     txtNet.Text = ""
     txtAgency.Text = ""
     txtproduct.Text = ""
     txtCrntPer.Text = 0
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
 If KeyAscii = 13 Then cmdAdd.SetFocus
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
