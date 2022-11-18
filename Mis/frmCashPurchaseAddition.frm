VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmCashPurchaseAddition 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Cash Purchase Addition"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cash Purchase Addition"
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
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   11655
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
         Left            =   4005
         TabIndex        =   3
         Top             =   1200
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
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
         Left            =   7125
         TabIndex        =   41
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox CboCustSupp 
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
         Height          =   360
         Left            =   1800
         TabIndex        =   11
         Top             =   3240
         Width           =   5175
      End
      Begin VB.ListBox lstSaleType 
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
         Height          =   1260
         Left            =   1800
         TabIndex        =   4
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Frame FraItem 
         BackColor       =   &H00FFFFC0&
         Height          =   855
         Left            =   240
         TabIndex        =   23
         Top             =   3720
         Width           =   11055
         Begin VB.TextBox txtcode 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   12
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtdesc 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   1320
            MaxLength       =   200
            TabIndex        =   13
            Top             =   480
            Width           =   4935
         End
         Begin VB.TextBox txtunit 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   6480
            MaxLength       =   10
            TabIndex        =   14
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtQuantity 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   7320
            MaxLength       =   10
            TabIndex        =   15
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtRate 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   8640
            TabIndex        =   17
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtAmount 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   9720
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label16 
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
            TabIndex        =   29
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Rate"
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
            Left            =   8760
            TabIndex        =   28
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Quantity"
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
            Left            =   7440
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Unit"
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
            Left            =   6720
            TabIndex        =   26
            Top             =   240
            Width           =   420
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
            Left            =   1560
            TabIndex        =   25
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Code"
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
            TabIndex        =   24
            Top             =   240
            Width           =   570
         End
      End
      Begin VB.TextBox txtbill1 
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
         Height          =   405
         Left            =   7080
         MaxLength       =   6
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtbill2 
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
         Height          =   405
         Left            =   8640
         MaxLength       =   6
         TabIndex        =   6
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtLpoNo 
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
         Height          =   405
         Left            =   7080
         MaxLength       =   20
         TabIndex        =   1
         Top             =   480
         Width           =   2895
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
         Picture         =   "frmCashPurchaseAddition.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   7320
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
         Picture         =   "frmCashPurchaseAddition.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   7320
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00FFFF80&
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
         Left            =   6360
         Picture         =   "frmCashPurchaseAddition.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7320
         Width           =   975
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
         Picture         =   "frmCashPurchaseAddition.frx":0986
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   7320
         Width           =   975
      End
      Begin VB.TextBox txtBill3 
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
         Height          =   405
         Left            =   10200
         MaxLength       =   6
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   1320
         Top             =   7560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1695
         Left            =   240
         TabIndex        =   30
         Top             =   4800
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2990
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
         BackColorFixed  =   -2147483635
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
         Left            =   4005
         TabIndex        =   0
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
      Begin PVMaskEditLib.PVMaskEdit txtbilldate1 
         Height          =   375
         Left            =   7080
         TabIndex        =   8
         Top             =   2640
         Width           =   1335
         _Version        =   65541
         _ExtentX        =   2355
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
      Begin PVMaskEditLib.PVMaskEdit txtbilldate2 
         Height          =   375
         Left            =   8640
         TabIndex        =   9
         Top             =   2640
         Width           =   1335
         _Version        =   65541
         _ExtentX        =   2355
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
      Begin PVMaskEditLib.PVMaskEdit txtbilldate3 
         Height          =   375
         Left            =   10200
         TabIndex        =   10
         Top             =   2640
         Width           =   1335
         _Version        =   65541
         _ExtentX        =   2355
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
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Total Amt"
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
         Left            =   2940
         TabIndex        =   44
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label Label6 
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
         TabIndex        =   43
         Top             =   1320
         Width           =   1410
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
         Left            =   5865
         TabIndex        =   42
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004000&
         BorderWidth     =   2
         X1              =   0
         X2              =   11640
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Bill Date"
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
         Left            =   5760
         TabIndex        =   40
         Top             =   2760
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Purchase Type"
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
         TabIndex        =   39
         Top             =   1800
         Width           =   1590
      End
      Begin VB.Label Label12 
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
         Left            =   8400
         TabIndex        =   38
         Top             =   6720
         Width           =   1380
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "LPO Number"
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
         Left            =   5760
         TabIndex        =   37
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Bill Number"
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
         Left            =   5760
         TabIndex        =   36
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblgrAmount 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   9960
         TabIndex        =   35
         Top             =   6600
         Width           =   1365
      End
      Begin VB.Label lblVoucNo 
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
         Left            =   1800
         TabIndex        =   34
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Party Name "
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
         TabIndex        =   33
         Top             =   3360
         Width           =   1530
      End
      Begin VB.Label Label2 
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
         Left            =   3360
         TabIndex        =   32
         Top             =   600
         Width           =   510
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
         TabIndex        =   31
         Top             =   600
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmCashPurchaseAddition"
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
Dim con As Currency
Dim accode As Long
Dim X
Dim y
Dim i
Dim ctype

Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtTtlAmount.SetFocus
End Sub

Private Sub cboCurrency_LostFocus()
  accode = 0
 If cboCurrency.Text = "USD" Then
     accode = 103002
     txtConvRate.Text = ""
     lblConvRate.Visible = True
     txtConvRate.Visible = True
     txtConvRate.TabIndex = 4
     
     
    Else
     accode = 103001
     txtConvRate.Text = 1
     lblConvRate.Visible = False
     txtConvRate.Visible = False
     txtConvRate.TabIndex = 22
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

Private Sub CmdPrint_Click()
X = 0
Dim addr As String
Dim city As String
Dim coun As String
Dim tele As String
Dim fax As String

 X = InputBox("Enter Voucher Number to Print : ", "Print", "100000")
 If X = "" Then Exit Sub
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = "Select * from capr_mas where vouc_no =" & X & ""
      Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
          MsgBox " Voucher Number not found"
          Exit Sub
          Else
           ctype = rs!tcurrency
          Sqlqry2 = "Select * from supp_Fin where Supp_no='" & rs!cust_no & "'"
          Set rs1 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
          If rs1.RecordCount <> 0 Then
           If IsNull(rs1!Address) = True Then
            addr = ""
           Else
            addr = rs1!Address
           End If
           
           If IsNull(rs1!city) = True Then
            city = ""
           Else
            city = rs1!city
           End If
         
           If IsNull(rs1!country) = True Then
            coun = ""
           Else
            coun = rs1!country
           End If
           
           If IsNull(rs1!telephone) = True Then
            tele = ""
           Else
            tele = rs1!telephone
           End If
           
           If IsNull(rs1!fax) = True Then
            fax = ""
           Else
            fax = rs1!fax
           End If
  If ctype = "DHS" Then
   
  CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
  CrystalReport1.ReportFileName = App.Path & "\cashpur.rpt"
  CrystalReport1.SelectionFormula = "{Capr_tra.Vouc_no}=" & X & ""
  CrystalReport1.Formulas(0) = "xxx1='" & Inwords(rs!tra_amount) & " Only" & "'"
  CrystalReport1.Formulas(1) = "Raddr='" & Trim(addr) & "'"
  CrystalReport1.Formulas(2) = "Rcity='" & Trim(city) & "'"
  CrystalReport1.Formulas(3) = "Rcoun='" & Trim(coun) & "'"
  CrystalReport1.Formulas(4) = "Rtele='" & Trim(tele) & "'"
  CrystalReport1.Formulas(5) = "Rfax='" & Trim(fax) & "'"
  CrystalReport1.Formulas(6) = "curtype='" & cboCurrency.Text & "'"
  CrystalReport1.WindowState = crptMaximized
  CrystalReport1.Action = 1
  
  Else
      
  CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
  CrystalReport1.ReportFileName = App.Path & "\cashpur.rpt"
  CrystalReport1.SelectionFormula = "{Capr_tra.Vouc_no}=" & X & ""
  CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(rs!tra_amount) & " Only" & "'"
  CrystalReport1.Formulas(1) = "Raddr='" & Trim(addr) & "'"
  CrystalReport1.Formulas(2) = "Rcity='" & Trim(city) & "'"
  CrystalReport1.Formulas(3) = "Rcoun='" & Trim(coun) & "'"
  CrystalReport1.Formulas(4) = "Rtele='" & Trim(tele) & "'"
  CrystalReport1.Formulas(5) = "Rfax='" & Trim(fax) & "'"
  CrystalReport1.Formulas(6) = "curtype='" & cboCurrency.Text & "'"
  CrystalReport1.WindowState = crptMaximized
  CrystalReport1.Action = 1
End If
  End If
  End If
End Sub

Private Sub CmdSave_Click()
     If Val(txtTtlAmount.Text) = Val(lblgrAmount.Caption) Then
   
         If ValidateData = True Then
            cur = ""
            con = 1
            accode = 0
       If cboCurrency.Text = "USD" Then
          cur = "USD"
          accode = 103002
          con = Val(Trim(txtConvRate.Text))
          totdhs = Round(Val(txtTtlAmount) * convertion, 2)
          totusd = Val(txtTtlAmount)
        Else
          cur = "DHS"
          con = 1
          accode = 103001
          totdhs = Val(txtTtlAmount)
          totusd = Round(Val(txtTtlAmount) / convertion, 2)
        End If
    
  
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         Sqlqry1 = "Select * from DUMCAPR1"
         Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
         If rs.RecordCount = 0 Then
           MsgBox " Transactions are not recorded"
           Exit Sub
         Else
           Set ws = DBEngine.Workspaces(0)
           Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           Sqlqry2 = "Insert into CAPR_MAS values('" & Val(lblVoucNo.Caption) & "','PCA','" _
                                & UCase(Mid(lstSaleType, 1, 6)) & "', '" _
                                & Mid(lstSaleType, 10, 25) & "','" _
                                & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                & findfirstfixup(Trim(txtLpoNo)) & "','" _
                                & Mid(CboCustSupp, 1, 4) & "','" _
                                & findfirstfixup(Mid(CboCustSupp, 8, 25)) & "','" _
                                & Trim(cboCurrency.Text) & "'," _
                                & con & ", " _
                                & Val(Trim(txtTtlAmount.Text)) & ", '" _
                                & Val(lblgrAmount.Caption) * con & "','" _
                                & Trim(txtbill1) & "','" _
                                & Trim(txtbill2) & "','" _
                                & Trim(txtBill3) & "',' " _
                                & Format(txtbilldate1.TextWithMask, "dd/mm/yyyy") & "','" _
                                & Format(txtbilldate2.TextWithMask, "dd/mm/yyyy") & "','" _
                                & Format(txtbilldate3.TextWithMask, "dd/mm/yyyy") & "'," & totdhs & ", " & totusd & ",'" & accode & "' ,'N')"
                              

           ws.BeginTrans
           db.Execute (Sqlqry2)
           ws.CommitTrans
         
         rs.MoveFirst
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         Sqlqry3 = "Insert into CAPR_TRA values('" & rs!vouc_no & "','" & Trim(rs!tDate) & "','" _
                                & rs!it_code & "','" _
                                & findfirstfixup(rs!it_desc) & "','" _
                                & rs!it_unit & "'," _
                                & rs!it_qty & "," _
                                & rs!it_Rate & ",'" _
                                & Trim(cboCurrency.Text) & "'," _
                                & Val(Trim(txtConvRate.Text)) & ", " _
                                & rs!tra_amount & "," _
                                & Val(txtConvRate.Text) * Val(rs!tra_amount) & ")"
                                
          ws.BeginTrans
          db.Execute (Sqlqry3)
          ws.CommitTrans
          rs.MoveNext
          Loop
     End If
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Update docu_mas set doc_no='" & lblVoucNo & "' where doc_type='PCA'"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     textclear
     lblVoucNo = lblVoucNo + 1
     MsgBox " Record is inserted", vbInformation, "Status"
   End If
 Else
 MsgBox "Total Amount is not equal to Entered Amount"
Exit Sub
End If
End Sub

Private Sub Form_Load()
 txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
 cboCurrency.AddItem "DHS"
 cboCurrency.AddItem "USD"
 lblConvRate.Visible = False
 txtConvRate.Visible = False
 accode = 0
 
 AutoIncrementVoucher
 PopulateAcctSuppCust
 PopulateSaleCode
 lblgrAmount.Caption = 0
 
 Flexitems
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
  Sqlqry = "Delete * from DUMCAPR1"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
End Sub

Private Sub AutoIncrementVoucher()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select DOC_no from DOCU_MAS WHERE DOC_TYPE='PCA'"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
If rs.RecordCount = 0 Then
   MsgBox "Document type 'PCA' not found"
   Exit Sub
Else
   lblVoucNo = Val(rs!doc_no) + 1
End If
End Sub

Private Sub PopulateAcctSuppCust()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry1 = "Select * from Supp_fin order by Supp_no"
Sqlqry2 = "Select * from agndtls order by agentname"
Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)

CboCustSupp.Clear

If rs1.RecordCount = 0 Then
    MsgBox "No Records found in the Supplier register"
Else
   rs1.MoveFirst
   Do Until rs1.EOF
      CboCustSupp.AddItem rs1!Supp_no & "  :  " & rs1!Supp_name
      rs1.MoveNext
   Loop
End If

'If rs2.RecordCount = 0 Then
'    MsgBox "No Records found in the Customer Register"
'Else
'   rs2.MoveFirst
'   Do Until rs2.EOF
'      CboCustSupp.AddItem "AGNC" & "  :  " & rs2!agentname
'      rs2.MoveNext
'   Loop
'End If

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
ElseIf lstSaleType.SelCount = 0 Then
  MsgBox "Select Purchase type", vbInformation, "Invalid Entry"
  lstSaleType.SetFocus
  Exit Function
ElseIf txtbill1.Text = "" Then
  MsgBox "Invalid Delivery Number", vbInformation, "Invalid Entry"
  txtbill1.SetFocus
  Exit Function
ElseIf IsDate(txtbilldate1.TextWithMask) = False Then
  MsgBox "Invalid Bill Date", vbInformation, "Invalid Entry"
  txtbilldate1.SetFocus
  Exit Function
ElseIf txtQuantity.Text = "" Or IsNumeric(txtQuantity) = False Then
  MsgBox "Invalid quantity", vbInformation, "Invalid Entry"
  txtQuantity.SetFocus
  Exit Function
ElseIf txtRate.Text = "" Or IsNumeric(txtRate) = False Then
  MsgBox "Invalid rate", vbInformation, "Invalid Entry"
  txtRate.SetFocus
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
    .ColWidth(0) = 1300
    .ColWidth(1) = 6300
    .ColWidth(2) = 700
    .ColWidth(3) = 1025
    .ColWidth(4) = 750
    .ColWidth(5) = 850
    
    .Col = 1
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Description"
    .Col = 2
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Unit"
    .Col = 3
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Quantity"
    .Col = 4
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Rate"
    .Col = 5
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Amount"
    .Row = 0
    .Col = 1
  End With
End Sub

Private Sub Msflexgrid1_dblclick()
 Dim i
 Dim j
 Dim X
 
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
        txtcode = .Text
        .Col = 1
        txtdesc = .Text
        .Col = 2
        txtunit = .Text
        .Col = 3
        txtQuantity = .Text
        .Col = 4
        txtRate = .Text
        .Col = 5
        txtAmount = .Text
        
        lblgrAmount.Caption = Val(lblgrAmount.Caption) - Val(txtAmount)
        
                
                
        .RemoveItem (j)
        
        Sqlqry1 = "Delete * from dumcapr1 where it_Code='" & txtcode & "' and it_desc ='" & txtdesc & "' and It_value =" & Val(txtAmount) & ""
        ws.BeginTrans
        db.Execute Sqlqry1
        ws.CommitTrans
        
       txtcode.SetFocus
     End With
    End If
   End If
  End If
End Sub

Private Sub cboCustSupp_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtcode.SetFocus
End Sub

Private Sub lstSaleType_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtbill1.SetFocus
End Sub

Private Sub txtBill3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtbilldate1.SetFocus
End Sub

Private Sub txtBillDate1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtbilldate2.SetFocus
End Sub

Private Sub txtbilldate1_LostFocus()
If IsDate(txtdate.TextWithMask) = False Then
   MsgBox "Invalid Date", vbInformation, "Invalid Entry"
   txtdate.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtBillDate2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtbilldate3.SetFocus
End Sub

Private Sub txtbilldate2_LostFocus()
If IsDate(txtdate.TextWithMask) = False Then
   MsgBox "Invalid Date", vbInformation, "Invalid Entry"
   txtdate.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtBillDate3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboCustSupp.SetFocus
End Sub

Private Sub txtbilldate3_LostFocus()
If IsDate(txtdate.TextWithMask) = False Then
   MsgBox "Invalid Date", vbInformation, "Invalid Entry"
   txtdate.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdesc.SetFocus
End Sub

Private Sub txtConvRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstSaleType.SetFocus
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtLpoNo.SetFocus
End Sub

Private Function textclear()
     txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
     lstSaleType.ListIndex = 0
     txtLpoNo.Text = ""
     CboCustSupp.Text = ""
     txtbill1.Text = ""
     txtbill2.Text = ""
     txtcode.Text = ""
     txtdesc.Text = ""
     txtunit.Text = ""
     txtQuantity.Text = ""
     txtbilldate1.TextWithMask = ""
     txtbilldate2.TextWithMask = ""
     txtbilldate3.TextWithMask = ""
     txtRate.Text = ""
     txtAmount.Text = ""
     lblgrAmount.Caption = "0.00"
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from DUMCAPR1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     MSFlexGrid1.Clear
     lstSaleType.SetFocus
End Function

Private Sub PopulateSaleCode()
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Select * from Acct_mas where acct_code>='" & 401001 & "' and acct_code<='" & 404000 & "' ORDER BY ACCT_CODE"
 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

lstSaleType.Clear

If rs.RecordCount = 0 Then
   MsgBox "No Records found in the Code range 401001 to 404000"
   Exit Sub
Else
   rs.MoveFirst
   Do Until rs.EOF
     lstSaleType.AddItem rs!acct_code & " : " & rs!acct_name
     rs.MoveNext
   Loop
End If

End Sub
Private Sub txtbill1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtbill2.SetFocus
End Sub

Private Sub txtbill2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtBill3.SetFocus
End Sub

Private Sub txtdate_LostFocus()
If IsDate(txtdate.TextWithMask) = False Then
   MsgBox "Invalid Date", vbInformation, "Invalid Entry"
   txtdate.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtunit.SetFocus
End Sub

Private Sub txtLpoNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboCurrency.SetFocus
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRate.SetFocus
txtAmount.Text = Val(txtQuantity) * Val(txtRate)
End Sub
Private Sub txtRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtcode.SetFocus
End Sub
Private Sub txtRate_LostFocus()
Dim con As Currency
 txtAmount.Text = Val(txtQuantity) * Val(txtRate)
 If ValidateData = True Then
    
    If cboCurrency = "USD" Then
      con = Val(txtConvRate)
    Else
      con = 1
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = "select * from DUMCAPR1"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Sqlqry = " Insert into DUMCAPR1 values('" & lblVoucNo & "','" _
                              & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                              & Trim(txtcode) & "','" _
                              & findfirstfixup(Trim(txtdesc)) & "','" _
                              & Trim(txtunit) & "', " _
                              & Trim(txtQuantity) & "," _
                              & Trim(txtRate) & ",'" _
                              & Trim(cboCurrency.Text) & "'," _
                              & con & ", " _
                              & Val(Trim(txtAmount)) & "," _
                              & con * Val(Trim(txtAmount)) & ")"

        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
            
        Sqlqry1 = "select * from DUMCAPR1"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
           MSFlexGrid1.Clear
           Exit Sub
        Else
           Flexitems
           rs.MoveFirst
           Do Until rs.EOF
             MSFlexGrid1.AddItem rs!it_code & Chr(9) & rs!it_desc & Chr(9) & rs!it_unit & Chr(9) & rs!it_qty & Chr(9) & rs!it_Rate & Chr(9) & rs!tra_amount
             rs.MoveNext
           Loop
        End If
        lblgrAmount.Caption = Val(txtAmount)
        If txtTtlAmount.Text = Val(lblgrAmount.Caption) Then
          cmdSave.SetFocus
        Else
          txtcode.SetFocus
        End If
               
   Else
       X = 0
       y = 0
       rs.MoveFirst
       Do Until rs.EOF
        X = X + rs!tra_amount
        rs.MoveNext
        lblgrAmount.Caption = X
       Loop
       lblgrAmount.Caption = X + Val(txtAmount)
              
       Sqlqry = " Insert into DUMCAPR1 values('" & lblVoucNo & "','" _
                              & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                              & Trim(txtcode) & "','" _
                              & findfirstfixup(Trim(txtdesc)) & "','" _
                              & Trim(txtunit) & "', " _
                              & Trim(txtQuantity) & "," _
                              & Trim(txtRate) & ",'" _
                              & Trim(cboCurrency.Text) & "'," _
                              & con & ", " _
                              & Val(Trim(txtAmount)) & "," _
                              & con * Val(Trim(txtAmount)) & ")"

        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        
        Sqlqry1 = "select * from DUMCAPR1"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
           If rs.RecordCount = 0 Then
              MSFlexGrid1.Clear
              Exit Sub
            Else
              Flexitems
              rs.MoveFirst
               Do Until rs.EOF
                MSFlexGrid1.AddItem rs!it_code & Chr(9) & rs!it_desc & Chr(9) & rs!it_unit & Chr(9) & rs!it_qty & Chr(9) & rs!it_Rate & Chr(9) & rs!tra_amount
                rs.MoveNext
               Loop
            End If
         End If
     End If
   If txtTtlAmount.Text = Val(lblgrAmount.Caption) Then
     cmdSave.SetFocus
   Else
     txtcode.SetFocus
   End If
End Sub

Private Sub txtTtlAmount_KeyPress(KeyAscii As Integer)
If cboCurrency.Text = "USD" Then
 If KeyAscii = 13 Then txtConvRate.SetFocus
Else
 If KeyAscii = 13 Then lstSaleType.SetFocus
End If

End Sub

Private Sub txtunit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtQuantity.SetFocus
End Sub
