VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "PVMASK.OCX"
Begin VB.Form frmBOorg 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Booking Order"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1560
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4560
      Picture         =   "frmBOorg.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "<<&Back<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6960
      Picture         =   "frmBOorg.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFC0&
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
      Height          =   660
      Left            =   5760
      Picture         =   "frmBOorg.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Add"
      DisabledPicture =   "frmBOorg.frx":0306
      DownPicture     =   "frmBOorg.frx":0838
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3360
      MaskColor       =   &H008080FF&
      Picture         =   "frmBOorg.frx":0D6A
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   7695
      Left            =   120
      TabIndex        =   56
      Top             =   120
      Width           =   11655
      Begin VB.ComboBox CboClient 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   315
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   166
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox CboAgency 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   315
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   165
         Top             =   1320
         Width           =   2655
      End
      Begin VB.ComboBox cboregion 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9360
         TabIndex        =   158
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtremarks 
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
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   2280
         Width           =   10095
      End
      Begin VB.TextBox txtConvRate 
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
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   9360
         TabIndex        =   45
         Top             =   810
         Width           =   735
      End
      Begin VB.ComboBox cboCurrency 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cboMediatype 
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
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtboref 
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
         Left            =   5520
         TabIndex        =   6
         Top             =   1815
         Width           =   2895
      End
      Begin VB.ComboBox cboProduct 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   2655
      End
      Begin VB.ComboBox cboyear 
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
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cbomonth 
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
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin PVMaskEditLib.PVMaskEdit txtdate 
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   1455
         _Version        =   65541
         _ExtentX        =   2566
         _ExtentY        =   556
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
      Begin VB.Frame Fracin 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H000000FF&
         Height          =   4935
         Left            =   120
         TabIndex        =   85
         Top             =   2640
         Width           =   11415
         Begin VB.TextBox txtDisccin 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4200
            MaxLength       =   2
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   4440
            Width           =   495
         End
         Begin VB.ComboBox cbolength 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   5280
            TabIndex        =   48
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox cbodaycin 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox cbosubmedia 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtbarterCin 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   3960
            Width           =   1215
         End
         Begin VB.TextBox txtfreecin 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   3960
            Width           =   1215
         End
         Begin VB.TextBox txtAddDiscountCin 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6960
            MaxLength       =   10
            TabIndex        =   55
            Top             =   4440
            Width           =   1215
         End
         Begin VB.TextBox txtGrAmountCin 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9840
            TabIndex        =   122
            TabStop         =   0   'False
            Top             =   3960
            Width           =   1335
         End
         Begin VB.TextBox txtNetAmountCin 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9840
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   4440
            Width           =   1335
         End
         Begin VB.ComboBox cboMatCin 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   7680
            TabIndex        =   50
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtDesccin 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            TabIndex        =   49
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox cbotypecin 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   9480
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtamountcin 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10440
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   720
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexcin 
            Height          =   2415
            Left            =   120
            TabIndex        =   86
            Top             =   1200
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   4260
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   -2147483624
            ForeColor       =   128
            BackColorBkg    =   16777215
            GridLines       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtComPerCin 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   4440
            Width           =   495
         End
         Begin PVMaskEditLib.PVMaskEdit txtCinDateFrom 
            Height          =   255
            Left            =   1800
            TabIndex        =   160
            Top             =   720
            Width           =   1095
            _Version        =   65541
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin PVMaskEditLib.PVMaskEdit txtCinDateTo 
            Height          =   255
            Left            =   3000
            TabIndex        =   161
            Top             =   720
            Width           =   1095
            _Version        =   65541
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin VB.Label Label73 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Date From"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   163
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label78 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Date To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   162
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label63 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Discount (%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   147
            Top             =   4440
            Width           =   1335
         End
         Begin VB.Label Label60 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Sub Media"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   143
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Total Barter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   142
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label Label58 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Total Free"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Agency Com. (%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   4440
            Width           =   1455
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   " Net Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8160
            TabIndex        =   125
            Top             =   4440
            Width           =   1455
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Gross Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8160
            TabIndex        =   124
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Additional Disc."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   123
            Top             =   4440
            Width           =   1815
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00004080&
            X1              =   0
            X2              =   11400
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Label Label28 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Days"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   92
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   6120
            TabIndex        =   91
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Sec."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   90
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9480
            TabIndex        =   89
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Material"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   88
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Amount "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10440
            TabIndex        =   87
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame FraMag 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H000000FF&
         Height          =   4935
         Left            =   120
         TabIndex        =   76
         Top             =   2640
         Width           =   11415
         Begin VB.TextBox txtComments 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5880
            TabIndex        =   13
            Top             =   600
            Width           =   1815
         End
         Begin VB.ComboBox cbospace 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   2280
            TabIndex        =   11
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtSurcharge 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7560
            MaxLength       =   10
            TabIndex        =   20
            Top             =   4320
            Width           =   975
         End
         Begin VB.TextBox txtdiscmag 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3480
            MaxLength       =   2
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   4320
            Width           =   495
         End
         Begin VB.TextBox txtbartermag 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   129
            TabStop         =   0   'False
            Top             =   3840
            Width           =   1215
         End
         Begin VB.TextBox txtfreemag 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   3840
            Width           =   1215
         End
         Begin VB.TextBox txtnetamountmag 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9840
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   4320
            Width           =   1335
         End
         Begin VB.TextBox txtGrAmountmag 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9840
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox txtadddiscountmag 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            MaxLength       =   10
            TabIndex        =   19
            Top             =   4320
            Width           =   975
         End
         Begin VB.TextBox txtcompermag 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   4320
            Width           =   495
         End
         Begin VB.TextBox txtissueno 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            MaxLength       =   5
            TabIndex        =   8
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtamountmag 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10440
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cbotypemag 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   9480
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtdescmag 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4080
            TabIndex        =   12
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox cbomatmag 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   7800
            TabIndex        =   14
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtPage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   10
            Top             =   600
            Width           =   375
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexmag 
            Height          =   2535
            Left            =   120
            TabIndex        =   102
            Top             =   1080
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   4471
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   -2147483624
            ForeColor       =   128
            BackColorBkg    =   16777215
            GridLines       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin PVMaskEditLib.PVMaskEdit txtissdate 
            Height          =   255
            Left            =   600
            TabIndex        =   9
            Top             =   600
            Width           =   1095
            _Version        =   65541
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin VB.Label Label74 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Comments"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   164
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label72 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   720
            TabIndex        =   159
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Space"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   155
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Surcharge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   154
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Label Label64 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Disc (%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   148
            Top             =   4320
            Width           =   975
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Total Barter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   130
            Top             =   3840
            Width           =   1335
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Total Free"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   128
            Top             =   3840
            Width           =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00004080&
            X1              =   0
            X2              =   11400
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Add. Disc."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   106
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Gross Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8160
            TabIndex        =   105
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   " Net Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8040
            TabIndex        =   104
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Agency Com. (%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   4320
            Width           =   1815
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Amount "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10080
            TabIndex        =   82
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Material"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7800
            TabIndex        =   81
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9480
            TabIndex        =   80
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Page"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   79
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label22 
            BackColor       =   &H00FFFFC0&
            Caption         =   "References"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   78
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label24 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Iss #"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraTV 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H000000FF&
         Height          =   4935
         Left            =   120
         TabIndex        =   66
         Top             =   2640
         Width           =   11415
         Begin VB.TextBox txtdiscTV 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   4320
            Width           =   495
         End
         Begin VB.TextBox txtbarterTv 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox txtfreeTv 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox txtcompertv 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   4320
            Width           =   495
         End
         Begin VB.TextBox txtadddiscounttv 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6960
            MaxLength       =   10
            TabIndex        =   34
            Top             =   4320
            Width           =   1215
         End
         Begin VB.TextBox txtgramountTV 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9840
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   3720
            Width           =   1335
         End
         Begin VB.TextBox txtnetamountTV 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9840
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   4320
            Width           =   1335
         End
         Begin VB.TextBox txtDaytv 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            MaxLength       =   5
            TabIndex        =   24
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtamounttv 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9960
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtRatetv 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8880
            TabIndex        =   30
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtSpots 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8160
            MaxLength       =   3
            TabIndex        =   29
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtdesctv 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            TabIndex        =   26
            Top             =   480
            Width           =   2535
         End
         Begin PVMaskEditLib.PVMaskEdit PVMaskTime 
            Height          =   255
            Left            =   840
            TabIndex        =   25
            Top             =   480
            Width           =   615
            _Version        =   65541
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            Text            =   ""
            Mask            =   "##:##"
            BackColor       =   16777215
            Alignment       =   1
         End
         Begin VB.ComboBox cboMattv 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   315
            Left            =   4200
            TabIndex        =   27
            Top             =   480
            Width           =   2535
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlextv 
            Height          =   2775
            Left            =   240
            TabIndex        =   75
            Top             =   840
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   4895
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   -2147483624
            ForeColor       =   128
            BackColorBkg    =   16777215
            GridLines       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.ComboBox cbotypetv 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   315
            Left            =   6840
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label65 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Discount (%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   149
            Top             =   4320
            Width           =   1455
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Total Barter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4680
            TabIndex        =   134
            Top             =   3720
            Width           =   1335
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Total Free"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   132
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Agency Com. (%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   4320
            Width           =   1815
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   " Net Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8520
            TabIndex        =   113
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Gross Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8280
            TabIndex        =   112
            Top             =   3720
            Width           =   1335
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Additional Disc."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   111
            Top             =   4320
            Width           =   1695
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00004080&
            X1              =   0
            X2              =   11400
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Amount "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9960
            TabIndex        =   74
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Material"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   72
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Type"
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
            Left            =   6960
            TabIndex        =   71
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Time"
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
            Left            =   840
            TabIndex        =   70
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   1560
            TabIndex        =   69
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Day"
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
            Left            =   240
            TabIndex        =   67
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Spots"
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
            Left            =   8040
            TabIndex        =   73
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Rate "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9000
            TabIndex        =   68
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Fraemp 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H000000FF&
         Height          =   4935
         Left            =   120
         TabIndex        =   83
         Top             =   2640
         Width           =   11415
         Begin MSFlexGridLib.MSFlexGrid MSFleext 
            Height          =   4335
            Left            =   120
            TabIndex        =   84
            Top             =   360
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   7646
            _Version        =   393216
            Rows            =   18
            Cols            =   10
            FixedCols       =   0
            BackColor       =   -2147483624
            ForeColor       =   8388736
            BackColorBkg    =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Fraol 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H000000FF&
         Height          =   4935
         Left            =   120
         TabIndex        =   93
         Top             =   2640
         Width           =   11415
         Begin VB.TextBox txtdiscOL 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   4440
            Width           =   495
         End
         Begin VB.TextBox txtbarterol 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   3960
            Width           =   1215
         End
         Begin VB.TextBox txtfreeol 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   3960
            Width           =   1215
         End
         Begin VB.TextBox txtcomperOL 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   4440
            Width           =   495
         End
         Begin VB.TextBox txtAddDiscountOL 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7200
            MaxLength       =   10
            TabIndex        =   43
            Top             =   4440
            Width           =   1095
         End
         Begin VB.TextBox txtGrAmountOL 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9840
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   3960
            Width           =   1335
         End
         Begin VB.TextBox txtNetAmountOL 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9840
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   4440
            Width           =   1335
         End
         Begin VB.TextBox txtrateol 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9120
            TabIndex        =   40
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtamountol 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10200
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cbotypeol 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   315
            Left            =   6960
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtdescol 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2760
            TabIndex        =   35
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox cbomatol 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   315
            Left            =   4560
            TabIndex        =   36
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txtimpression 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   8040
            TabIndex        =   38
            Top             =   600
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexol 
            Height          =   2655
            Left            =   240
            TabIndex        =   94
            Top             =   1080
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   4683
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   -2147483624
            ForeColor       =   4194432
            BackColorFixed  =   -2147483631
            BackColorBkg    =   16777215
            AllowBigSelection=   0   'False
            FocusRect       =   0
            HighLight       =   0
            GridLines       =   2
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
            Height          =   255
            Left            =   120
            TabIndex        =   152
            Top             =   600
            Width           =   1215
            _Version        =   65541
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin PVMaskEditLib.PVMaskEdit txtdateto 
            Height          =   255
            Left            =   1440
            TabIndex        =   153
            Top             =   600
            Width           =   1215
            _Version        =   65541
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin VB.Label Label67 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Date To"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   151
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Discount (%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   150
            Top             =   4440
            Width           =   1455
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Total Barter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   138
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Total Free"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   136
            Top             =   3960
            Width           =   1215
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Agency Com.(%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   4440
            Width           =   1815
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   " Net Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8160
            TabIndex        =   119
            Top             =   4440
            Width           =   1455
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Gross Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8160
            TabIndex        =   118
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Additional Disc."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   117
            Top             =   4440
            Width           =   1815
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00004080&
            X1              =   0
            X2              =   11400
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Label Label35 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Rate  / CPM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   101
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10200
            TabIndex        =   100
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6960
            TabIndex        =   98
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label31 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impressions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8040
            TabIndex        =   97
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   96
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label29 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Date from"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Material"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4800
            TabIndex        =   99
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label71 
         BackColor       =   &H00FFFFC0&
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
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   360
         TabIndex        =   157
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label70 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Region"
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
         Height          =   255
         Left            =   8520
         TabIndex        =   156
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label62 
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
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   840
         TabIndex        =   146
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblConvRate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Convertion"
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
         Left            =   8160
         TabIndex        =   145
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label61 
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
         ForeColor       =   &H00004000&
         Height          =   345
         Left            =   4560
         TabIndex        =   144
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Media Type"
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
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "B.O Ref #"
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
         Height          =   255
         Left            =   4440
         TabIndex        =   64
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblserialno 
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
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   9360
         TabIndex        =   63
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Serial #"
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
         Height          =   255
         Left            =   8280
         TabIndex        =   62
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Height          =   375
         Left            =   4200
         TabIndex        =   61
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Client"
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
         Height          =   255
         Left            =   8640
         TabIndex        =   60
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Height          =   255
         Left            =   360
         TabIndex        =   59
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Year"
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
         Height          =   255
         Left            =   360
         TabIndex        =   58
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Month"
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
         Height          =   255
         Left            =   4440
         TabIndex        =   57
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmBOorg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim dt As Long
Dim m As Long
Dim Y As Long
Dim mtype As Long
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim X, Z As Integer
Dim agdisc As Currency
Dim extdisc As Currency
Dim adddisc As Currency
Dim AddDiscEach As Currency
Dim Nettra As Currency
Dim NOS As Integer
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim ws As Workspace
Dim invdate As Date

Private Sub Flexitemsol()
  With MSFlexol
    .Clear
    .AllowUserResizing = flexResizeColumns
    .Rows = 1
    .Cols = 8
    .Col = 0
    .Row = 0
    .Text = "DT_From"
    .ColAlignment(0) = 0
    .ColWidth(0) = 900
    .ColWidth(1) = 900
    .ColWidth(2) = 3000
    .ColWidth(3) = 2400
    .ColWidth(4) = 800
    .ColWidth(5) = 900
    .ColWidth(6) = 900
    .ColWidth(7) = 1000
    
    .Col = 1
    .Text = "DT_To"
    .Col = 2
    .Text = "Description"
    .Col = 3
    .Text = "Material"
    .Col = 4
    .Text = "P_Type"
    .Col = 5
    .Text = "Impres."
    .Col = 6
    .Text = "Rate"
    .Col = 7
    .Text = "Amount"
    
    .Row = 0
    .Col = 1
  
  End With
End Sub
Private Sub CboAgency_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CboClient.SetFocus
End Sub
Private Sub CboClient_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboMediatype.SetFocus
End Sub

Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If cboCurrency.Text = "USD" Then
    lblConvRate.Visible = True
    txtConvRate.Visible = True
    txtConvRate.Text = ""
    txtConvRate.SetFocus
 Else
     lblConvRate.Visible = False
     txtConvRate.Visible = False
     txtConvRate.Text = 1
     cboProduct.SetFocus
 End If
 End If
End Sub

Private Sub cboCurrency_LostFocus()
 If cboCurrency.Text = "USD" Then
     lblConvRate.Visible = True
     txtConvRate.Visible = True
     txtConvRate.Text = ""
     txtConvRate.TabIndex = 4
    Else
     lblConvRate.Visible = False
     txtConvRate.Visible = False
     txtConvRate.Text = 1
     txtConvRate.TabIndex = 38
 End If
End Sub

Private Sub cbolength_LostFocus()
 Dim C As String
 Dim X As Integer
 C = ""
 X = 0
 
 If Val(Trim(cbolength.Text)) = 10 Or Val(Trim(cbolength.Text)) = 30 Or Val(Trim(cbolength.Text)) = 60 Or Val(Trim(cbolength.Text)) = 90 Then
 
    If Mid(cbodaycin, 1, 1) = "B" Then
       C = "BIW" & Val(Trim(cbolength.Text))
    Else
       C = "MON" & Val(Trim(cbolength.Text))
    End If
    
 
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select " & C & " from cinema_rates where sub_media='" & Trim(cbosubmedia.Text) & "'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If IsNull(rs.Fields(0)) = False Then X = rs.Fields(0)
     txtamountcin.Text = X
  Else
   X = 0
   txtamountcin.Text = ""
  End If
           
End Sub
Private Sub cboMatCin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cbotypecin.SetFocus
End Sub
Private Sub cbomatmag_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbotypemag.SetFocus
End Sub
Private Sub cbomatol_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbotypeol.SetFocus
End Sub
Private Sub cboMattv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cbotypetv.SetFocus
End Sub
Private Sub cboMediatype_Click()
If Mid(cboMediatype.Text, 1, 3) = "Tel" Then
   fraTV.Visible = True
   Fraol.Visible = False
   Fracin.Visible = False
   FraMag.Visible = False
   Fraemp.Visible = False
   mtype = 1
   txtboref.SetFocus
      ' txtDaytv.SetFocus
   Flexitemstv
 ElseIf Mid(cboMediatype.Text, 1, 3) = "Onl" Then
   fraTV.Visible = False
   Fraol.Visible = True
   Fracin.Visible = False
   FraMag.Visible = False
   Fraemp.Visible = False
   mtype = 2
   txtboref.SetFocus
   
   'txtdatefrom.SetFocus
   Flexitemsol
 ElseIf Mid(cboMediatype.Text, 1, 3) = "Cin" Then
   fraTV.Visible = False
   Fraol.Visible = False
   Fracin.Visible = True
   FraMag.Visible = False
   Fraemp.Visible = False
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from cinema_rates order by sub_media"
     'Sqlqry = "Select * from media where media='Cinema' "
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
        rs.MoveFirst
        cbosubmedia.Clear
        Do Until rs.EOF
         cbosubmedia.AddItem rs!sub_media
         rs.MoveNext
        Loop
    End If
   
   cbodaycin.Clear
   cbodaycin.AddItem "Bi-Weekly"
   cbodaycin.AddItem "Monthly"
   
   cbolength.Clear
   cbolength.AddItem "10"
   cbolength.AddItem "30"
   cbolength.AddItem "60"
   cbolength.AddItem "90"
   cbolength.AddItem "OTHER"
   mtype = 3
   txtboref.SetFocus
   
   Flexitemscin
 ElseIf Mid(cboMediatype.Text, 1, 3) = "Mag" Then
   fraTV.Visible = False
   Fraol.Visible = False
   Fracin.Visible = False
   FraMag.Visible = True
   Fraemp.Visible = False
   mtype = 4
   txtboref.SetFocus
   Flexitemsmag
Else
   fraTV.Visible = False
   Fraol.Visible = False
   Fracin.Visible = False
   FraMag.Visible = False
   Fraemp.Visible = True
   txtboref.SetFocus
 End If
  
End Sub

Private Sub cbomonth_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdate.SetFocus
End Sub

Private Sub cbomonth_LostFocus()

 X = cbomonth.Text
 
 
If X = "January" Then
    dt = 31
ElseIf X = "February" Then
    dt = 28
ElseIf X = "March" Then
    dt = 31
ElseIf X = "April" Then
    dt = 30
ElseIf X = "May" Then
    dt = 31
ElseIf X = "June" Then
    dt = 30
ElseIf X = "July" Then
    dt = 31
ElseIf X = "August" Then
    dt = 31
ElseIf X = "September" Then
    dt = 30
ElseIf X = "October" Then
    dt = 31
ElseIf X = "November" Then
    dt = 30
Else
    dt = 31
End If

m = cbomonth.ListIndex + 1
Y = cboyear.Text
invdate = dt & " / " & m & " / " & Y
invdate = Format(invdate, "dd/mm/yyyy")

End Sub

Private Sub cboProduct_Click()

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from products where Product_Name='" & Trim(cboProduct.Text) & "'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount <> 0 Then
        
        txtcompermag.Text = Val(rs!Discount)
        txtComPerCin.Text = Val(rs!Discount)
        txtcompertv.Text = Val(rs!Discount)
        txtcomperOL.Text = Val(rs!Discount)
     
    End If
   cboMattv.Clear
   cbomatmag.Clear
   cboMatCin.Clear
   cbomatol.Clear
             
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = "Select * from material where Product='" & Trim(cboProduct.Text) & "'"
    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If rs1.RecordCount <> 0 Then
        
            cboMattv.Clear
            cbomatmag.Clear
            cboMatCin.Clear
            cbomatol.Clear
            
                rs1.MoveFirst
                
                Do Until rs1.EOF
                 cboMattv.AddItem rs1!Name
                 cboMatCin.AddItem rs1!Name
                 cbomatmag.AddItem rs1!Name
                 cbomatol.AddItem rs1!Name
                 rs1.MoveNext
                Loop
                CboAgency.SetFocus
        End If
 End Sub

Private Sub cboProduct_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboAgency.SetFocus
End Sub

Private Sub cbospace_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdescmag.SetFocus
End Sub

Private Sub cbosubmedia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCinDateFrom.SetFocus
End Sub

Private Sub cbotypecin_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtamountcin.SetFocus
End Sub

Private Sub cbotypemag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtamountmag.SetFocus
End Sub
Private Sub cbotypeol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtimpression.SetFocus
End Sub
Private Sub cbotypetv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtSpots.SetFocus
End Sub
Private Sub cboyear_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbomonth.SetFocus
End Sub
Private Sub cboyear_LostFocus()
    Y = cboyear.Text
    m = cbomonth.ListIndex + 1
 If cbomonth.ListIndex = 0 Then
   dt = 31
 ElseIf cbomonth.ListIndex = 1 Then
   dt = 28
 ElseIf cbomonth.ListIndex = 2 Then
   dt = 31
 ElseIf cbomonth.ListIndex = 3 Then
   dt = 30
 ElseIf cbomonth.ListIndex = 4 Then
   dt = 31
 ElseIf cbomonth.ListIndex = 5 Then
   dt = 30
 ElseIf cbomonth.ListIndex = 6 Then
   dt = 31
 ElseIf cbomonth.ListIndex = 7 Then
   dt = 31
 ElseIf cbomonth.ListIndex = 8 Then
   dt = 30
 ElseIf cbomonth.ListIndex = 9 Then
   dt = 31
 ElseIf cbomonth.ListIndex = 10 Then
   dt = 30
 Else
   dt = 31
 End If
 
Y = cboyear.Text
m = cbomonth.ListIndex + 1
invdate = dt & " / " & m & " / " & Y
invdate = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub cmdadd_Click()
  
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from Bo_TRAmagprn"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "delete * from Bo_TRAtvprn"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "delete * from Bo_TRAcinprn"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from Bo_TRAolprn"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
 
  
 
 If ValidateData = True Then
    If cboCurrency.Text = "DHS" Then
      txtConvRate.Text = 1
    End If
                  
                  
                  
  If Mid(Trim(cboMediatype.Text), 1, 3) = "Onl" Then
    If MSFlexol.Rows > 1 Then
   
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Insert into Bo_Mas values('" & Val(lblserialno.Caption) & "','" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" & Trim(cboCurrency.Text) & "'," & Val(txtConvRate) & "," & Val(Trim(txtGrAmountOL.Text)) & "," & Val(Trim(txtNetAmountOL.Text)) & ",'" & cboyear.Text & "','" _
                                     & Trim(cbomonth.Text) & " '," & Val(cbomonth.ListIndex) & ",'" _
                                     & findfirstfixup(Trim(cboregion.Text)) & "','" & findfirstfixup(Trim(txtremarks.Text)) & "','" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Online','" _
                                     & Trim(Mid(cboMediatype, 9, 30)) & "','" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "'," _
                                     & Val(txtGrAmountOL.Text) * Val(txtConvRate) & "," _
                                     & Val(Trim(txtfreeol.Text)) & "," _
                                     & Val(Trim(txtbarterol.Text)) & ",'" _
                                     & Val(Trim(txtcomperOL.Text)) & "','" _
                                     & Val(Trim(txtdiscOL.Text)) & "'," _
                                     & Val(Trim(txtAddDiscountOL.Text)) & "," & 0 & "," _
                                     & Val(Trim(txtNetAmountOL.Text)) * Val(txtConvRate) & ",'" & Format(invdate, "dd/mm/yyyy") & "','301000','N','N')"
       ws.BeginTrans
       db.Execute (Sqlqry)
       ws.CommitTrans
      
    Sqlqry1 = "Select * from dumBo_TRAolbo"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount = 0 Then
         MsgBox " Transactions are not recorded"
         Exit Sub
      Else
         rs.MoveFirst
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           Sqlqry2 = " Insert into Bo_traol values('" & rs!serial_no & "','" & Trim(rs!date_From) & "','" & Trim(rs!date_to) & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "','" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                     & Trim(rs!sub_media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Type) & "'," _
                                     & Val(Trim(rs!impressions)) & "," _
                                     & Val(Trim(rs!Rate)) & ",'" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!amount) & ")"
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
             
             Sqlqry2 = " Insert into Bo_traolprn values('" & rs!serial_no & "','" & Trim(rs!date_From) & "','" & Trim(rs!date_to) & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "','" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                     & Trim(rs!sub_media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Type) & "'," _
                                     & Val(Trim(rs!impressions)) & "," _
                                     & Val(Trim(rs!Rate)) & ",'" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!amount) & ")"
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
            
          rs.MoveNext
         Loop
       End If
        
         lblserialno.Caption = Val(lblserialno.Caption) + 1
        
         MsgBox " Record is inserted", vbInformation, "Status"
        
         X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
        If X = vbYes Then
         With CrystalReport1
          .DataFiles(0) = App.Path & "\misov.mdb"
          .ReportFileName = App.Path & "\bool.rpt"
          .WindowState = crptMaximized
          .Action = 1
         End With
        End If
      Else
       MsgBox "No Transactions are recorded"
       Exit Sub
   End If
     
     
  ElseIf Mid(Trim(cboMediatype.Text), 1, 3) = "Tel" Then
   If MSFlextv.Rows > 1 Then
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Insert into Bo_Mas values('" & Val(lblserialno.Caption) & "','" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" & Trim(cboCurrency.Text) & "'," & Val(txtConvRate) & "," & Val(Trim(txtgramountTV.Text)) & "," & Val(Trim(txtnetamountTV.Text)) & ",'" & cboyear.Text & "','" _
                                     & Trim(cbomonth.Text) & " '," & Val(cbomonth.ListIndex) & ",'" _
                                     & findfirstfixup(Trim(cboregion.Text)) & "','" & findfirstfixup(Trim(txtremarks.Text)) & "','" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Television','" _
                                     & Trim(Mid(cboMediatype, 13, 30)) & "','" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "'," _
                                     & Val(txtgramountTV.Text) * Val(txtConvRate) & "," _
                                     & Val(Trim(txtfreeTv.Text)) & "," _
                                     & Val(Trim(txtbarterTv.Text)) & ",'" _
                                     & Val(Trim(txtdiscTV.Text)) & "','" _
                                     & Val(Trim(txtcompertv.Text)) & "'," _
                                     & Val(Trim(txtadddiscounttv.Text)) & "," & 0 & ", " _
                                     & Val(Trim(txtnetamountTV.Text)) * Val(txtConvRate) & ",'" & Format(invdate, "dd/mm/yyyy") & "','301000','N','N')"
       ws.BeginTrans
       db.Execute (Sqlqry)
       ws.CommitTrans
    
    
    Sqlqry1 = "Select * from dumBo_TRAtvbo"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount = 0 Then
         MsgBox " Transactions are not recorded"
         Exit Sub
      Else
         rs.MoveFirst
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           Sqlqry2 = " Insert into bo_tratv values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "','" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                     & Trim(rs!sub_media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!Day) & "','" _
                                     & Trim(rs!Time) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Type) & "'," _
                                     & Val(Trim(rs!spots)) & "," _
                                     & Val(Trim(rs!Rate)) & ",'" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!amount) & ")"
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
            
            Sqlqry2 = " Insert into bo_tratvprn values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "','" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                     & Trim(rs!sub_media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!Day) & "','" _
                                     & Trim(rs!Time) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Type) & "'," _
                                     & Val(Trim(rs!spots)) & "," _
                                     & Val(Trim(rs!Rate)) & ",'" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!amount) & ")"
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
          rs.MoveNext
         Loop
       End If
        
        lblserialno.Caption = Val(lblserialno.Caption) + 1
        MsgBox " Record is inserted", vbInformation, "Status"
        
        
        X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
        
        If X = vbYes Then
                With CrystalReport1
                    .DataFiles(0) = App.Path & "\misov.mdb"
                    .ReportFileName = App.Path & "\BoTv.rpt"
                    .WindowState = crptMaximized
                    .Action = 1
                End With
         
        End If
        
     Else
      MsgBox "No Transactions are recorded"
      Exit Sub
     End If
     
  ElseIf Mid(Trim(cboMediatype.Text), 1, 3) = "Cin" Then
   If MSFlexcin.Rows > 1 Then
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = " Insert into Bo_Mas values('" & Val(lblserialno.Caption) & "','" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" & Trim(cboCurrency.Text) & "'," & Val(txtConvRate) & "," & Val(Trim(txtGrAmountCin.Text)) & "," & Val(Trim(txtNetAmountCin.Text)) & ",'" & cboyear.Text & "','" _
                                             & Trim(cbomonth.Text) & " '," & Val(cbomonth.ListIndex) & ",'" _
                                             & findfirstfixup(Trim(cboregion.Text)) & "','" & findfirstfixup(Trim(txtremarks.Text)) & "','" _
                                             & findfirstfixup(cboProduct.Text) & "','" _
                                             & findfirstfixup(CboClient) & "','" _
                                             & findfirstfixup(CboAgency) & "','Cinema','Cinema','" _
                                             & findfirstfixup(Trim(txtboref.Text)) & "'," _
                                             & Val(txtGrAmountCin.Text) * Val(txtConvRate) & "," _
                                             & Val(Trim(txtfreecin.Text)) & "," _
                                             & Val(Trim(txtbarterCin.Text)) & ",'" _
                                             & Val(Trim(txtDisccin.Text)) & "','" _
                                             & Val(Trim(txtComPerCin.Text)) & "'," _
                                             & Val(Trim(txtAddDiscountCin.Text)) & "," & 0 & "," _
                                             & Val(Trim(txtNetAmountCin.Text)) * Val(txtConvRate) & ",'" & Format(invdate, "dd/mm/yyyy") & "','301000','N','N')"
               ws.BeginTrans
               db.Execute (Sqlqry)
               ws.CommitTrans
       
      agdisc = 0
      extdisc = 0
      adddisc = 0
      NOS = 0
      AddDiscEach = 0
       
      agdisc = Val(txtComPerCin.Text)
      extdisc = Val(txtDisccin.Text)
      adddisc = Val(txtAddDiscountCin.Text)
      
      
    Sqlqry1 = " select * from Dumbo_tracinbo where type ='Paid'"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    
    
    rs.MoveFirst
    Do Until rs.EOF
      NOS = rs.RecordCount
     rs.MoveNext
    Loop
    
    AddDiscEach = Round(adddisc / NOS, 3)
     
     
    Sqlqry1 = "Select * from dumBo_TRAcinbo"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount <> 0 Then
         rs.MoveFirst
         Nettra = 0
         Do Until rs.EOF
          If rs!Type <> "Paid" Then
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                          Sqlqry2 = " Insert into bo_tracin values('" & rs!serial_no & "','" & rs!Year & "','" _
                                                    & Trim(rs!Month) & "','" _
                                                    & findfirstfixup(rs!Product) & "','" _
                                                    & findfirstfixup(rs!client) & "','" _
                                                    & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                                    & Trim(rs!sub_media) & "','" _
                                                    & Trim(rs!datefrom) & "','" _
                                                    & Trim(rs!dateto) & "','" _
                                                    & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                                    & Trim(rs!Day) & "','" _
                                                    & Trim(rs!Length) & "','" _
                                                    & findfirstfixup(Trim(rs!Description)) & "','" _
                                                    & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                                    & Trim(rs!Type) & "','" _
                                                    & Trim(rs!tcurrency) & "'," _
                                                    & Trim(rs!tconvertion) & "," _
                                                    & Trim(rs!tra_amount) & "," _
                                                    & Trim(rs!amount) & "," & 0 & ")"
                           ws.BeginTrans
                           db.Execute (Sqlqry2)
                           ws.CommitTrans
                           
                           
                    Sqlqry2 = " Insert into bo_tracinprn values('" & rs!serial_no & "','" & rs!Year & "','" _
                                             & Trim(rs!Month) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                             & Trim(rs!sub_media) & "','" _
                                             & Trim(rs!datefrom) & "','" _
                                             & Trim(rs!dateto) & "','" _
                                             & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                             & Trim(rs!Day) & "','" _
                                             & Trim(rs!Length) & "','" _
                                             & findfirstfixup(Trim(rs!Description)) & "','" _
                                             & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                             & Trim(rs!Type) & "','" _
                                             & Trim(rs!tcurrency) & "'," _
                                             & Trim(rs!tconvertion) & "," _
                                             & Trim(rs!tra_amount) & "," _
                                             & Trim(rs!amount) & ")"
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
          Else
             Nettra = Val(rs!tra_amount) - (Val(rs!tra_amount) * agdisc / 100) - ((Val(rs!tra_amount) - Val(rs!tra_amount) * agdisc / 100) * extdisc / 100) - AddDiscEach
       
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                          Sqlqry2 = " Insert into bo_tracin values('" & rs!serial_no & "','" & rs!Year & "','" _
                                                    & Trim(rs!Month) & "','" _
                                                    & findfirstfixup(rs!Product) & "','" _
                                                    & findfirstfixup(rs!client) & "','" _
                                                    & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                                    & Trim(rs!sub_media) & "','" _
                                                    & Trim(rs!datefrom) & "','" _
                                                    & Trim(rs!dateto) & "','" _
                                                    & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                                    & Trim(rs!Day) & "','" _
                                                    & Trim(rs!Length) & "','" _
                                                    & findfirstfixup(Trim(rs!Description)) & "','" _
                                                    & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                                    & Trim(rs!Type) & "','" _
                                                    & Trim(rs!tcurrency) & "'," _
                                                    & Trim(rs!tconvertion) & "," _
                                                    & Trim(rs!tra_amount) & "," _
                                                    & Trim(rs!amount) & "," & Nettra & ")"
                           ws.BeginTrans
                           db.Execute (Sqlqry2)
                           ws.CommitTrans
                           
                           
                    Sqlqry2 = " Insert into bo_tracinprn values('" & rs!serial_no & "','" & rs!Year & "','" _
                                             & Trim(rs!Month) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                             & Trim(rs!sub_media) & "','" _
                                             & Trim(rs!datefrom) & "','" _
                                             & Trim(rs!dateto) & "','" _
                                             & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                             & Trim(rs!Day) & "','" _
                                             & Trim(rs!Length) & "','" _
                                             & findfirstfixup(Trim(rs!Description)) & "','" _
                                             & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                             & Trim(rs!Type) & "','" _
                                             & Trim(rs!tcurrency) & "'," _
                                             & Trim(rs!tconvertion) & "," _
                                             & Trim(rs!tra_amount) & "," _
                                             & Trim(rs!amount) & ")"
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
            End If
            
            
            
          rs.MoveNext
         Loop
       End If
        
        lblserialno.Caption = Val(lblserialno.Caption) + 1
        MsgBox " Record is inserted", vbInformation, "Status"
        X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
        If X = vbYes Then
         With CrystalReport1
          .DataFiles(0) = App.Path & "\misov.mdb"
          .ReportFileName = App.Path & "\bocin.rpt"
          .WindowState = crptMaximized
          .Action = 1
         End With
        End If
     
     Else
      MsgBox "No Transactions are recorded"
      Exit Sub
     End If
    
 Else
  If MSFlexmag.Rows > 1 Then
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Insert into Bo_Mas values('" & Val(lblserialno.Caption) & "','" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" & Trim(cboCurrency.Text) & "'," & Val(txtConvRate) & "," & Val(Trim(txtGrAmountmag.Text)) & "," & Val(Trim(txtnetamountmag.Text)) & ",'" & cboyear.Text & "','" _
                                     & Trim(cbomonth.Text) & " '," & Val(cbomonth.ListIndex) & ",'" _
                                     & findfirstfixup(Trim(cboregion.Text)) & "','" & findfirstfixup(Trim(txtremarks.Text)) & "','" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Magazine','" _
                                     & Trim(Mid(cboMediatype, 11, 30)) & "','" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "'," _
                                     & Val(txtGrAmountmag.Text) * Val(txtConvRate) & "," _
                                     & Val(Trim(txtfreemag.Text)) & "," _
                                     & Val(Trim(txtbartermag.Text)) & ",'" _
                                     & Val(Trim(txtdiscmag.Text)) & "','" _
                                     & Val(Trim(txtcompermag.Text)) & "'," _
                                     & Val(Trim(txtadddiscountmag.Text)) & "," & Val(Trim(txtSurcharge)) & "," _
                                     & Val(Trim(txtnetamountmag.Text)) * Val(txtConvRate) & ",'" & Format(invdate, "dd/mm/yyyy") & "','301000','N','N')"
       ws.BeginTrans
       db.Execute (Sqlqry)
       ws.CommitTrans
    
    Sqlqry1 = "Select * from dumBo_TRAmagbo"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount = 0 Then
         MsgBox " Transactions are not recorded"
         Exit Sub
      Else
         rs.MoveFirst
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           Sqlqry2 = " Insert into bo_tramag values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(cbomonth.ListIndex) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                     & Trim(rs!sub_media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tdate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!amount) & ")"
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
            
            
            Sqlqry2 = " Insert into bo_tramagprn values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(cbomonth.ListIndex) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                     & Trim(rs!sub_media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tdate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!amount) & ")"
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
          rs.MoveNext
         Loop
       End If
   lblserialno.Caption = Val(lblserialno.Caption) + 1
  
   MsgBox " Record is inserted", vbInformation, "Status"
  
    X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
    If X = vbYes Then
     With CrystalReport1
      .DataFiles(0) = App.Path & "\misov.mdb"
      .ReportFileName = App.Path & "\bomag.rpt"
      .WindowState = crptMaximized
      .Action = 1
     End With
    End If
 
  Else
    MsgBox "No Transactions are recorded"
    Exit Sub
  End If
  End If
  textclear
 End If

End Sub

Private Sub CmdBack_Click()
 Unload Me
End Sub
 Private Sub textclear()

   cboProduct.ListIndex = -1
   CboAgency.ListIndex = -1
   CboClient.ListIndex = -1
   txtboref.Text = ""
   cboMediatype.ListIndex = -1
   txtdescol.Text = ""
   cbomatol.ListIndex = -1
   cbotypeol.ListIndex = -1
   txtimpression.Text = ""
   txtrateol.Text = ""
   txtamountol.Text = ""
   
   txtDaytv.Text = ""
   PVMaskTime.TextWithMask = ""
   txtdesctv.Text = ""
   cboMattv.ListIndex = -1
   cbotypetv.ListIndex = -1
   txtSpots.Text = ""
   txtRatetv.Text = ""
   txtamounttv.Text = ""

   cbodaycin.ListIndex = -1
   cbolength.ListIndex = -1
   txtCinDateFrom.TextWithMask = Format(Now(), "dd/mm/yyyy")
   txtCinDateTo.TextWithMask = Format(Now(), "dd/mm/yyyy")
   txtDesccin.Text = ""
   cbosubmedia.ListIndex = -1
   cboMatCin.ListIndex = -1
   cbotypecin.ListIndex = -1
   txtamountcin.Text = ""
        
   txtissueno.Text = ""
   txtPage.Text = ""
   txtdescmag.Text = ""
   txtComments.Text = ""
   cbomatmag.ListIndex = -1
   cbotypemag.ListIndex = -1
   txtamountmag.Text = ""
     
     txtGrAmountmag.Text = ""
     txtGrAmountCin.Text = ""
     txtgramountTV.Text = ""
     txtGrAmountOL.Text = ""
     
     txtremarks.Text = ""
     cboregion.Text = ""
     
     txtnetamountmag.Text = ""
     txtNetAmountCin.Text = ""
     txtnetamountTV.Text = ""
     txtNetAmountOL.Text = ""
     
     txtadddiscountmag.Text = ""
     txtAddDiscountCin.Text = ""
     txtadddiscounttv.Text = ""
     txtAddDiscountOL.Text = ""
     
     txtcompermag.Text = ""
     txtComPerCin.Text = ""
     txtcompertv.Text = ""
     txtcomperOL.Text = ""
     
     txtdiscmag.Text = ""
     txtDisccin.Text = ""
     txtdiscTV.Text = ""
     txtdiscOL.Text = ""
     
     txtSurcharge.Text = ""
     
     cboCurrency.ListIndex = -1
     lblConvRate.Visible = False
     txtConvRate.Text = ""
     txtConvRate.Visible = False
     
     txtfreemag.Text = ""
     txtfreecin.Text = ""
     txtfreeTv.Text = ""
     txtfreeol.Text = ""
     
     txtbartermag.Text = ""
     txtbarterCin.Text = ""
     txtbarterTv.Text = ""
     txtbarterol.Text = ""
          
     Flexitemstv
     Flexitemsol
     Flexitemscin
     Flexitemsmag
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from dumBo_TRAolbo"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from dumBo_TRAtvbo"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from dumBo_TRAcinbo"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from dumBo_TRAmagbo"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
          
  End Sub
  
Private Sub cmdClear_Click()
  textclear
End Sub

Private Sub Form_Load()
U = 0
Y = 0
Z = 0

mtype = 0

cbospace.AddItem "FPC"
cbospace.AddItem "OBC"
cbospace.AddItem "IFC"
cbospace.AddItem "DPS"
cbospace.AddItem "SPO"
cbospace.AddItem "OTH"

lblConvRate.Visible = False
txtConvRate.Visible = False

cboCurrency.AddItem "DHS"
cboCurrency.AddItem "USD"

cbomonth.AddItem "January"
cbomonth.AddItem "February"
cbomonth.AddItem "March"
cbomonth.AddItem "April"
cbomonth.AddItem "May"
cbomonth.AddItem "June"
cbomonth.AddItem "July"
cbomonth.AddItem "August"
cbomonth.AddItem "September"
cbomonth.AddItem "October"
cbomonth.AddItem "November"
cbomonth.AddItem "December"


i = 2000
dt = 28
For i = 2000 To 2100
 cboyear.AddItem i
Next
X = 0

 cboyear.Text = Year(Now())
 
 X = Month(Now())
 
 
If X = 1 Then
   cbomonth.ListIndex = 0
   dt = 31
ElseIf X = 2 Then
   cbomonth.ListIndex = 1
   dt = 28
ElseIf X = 3 Then
   cbomonth.ListIndex = 2
   dt = 31
ElseIf X = 4 Then
   cbomonth.ListIndex = 3
   dt = 30
ElseIf X = 5 Then
   cbomonth.ListIndex = 4
   dt = 31
ElseIf X = 6 Then
   cbomonth.ListIndex = 5
   dt = 30
ElseIf X = 7 Then
   cbomonth.ListIndex = 6
   dt = 31
ElseIf X = 8 Then
   cbomonth.ListIndex = 7
   dt = 31
ElseIf X = 9 Then
   cbomonth.ListIndex = 8
   dt = 30
ElseIf X = 10 Then
   cbomonth.ListIndex = 9
   dt = 31
ElseIf X = 11 Then
   cbomonth.ListIndex = 10
   dt = 30
Else
   cbomonth.ListIndex = 11
   dt = 31
End If

txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")

populateproducts
populateagency
populateclient
populatemedia
Populateregion
AutoIncrementVoucher

   fraTV.Visible = False
   Fraol.Visible = False
   Fracin.Visible = False
   FraMag.Visible = False
   Fraemp.Visible = True


cbotypetv.AddItem "Paid"
cbotypetv.AddItem "Free"
cbotypetv.AddItem "Barter"

cbotypecin.AddItem "Paid"
cbotypecin.AddItem "Free"
cbotypecin.AddItem "Barter"

cbotypemag.AddItem "Paid"
cbotypemag.AddItem "Free"
cbotypemag.AddItem "Barter"

cbotypeol.AddItem "Paid"
cbotypeol.AddItem "Free"
cbotypeol.AddItem "Barter"

 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "delete * from dumBo_TRAmagbo"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "delete * from dumBo_TRAtvbo"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumBo_TRAcinbo"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumBo_TRAolbo"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
       
End Sub

Private Sub populateproducts()
    cboProduct.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from products Order by Product_Name"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        rs.MoveFirst
            Do Until rs.EOF
              cboProduct.AddItem rs!product_name
            rs.MoveNext
       Loop
    End If
 End Sub

Private Sub populateagency()
  CboAgency.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from agndtls Order by agentName"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        rs.MoveFirst
            Do Until rs.EOF
              CboAgency.AddItem rs!agentname
            rs.MoveNext
          Loop
    End If
End Sub
Private Sub populateclient()
  CboClient.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from clientdtls Order by ClientName"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        rs.MoveFirst
            Do Until rs.EOF
              CboClient.AddItem rs!clientname
            rs.MoveNext
       Loop
    End If
End Sub
Private Sub populatemedia()
    cboMediatype.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Media Order by Media_Type"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        cboMediatype.AddItem "Cinema"
        rs.MoveFirst
            Do Until rs.EOF
              If Mid(rs!Media_Type, 1, 3) = "Cin" Then
               rs.MoveNext
              Else
                cboMediatype.AddItem rs!Media_Type & "  " & Trim(rs!sub_media)
                rs.MoveNext
              End If
       Loop
    End If
 End Sub

Private Sub Populateregion()
    cboregion.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select distinct(region) from bo_mas Order by region"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
       
        rs.MoveFirst
       Do Until rs.EOF
        If IsNull(rs!region) = True Then
         rs.MoveNext
        Else
          If rs!region = "" Then
           rs.MoveNext
          End If
          
         cboregion.AddItem rs!region
         rs.MoveNext
        End If
       Loop
    End If
 End Sub
Private Sub AutoIncrementVoucher()
X = 0
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select serial_no from BO_mas order by serial_no"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
If rs.RecordCount <> 0 Then
   rs.MoveLast
   X = Val(rs!serial_no)
   lblserialno.Caption = X + 1
Else
   lblserialno = 100001
End If
End Sub

Private Sub PVMaskEdit1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPage.SetFocus
End Sub

Private Sub PVMaskTime_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdesctv.SetFocus
End Sub

Private Sub txtAddDiscountCin_LostFocus()
  If txtDisccin.Text = "" Then txtDisccin.Text = 0
  If txtAddDiscountCin.Text = "" Then txtAddDiscountCin.Text = 0
 txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100) - (Val(Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)) * txtDisccin / 100) - Val(txtAddDiscountCin.Text)
 cmdAdd.SetFocus
End Sub

Private Sub txtadddiscountmag_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtSurcharge.SetFocus
   
End Sub
Private Sub txtadddiscountCIN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd.SetFocus

End Sub
Private Sub txtadddiscountmag_LostFocus()
  If txtdiscmag.Text = "" Then txtdiscmag.Text = 0
  If txtadddiscountmag.Text = "" Then txtadddiscountmag.Text = 0
  If txtSurcharge.Text = "" Then txtSurcharge.Text = 0
  
 txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text) + Val(txtSurcharge)
 txtSurcharge.SetFocus
End Sub
Private Sub txtadddiscountOL_LostFocus()
  If txtdiscOL.Text = "" Then txtdiscOL.Text = 0
  If txtAddDiscountOL.Text = "" Then txtAddDiscountOL.Text = 0

 txtNetAmountOL.Text = Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100) - (Val(Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100)) * txtdiscOL / 100) - Val(txtAddDiscountOL.Text)
 cmdAdd.SetFocus
End Sub
Private Sub txtadddiscountTV_LostFocus()
 If txtdiscTV.Text = "" Then txtdiscTV.Text = 0
 If txtadddiscounttv.Text = "" Then txtadddiscounttv.Text = 0
 txtnetamountTV.Text = Val(txtgramountTV.Text) - (Val(txtgramountTV.Text) * Val(txtcompertv.Text) / 100) - (Val(Val(txtgramountTV.Text) - (Val(txtgramountTV.Text) * Val(txtcompertv.Text) / 100)) * txtdiscTV / 100) - Val(txtadddiscounttv)
 cmdAdd.SetFocus
End Sub
Private Sub txtadddiscountTV_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub

Private Sub txtadddiscountOL_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdAdd.SetFocus
 End Sub
Private Sub txtamountcin_GotFocus()
 SendKeys "{Home} + {End}"
End Sub

Private Sub txtamountcin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cbosubmedia.SetFocus
End Sub

Private Sub txtamountcin_LostFocus()
  
  If ValidateData = True Then
      
   If IsNumeric(txtamountcin.Text) = False Or IsNull(txtamountcin) = True Then
      MsgBox "Invalid Amount", vbInformation, "Invalid Entry"
      cbosubmedia.SetFocus
      Exit Sub
   End If
    
   If cbosubmedia.Text = "" Then
      MsgBox "Invalid Sub Media", vbInformation, "Invalid Entry"
      cbosubmedia.SetFocus
      Exit Sub
   End If
   
   If cbodaycin.Text = "" Then
      MsgBox "Invalid Days", vbInformation, "Invalid Entry"
      cbodaycin.SetFocus
      Exit Sub
   End If
     
   If cbolength.Text = "" Then
      MsgBox "Invalid Seconds", vbInformation, "Invalid Entry"
      cbolength.SetFocus
      Exit Sub
   End If
    
  
       
    If cbotypecin.Text = "" Then
       MsgBox "Invalid Payment Type", vbInformation, "Invalid Entry"
       cbotypecin.SetFocus
       Exit Sub
    End If
             
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = " select * from dumBo_TRAcinbo"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
     If rs.RecordCount = 0 Then
       Sqlqry1 = " Insert into dumBo_TRAcinbo values('" & Val(lblserialno.Caption) & "','" & cboyear.Text & "','" _
                                     & cbomonth.Text & "','" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Cinema','" _
                                     & Trim(cbosubmedia.Text) & "',#" _
                                     & Format(txtCinDateFrom.TextWithMask, "dd/mm/yyyy") & "#,#" _
                                     & Format(txtCinDateTo.TextWithMask, "dd/mm/yyyy") & "#,'" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "','" _
                                     & Trim(cbodaycin.Text) & "','" _
                                     & Trim(cbolength.Text) & "','" _
                                     & findfirstfixup(Trim(txtDesccin.Text)) & "','" _
                                     & findfirstfixup(Trim(cboMatCin.Text)) & "','" _
                                     & Trim(cbotypecin.Text) & "','" _
                                     & Trim(cboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(Trim(txtamountcin.Text)) & ", " & Val(Trim(txtamountcin.Text)) * Val(txtConvRate.Text) & ")"

        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
        
        Sqlqry2 = "select * from dumBo_TRAcinbo"
        Set rs = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlexcin.Clear
            Exit Sub
        Else
            Flexitemscin
            rs.MoveFirst
            Do Until rs.EOF
              MSFlexcin.AddItem rs!sub_media & Chr(9) & Format(rs!datefrom, "dd/mm/yyyy") & Chr(9) & Format(rs!dateto, "DD/MM/YYYY") & Chr(9) & rs!Day & Chr(9) & rs!Length & Chr(9) & rs!Description & Chr(9) & rs!mat_code & Chr(9) & rs!Type & Chr(9) & rs!tra_amount
              rs.MoveNext
            Loop
        End If
       
            
       If cbotypecin.Text = "Paid" Then
            txtGrAmountCin.Text = Val(txtamountcin.Text)
            txtGrAmountCin.Alignment = 2
             If txtDisccin.Text = "" Then
                  txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)
                Else
                   If txtAddDiscountCin.Text = "" Then
                    txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100) - (Val(Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)) * txtDisccin / 100)
                   Else
                    txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100) - (Val(Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)) * txtDisccin / 100) - Val(txtAddDiscountCin)
                   End If
                End If

       ElseIf cbotypecin.Text = "Free" Then
             txtfreecin.Text = Val(txtamountcin.Text)
       Else
            txtbarterCin.Text = Val(txtamountcin.Text)
       End If
                 
   Else
        U = 0
        Y = 0
        Z = 0
        
        rs.MoveFirst
         Sqlqry1 = " Insert into dumBo_TRAcinbo values('" & Val(lblserialno.Caption) & "','" & cboyear.Text & "','" _
                                     & cbomonth.Text & "','" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Cinema','" _
                                     & Trim(cbosubmedia) & "',#" _
                                     & Format(txtCinDateFrom.TextWithMask, "dd/mm/yyyy") & "#,#" _
                                     & Format(txtCinDateTo.TextWithMask, "dd/mm/yyyy") & "#,'" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "','" _
                                     & Trim(cbodaycin.Text) & "','" _
                                     & Trim(cbolength.Text) & "','" _
                                     & findfirstfixup(Trim(txtDesccin.Text)) & "','" _
                                     & findfirstfixup(Trim(cboMatCin.Text)) & "','" _
                                     & Trim(cbotypecin.Text) & "','" _
                                     & Trim(cboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(Trim(txtamountcin.Text)) & ", " & Val(Trim(txtamountcin.Text)) * Val(txtConvRate.Text) & ")"


        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
       
        
        Sqlqry1 = "select * from dumBo_TRAcinbo"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlexcin.Clear
            Exit Sub
        Else
            Flexitemscin
            rs.MoveFirst
            Do Until rs.EOF
                     MSFlexcin.AddItem rs!sub_media & Chr(9) & rs!datefrom & Chr(9) & rs!dateto & Chr(9) & rs!Day & Chr(9) & rs!Length & Chr(9) & rs!Description & Chr(9) & rs!mat_code & Chr(9) & rs!Type & Chr(9) & rs!tra_amount
              rs.MoveNext
            Loop
        End If
        
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAcinbo where type='Paid'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then U = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAcinbo where type='Free'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Y = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAcinbo where type='Barter'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Z = rs1.Fields(0)
            
            
            txtGrAmountCin.Text = U
            txtfreecin.Text = Y
            txtbarterCin.Text = Z
            
              If txtDisccin.Text = "" Then
                  txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)
                Else
                   If txtAddDiscountCin.Text = "" Then
                    txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100) - (Val(Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)) * txtDisccin / 100)
                   Else
                    txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100) - (Val(Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)) * txtDisccin / 100) - Val(txtAddDiscountCin)
                   End If
                End If

            
          
             
          U = 0
          Y = 0
          Z = 0
          
      
      cbosubmedia.SetFocus
      End If
    Else
     cbosubmedia.SetFocus
     Exit Sub
 End If
 
End Sub
Private Sub txtamountmag_GotFocus()
 SendKeys "{Home} + {End}"
End Sub
Private Sub txtamountmag_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtissueno.SetFocus
End Sub
Private Sub txtamountol_GotFocus()
 SendKeys "{Home} + {End}"
End Sub
Private Sub txtamountol_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdayol.SetFocus
End Sub

Private Sub txtboref_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboregion.SetFocus
End Sub
Private Sub cbodaycin_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbolength.SetFocus
End Sub
Private Sub txtdayol_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdescol.SetFocus
End Sub

Private Sub txtCinDateFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCinDateTo.SetFocus
End Sub

Private Sub txtCinDateFrom_LostFocus()
    If IsDate(txtCinDateFrom.TextWithMask) = False Then
       MsgBox "Invalid Cinema Date From", vbInformation, "Invalid Entry"
       txtCinDateFrom.SetFocus
       SendKeys " {Home} + {End} "
    End If
End Sub

Private Sub txtCinDateTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cbodaycin.SetFocus
End Sub

Private Sub txtCinDateTo_LostFocus()
If IsDate(txtCinDateTo.TextWithMask) = False Then
   MsgBox "Invalid Cinema Date To", vbInformation, "Invalid Entry"
   txtCinDateTo.SetFocus
   SendKeys " {Home} + {End} "
End If

End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cbomatmag.SetFocus
End Sub

Private Sub txtComPerCin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtDisccin.SetFocus
End Sub

Private Sub txtComPerCin_LostFocus()
If txtDisccin.Text = "" Then
    txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)
  Else
     If txtAddDiscountCin.Text = "" Then
      txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100) - (Val(Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)) * txtDisccin / 100)
     Else
      txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100) - (Val(Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)) * txtDisccin / 100) - Val(txtAddDiscountCin)
     End If
  End If
End Sub

Private Sub txtcompermag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdiscmag.SetFocus
End Sub

Private Sub txtcompermag_LostFocus()
If txtcompermag.Text = "" Then
    If txtdiscmag.Text = "" Then
       txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)
     Else
       If txtadddiscountmag.Text = "" Then
        txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100)
       Else
         If txtSurcharge.Text = "" Then
          txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text)
         Else
          txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text) + Val(txtSurcharge)
         End If
       End If
     End If
  Else
      If txtdiscmag.Text = "" Then
       txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)
     Else
       If txtadddiscountmag.Text = "" Then
        txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100)
       Else
         If txtSurcharge.Text = "" Then
          txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text)
         Else
          txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text) + Val(txtSurcharge)
         End If
       End If
     End If
  End If
End Sub
Private Sub txtcomperOL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdiscOL.SetFocus
End Sub

Private Sub txtcompertv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdiscTV.SetFocus
End Sub

Private Sub txtcompertv_LostFocus()
 If txtdiscTV.Text = "" Then txtdiscTV.Text = 0
 If txtadddiscounttv.Text = "" Then txtadddiscounttv.Text = 0
 txtnetamountTV.Text = Val(txtgramountTV.Text) - (Val(txtgramountTV.Text) * Val(txtcompertv.Text) / 100) - (Val(Val(txtgramountTV.Text) - (Val(txtgramountTV.Text) * Val(txtcompertv.Text) / 100)) * txtdiscTV / 100) - Val(txtadddiscounttv)
End Sub

Private Sub txtConvRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboProduct.SetFocus
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

Private Sub txtdatefrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdateto.SetFocus
End Sub

Private Sub txtdatefrom_LostFocus()
If IsDate(txtdatefrom.TextWithMask) = False Then
   MsgBox "Invalid Date From", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtdateto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdescol.SetFocus
End Sub

Private Sub txtdateto_LostFocus()
If IsDate(txtdateto.TextWithMask) = False Then
   MsgBox "Invalid Date To", vbInformation, "Invalid Entry"
   txtdateto.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtDaytv_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then PVMaskTime.SetFocus
End Sub
Private Sub txtDesccin_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboMatCin.SetFocus
End Sub
Private Sub txtdescmag_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtComments.SetFocus
End Sub
Private Sub txtdescol_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbomatol.SetFocus
End Sub
Private Sub txtdesctv_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboMattv.SetFocus
End Sub
Private Sub txtDisccin_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtAddDiscountCin.SetFocus
End Sub

Private Sub txtDisccin_LostFocus()
  If txtDisccin.Text = "" Then
    txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)
  Else
     If txtAddDiscountCin.Text = "" Then
      txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100) - (Val(Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)) * txtDisccin / 100)
     Else
      txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100) - (Val(Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)) * txtDisccin / 100) - Val(txtAddDiscountCin)
     End If
  End If
End Sub
Private Sub txtdiscmag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtadddiscountmag.SetFocus
End Sub

Private Sub txtdiscmag_LostFocus()
 If txtdiscmag.Text = "" Then
   txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)
 Else
   If txtadddiscountmag.Text = "" Then
    txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100)
   Else
     If txtSurcharge.Text = "" Then
      txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text)
     Else
      txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text) + Val(txtSurcharge)
     End If
   End If
 End If
End Sub

Private Sub txtdiscOL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAddDiscountOL.SetFocus
End Sub

Private Sub txtdiscOL_LostFocus()
If txtdiscOL.Text = "" Then
 txtNetAmountOL.Text = Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100)
Else
  If txtAddDiscountOL.Text = "" Then
   txtNetAmountOL.Text = Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100) - (Val(Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100)) * txtdiscOL / 100)
  Else
   txtNetAmountOL.Text = Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100) - (Val(Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100)) * txtdiscOL / 100) - Val(txtAddDiscountOL.Text)
  End If
End If
End Sub

Private Sub txtdiscTV_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtadddiscounttv.SetFocus
End Sub

Private Sub txtdiscTV_LostFocus()
 If txtdiscTV.Text = "" Then txtdiscTV.Text = 0
 If txtadddiscounttv.Text = "" Then txtadddiscounttv.Text = 0
 txtnetamountTV.Text = Val(txtgramountTV.Text) - (Val(txtgramountTV.Text) * Val(txtcompertv.Text) / 100) - (Val(Val(txtgramountTV.Text) - (Val(txtgramountTV.Text) * Val(txtcompertv.Text) / 100)) * txtdiscTV / 100) - Val(txtadddiscounttv)
End Sub

Private Sub txtimpression_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtrateol.SetFocus
End Sub

Private Sub txtissdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPage.SetFocus
End Sub

Private Sub txtissdate_LostFocus()
If IsDate(txtissdate.TextWithMask) = False Then
   MsgBox "Invalid issue Date", vbInformation, "Invalid Entry"
   txtissdate.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtissueno_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtissdate.SetFocus
End Sub
Private Sub cbolength_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtDesccin.SetFocus
End Sub
Private Sub txtPage_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbospace.SetFocus
End Sub
Private Sub txtrateol_GotFocus()
 SendKeys "{Home} + {End}"
End Sub
Private Sub txtrateol_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdatefrom.SetFocus
End Sub
Private Sub txtrateol_LostFocus()
   If ValidateData = True Then
      
   If IsNull(txtimpression) = True Or IsNumeric(txtimpression.Text) = False Then
     MsgBox " Invalid Impressions", vbInformation, "Invalid Entry"
     txtimpression.SetFocus
     Exit Sub
   End If
   
   If IsNull(txtrateol.Text) = True Or IsNumeric(txtrateol.Text) = False Then
     MsgBox " Invalid rate", vbInformation, "Invalid Entry"
     txtrateol.SetFocus
     Exit Sub
   End If
  
 
       
    If cbotypeol.Text = "" Then
       MsgBox "Invalid Payment Type", vbInformation, "Invalid Entry"
       cbotypeol.SetFocus
       Exit Sub
    End If
                
   txtamountol.Text = Val(txtimpression) / 1000 * Val(txtrateol.Text)
                
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = " select * from dumBo_TRAolbo"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If rs.RecordCount = 0 Then
       Sqlqry1 = " Insert into dumBo_TRAolbo values('" & Val(lblserialno.Caption) & "','" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','" & Format(txtdateto.TextWithMask, "dd/mm/yyyy") & "','" & cboyear.Text & "','" _
                                     & cbomonth.Text & "','" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Online','" _
                                     & Trim(Mid(cboMediatype, 8, 30)) & "','" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "','" _
                                     & findfirstfixup(Trim(txtdescol.Text)) & "','" _
                                     & findfirstfixup(Trim(cbomatol.Text)) & "','" _
                                     & Trim(cbotypeol.Text) & "'," _
                                     & Val(Trim(txtimpression.Text)) & "," _
                                     & Val(Trim(txtrateol.Text)) & ",'" _
                                     & Trim(cboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(Trim(txtamountol.Text)) & "," & Val(Trim(txtamountol.Text)) * Val(txtConvRate.Text) & " )"

        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
        
        Sqlqry1 = "select * from dumBo_TRAolbo"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlexol.Clear
            Exit Sub
        Else
            Flexitemol
            rs.MoveFirst
            Do Until rs.EOF
              MSFlexol.AddItem Format(rs!date_From, "dd/mm/yyyy") & Chr(9) & Format(rs!date_to, "dd/mm/yyyy") & Chr(9) & rs!Description & Chr(9) & rs!mat_code & Chr(9) & rs!Type & Chr(9) & rs!impressions & Chr(9) & rs!Rate & Chr(9) & rs!tra_amount
              rs.MoveNext
            Loop
        End If
            If cbotypeol.Text = "Paid" Then
                txtGrAmountOL.Text = Val(txtamountol.Text)
                txtGrAmountOL.Alignment = 2
                If txtdiscOL.Text = "" Then
                     txtNetAmountOL.Text = Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100)
                    Else
                      If txtAddDiscountOL.Text = "" Then
                       txtNetAmountOL.Text = Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100) - (Val(Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100)) * txtdiscOL / 100)
                      Else
                       txtNetAmountOL.Text = Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100) - (Val(Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100)) * txtdiscOL / 100) - Val(txtAddDiscountOL.Text)
                      End If
                    End If


            ElseIf cbotypeol.Text = "Free" Then
                txtfreeol.Text = Val(txtamountol.Text)
            Else
               txtbarterol.Text = Val(txtamountol.Text)
            End If
            
            
      Else
        
        U = 0
        Y = 0
        Z = 0
        
        rs.MoveFirst
       Sqlqry1 = " Insert into dumBo_TRAolbo values('" & Val(lblserialno.Caption) & "','" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','" & Format(txtdateto.TextWithMask, "dd/mm/yyyy") & "','" & cboyear.Text & "','" _
                                     & cbomonth.Text & "','" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Online','" _
                                     & Trim(Mid(cboMediatype, 8, 30)) & "','" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "','" _
                                     & findfirstfixup(Trim(txtdescol.Text)) & "','" _
                                     & findfirstfixup(Trim(cbomatol.Text)) & "','" _
                                     & Trim(cbotypeol.Text) & "'," _
                                     & Val(Trim(txtimpression.Text)) & "," _
                                     & Val(Trim(txtrateol.Text)) & ",'" _
                                     & Trim(cboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(Trim(txtamountol.Text)) & ", " & Val(Trim(txtamountol.Text)) * Val(txtConvRate.Text) & ")"
        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
       
        
        Sqlqry1 = "select * from dumBo_TRAolbo"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlexol.Clear
            Exit Sub
        Else
            Flexitemol
            rs.MoveFirst
            Do Until rs.EOF
               MSFlexol.AddItem Format(rs!date_From, "dd/mm/yyyy") & Chr(9) & Format(rs!date_to, "dd/mm/yyyy") & Chr(9) & rs!Description & Chr(9) & rs!mat_code & Chr(9) & rs!Type & Chr(9) & rs!impressions & Chr(9) & rs!Rate & Chr(9) & rs!tra_amount
              rs.MoveNext
            Loop
        End If
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAolbo where type='Paid'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then U = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAolbo where type='Free'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Y = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAolbo where type='Barter'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Z = rs1.Fields(0)
            
            
            txtGrAmountOL.Text = U
            txtfreeol.Text = Y
            txtbarterol.Text = Z
            
            If txtdiscOL.Text = "" Then
             txtNetAmountOL.Text = Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100)
            Else
              If txtAddDiscountOL.Text = "" Then
               txtNetAmountOL.Text = Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100) - (Val(Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100)) * txtdiscOL / 100)
              Else
               txtNetAmountOL.Text = Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100) - (Val(Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100)) * txtdiscOL / 100) - Val(txtAddDiscountOL.Text)
              End If
            End If

            
          U = 0
          Y = 0
          Z = 0
             
        
      End If
    Else
     Exit Sub
 End If
 
 
End Sub

Private Sub txtRatetv_GotFocus()
SendKeys "{Home} + {End}"
End Sub

Private Sub txtRatetv_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtDaytv.SetFocus
End Sub

Private Sub txtRatetv_LostFocus()


If ValidateData = True Then
      
   If IsNull(txtSpots.Text) = True Or IsNumeric(txtSpots.Text) = False Then
     MsgBox " Invalid Spots", vbInformation, "Invalid Entry"
     txtSpots.SetFocus
     Exit Sub
   End If
   
   If txtRatetv.Text = "" Or IsNumeric(txtRatetv.Text) = False Then
     MsgBox " Invalid rate", vbInformation, "Invalid Entry"
     txtRatetv.SetFocus
     Exit Sub
   End If
        
  
       
   If cbotypetv.Text = "" Then
       MsgBox "Invalid Payment Type", vbInformation, "Invalid Entry"
       cbotypetv.SetFocus
       Exit Sub
   End If
             
    txtamounttv.Text = Val(txtSpots.Text) * Val(txtRatetv)
     
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = " select * from dumBo_TRAtvbo"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If rs.RecordCount = 0 Then
       Sqlqry1 = " Insert into dumBo_TRAtvbo values('" & Val(lblserialno.Caption) & "','" & cboyear.Text & "','" _
                                     & cbomonth.Text & "','" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Television','" _
                                     & Trim(Mid(cboMediatype, 12, 30)) & "','" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "','" _
                                     & Trim(txtDaytv.Text) & "','" _
                                     & Trim(PVMaskTime.TextWithMask) & "','" _
                                     & findfirstfixup(Trim(txtdesctv.Text)) & "','" _
                                     & findfirstfixup(Trim(cboMattv.Text)) & "','" _
                                     & Trim(cbotypetv.Text) & "'," _
                                     & Val(Trim(txtSpots.Text)) & "," _
                                     & Val(Trim(txtRatetv.Text)) & ",'" _
                                     & Trim(cboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(Trim(txtamounttv.Text)) & ", " & Val(Trim(txtamounttv.Text)) * Val(txtConvRate.Text) & ")"


        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
        
        Sqlqry1 = "select * from dumBo_TRAtvbo"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlextv.Clear
            Exit Sub
        Else
            Flexitemstv
            rs.MoveFirst
            Do Until rs.EOF
              MSFlextv.AddItem rs!Day & Chr(9) & rs!Time & Chr(9) & rs!Description & Chr(9) & rs!mat_code & Chr(9) & rs!Type & Chr(9) & rs!spots & Chr(9) & rs!Rate & Chr(9) & rs!tra_amount
              rs.MoveNext
            Loop
        End If
               If cbotypetv.Text = "Paid" Then
                txtgramountTV.Text = Val(txtamounttv.Text)
                txtgramountTV.Alignment = 2
                If txtdiscTV.Text = "" Then txtdiscTV.Text = 0
                If txtadddiscounttv.Text = "" Then txtadddiscounttv.Text = 0
                txtnetamountTV.Text = Val(txtgramountTV.Text) - (Val(txtgramountTV.Text) * Val(txtcompertv.Text) / 100) - (Val(Val(txtgramountTV.Text) - (Val(txtgramountTV.Text) * Val(txtcompertv.Text) / 100)) * txtdiscTV / 100) - Val(txtadddiscounttv)
                
              
               ElseIf cbotypetv.Text = "Free" Then
                 txtfreeTv.Text = Val(txtamounttv.Text)
               Else
                 txtbarterTv.Text = Val(txtamounttv.Text)
               End If
               
  Else
       
      U = 0
      Y = 0
      Z = 0
            
       rs.MoveFirst
       Sqlqry1 = " Insert into dumBo_TRAtvbo values('" & Val(lblserialno.Caption) & "','" & cboyear.Text & "','" _
                                     & cbomonth.Text & "','" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Television','" _
                                     & Trim(Mid(cboMediatype, 12, 30)) & "','" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "','" _
                                     & Trim(txtDaytv.Text) & "','" _
                                     & Trim(PVMaskTime.TextWithMask) & "','" _
                                     & findfirstfixup(Trim(txtdesctv.Text)) & "','" _
                                     & findfirstfixup(Trim(cboMattv.Text)) & "','" _
                                     & Trim(cbotypetv.Text) & "'," _
                                     & Val(Trim(txtSpots.Text)) & "," _
                                     & Val(Trim(txtRatetv.Text)) & ",'" _
                                     & Trim(cboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(Trim(txtamounttv.Text)) & ", " & Val(Trim(txtamounttv.Text)) * Val(txtConvRate.Text) & ")"


        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
       
        
        Sqlqry1 = "select * from dumBo_TRAtvbo"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlextv.Clear
            Exit Sub
        Else
            Flexitemstv
            rs.MoveFirst
            Do Until rs.EOF
               MSFlextv.AddItem rs!Day & Chr(9) & rs!Time & Chr(9) & rs!Description & Chr(9) & rs!mat_code & Chr(9) & rs!Type & Chr(9) & rs!spots & Chr(9) & rs!Rate & Chr(9) & rs!tra_amount
              rs.MoveNext
            Loop
        End If
        
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAtvbo where type='Paid'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then U = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAtvbo where type='Free'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Y = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAtvbo where type='Barter'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Z = rs1.Fields(0)
            
            
            txtgramountTV.Text = U
            txtfreeTv.Text = Y
            txtbarterTv.Text = Z
            
            If txtdiscTV.Text = "" Then txtdiscTV.Text = 0
            If txtadddiscounttv.Text = "" Then txtadddiscounttv.Text = 0
            txtnetamountTV.Text = Val(txtgramountTV.Text) - (Val(txtgramountTV.Text) * Val(txtcompertv.Text) / 100) - (Val(Val(txtgramountTV.Text) - (Val(txtgramountTV.Text) * Val(txtcompertv.Text) / 100)) * txtdiscTV / 100) - Val(txtadddiscounttv)
            

             
          U = 0
          Y = 0
          Z = 0
               
    
      End If
    Else
     Exit Sub
 End If
End Sub

Private Sub CboRegion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtremarks.SetFocus
End Sub

Private Sub txtremarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If mtype = 1 Then
  txtDaytv.SetFocus
 ElseIf mtype = 2 Then
  txtdatefrom.SetFocus
 ElseIf mtype = 3 Then
  cbosubmedia.SetFocus
 ElseIf mtype = 4 Then
  txtissueno.SetFocus
 End If
End If
End Sub

Private Sub txtSpots_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtRatetv.SetFocus
End Sub

Private Function ValidateData()
ValidateData = False
If cboyear.Text = "" Then
  MsgBox "Invalid Year ", vbInformation, "Invalid Entry"
  cboyear.SetFocus
  SendKeys "{Home} + {End}"
  Exit Function
ElseIf cbomonth.Text = "" Then
  MsgBox "Invalid Month", vbInformation, "Invalid Entry"
  cbomonth.SetFocus
  Exit Function
ElseIf cboProduct.Text = "" Then
  MsgBox "Invalid Product", vbInformation, "Invalid Entry"
  cboProduct.SetFocus
  Exit Function
ElseIf cboMediatype.Text = "" Then
  MsgBox "Invalid Media Type", vbInformation, "Invalid Entry"
  cboMediatype.SetFocus
  Exit Function
ElseIf cboCurrency.Text = "" Then
  MsgBox "Select Currency Type", vbInformation, "Invalid Entry"
  cboCurrency.SetFocus
  Exit Function
ElseIf txtConvRate.Text = "" Then
  MsgBox "Enter Convertion Rate - - cannot be zero", vbInformation, "Invalid Entry"
  txtConvRate.SetFocus
  Exit Function
  
Else
  ValidateData = True
End If
End Function

Private Sub Flexitemsmag()
 With MSFlexmag
     .Clear
    .AllowUserResizing = flexResizeColumns
    .Rows = 1
    .Cols = 9
    .Col = 0
    .Row = 0
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Iss #"
    .ColAlignment(0) = 0
    .ColWidth(0) = 600
    .ColWidth(1) = 1100
    .ColWidth(2) = 500
    .ColWidth(3) = 1800
    .ColWidth(4) = 1800
    .ColWidth(5) = 1900
    .ColWidth(6) = 1500
    .ColWidth(7) = 800
    .ColWidth(8) = 900
    .Col = 1
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Date"
    .Col = 2
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Page"
    .Col = 3
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Space"
    
    .Col = 4
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Reference"
    
    .Col = 5
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Comments"
      
    .Col = 6
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Material"
    .Col = 7
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "P_Type"
    .Col = 8
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Amount"
    .Row = 0
    .Col = 1
  
  End With
End Sub

Private Sub Msflexmag_dblclick()
 Dim i
 Dim j
 Dim X
 Dim Y, Z, U
 
 X = MSFlexmag.Rows
 
 If X > 1 Then
   i = MsgBox(" Are you sure .. ! You want to Remove this transaction", vbInformation + vbYesNo)
    If i = vbYes Then
     With MSFlexmag
        j = .Row
        .Col = 0
        txtissueno = .Text
        .Col = 1
        txtdate.TextWithMask = .Text
        .Col = 2
        txtPage = .Text
        .Col = 3
        cbospace = .Text
        .Col = 4
        txtdescmag = .Text
        .Col = 5
        txtComments = .Text
        .Col = 6
        cbomatmag = .Text
        .Col = 7
        cbotypemag = .Text
        .Col = 8
        txtamountmag = .Text
                            
        .RemoveItem (j)
                             
        Sqlqry1 = "Select * from dumBo_TRAmagbo where  Issue_no ='" & Trim(txtissueno) & "' and Page ='" & Trim(txtPage) & "' and space='" & Trim(cbospace.Text) & "' and tdate =#" & Format(txtissdate.TextWithMask, "DD/MM/YYYY") & "# and tra_amount =" & Val(txtamountmag) & ""
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs1.RecordCount <> 0 Then
        rs1.MoveLast
        rs1.Delete
        End If
        U = 0
        Y = 0
        Z = 0
                
        
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAmagbo where type='Paid'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then U = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAmagbo where type='Free'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Y = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAmagbo where type='Barter'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Z = rs1.Fields(0)
            
            
          txtGrAmountmag.Text = U
          txtfreemag.Text = Y
          txtbartermag.Text = Z
          If txtdiscmag.Text = "" Then txtdiscmag.Text = 0
          If txtadddiscountmag.Text = "" Then txtadddiscountmag.Text = 0
          If txtSurcharge.Text = "" Then txtSurcharge.Text = 0
          txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text) + Val(txtSurcharge)
                     
         
             
          U = 0
          Y = 0
          Z = 0
          
        
        
     End With
    End If
 End If
End Sub

Private Sub txtamountmag_LostFocus()
 If ValidateData = True Then
      
   If IsNumeric(txtamountmag.Text) = False Or IsNull(txtamountmag.Text) = True Then
      MsgBox "Invalid Amount", vbInformation, "Invalid Entry"
      txtamountmag.SetFocus
      Exit Sub
   End If
       
    If cbotypemag.Text = "" Then
       MsgBox "Invalid Payment Type", vbInformation, "Invalid Entry"
       cbotypemag.SetFocus
       Exit Sub
    End If
             
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = " select * from dumBo_TRAmagbo"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If rs.RecordCount = 0 Then
    
       Sqlqry1 = " Insert into dumBo_TRAmagbo values('" & Val(lblserialno.Caption) & "','" & cboyear.Text & "','" _
                                     & cbomonth.Text & "'," & Val(cbomonth.ListIndex) & ",'" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Magazine','" _
                                     & Trim(Mid(cboMediatype, 10, 30)) & "','" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "','" _
                                     & Trim(txtissueno.Text) & "','" _
                                     & Format(txtissdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & Trim(txtPage.Text) & "','" _
                                     & findfirstfixup(Trim(txtdescmag.Text)) & "','" _
                                     & findfirstfixup(Trim(txtComments.Text)) & "','" _
                                     & findfirstfixup(Trim(cbomatmag.Text)) & "','" _
                                     & Trim(cbospace.Text) & "','" _
                                     & Trim(cbotypemag.Text) & "','" _
                                     & Trim(cboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(Trim(txtamountmag.Text)) & ", " & Val(Trim(txtamountmag.Text)) * Val(txtConvRate.Text) & ")"

        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
        
        Sqlqry1 = "select * from dumBo_TRAmagbo"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlexmag.Clear
            Exit Sub
        Else
            Flexitemsmag
            rs.MoveFirst
            Do Until rs.EOF
              MSFlexmag.AddItem rs!issue_no & Chr(9) & Format(rs!tdate, "dd/mm/yyyy") & Chr(9) & rs!Page & Chr(9) & rs!Space & Chr(9) & rs!Description & Chr(9) & rs!Comments & Chr(9) & rs!mat_code & Chr(9) & rs!Type & Chr(9) & rs!tra_amount
              rs.MoveNext
            Loop
        End If
          If cbotypemag.Text = "Paid" Then
            txtGrAmountmag.Text = Val(txtamountmag.Text)
            txtGrAmountmag.Alignment = 2
            If txtdiscmag.Text = "" Then
                  txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)
                Else
                  If txtadddiscountmag.Text = "" Then
                   txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100)
                  Else
                    If txtSurcharge.Text = "" Then
                     txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text)
                    Else
                     txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text) + Val(txtSurcharge)
                    End If
                  End If
                End If
            
          ElseIf cbotypemag.Text = "Free" Then
             txtfreemag.Text = Val(txtamountmag.Text)
          Else
             txtbartermag.Text = Val(txtamountmag.Text)
          End If
            
      Else
        U = 0
        Y = 0
        Z = 0
        
        rs.MoveFirst
         
       Sqlqry1 = " Insert into dumBo_TRAmagbo values('" & Val(lblserialno.Caption) & "','" & cboyear.Text & "','" _
                                     & cbomonth.Text & "'," & Val(cbomonth.ListIndex) & ",'" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Magazine','" _
                                     & Trim(Mid(cboMediatype, 10, 30)) & "','" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "','" _
                                     & Trim(txtissueno.Text) & "','" _
                                     & Format(txtissdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & Trim(txtPage.Text) & "','" _
                                     & findfirstfixup(Trim(txtdescmag.Text)) & "','" _
                                     & findfirstfixup(Trim(txtComments.Text)) & "','" _
                                     & findfirstfixup(Trim(cbomatmag.Text)) & "','" _
                                     & Trim(cbospace.Text) & "','" _
                                     & Trim(cbotypemag.Text) & "','" _
                                     & Trim(cboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(Trim(txtamountmag.Text)) & ", " & Val(Trim(txtamountmag.Text)) * Val(txtConvRate.Text) & ")"

        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
        
        Sqlqry1 = "select * from dumBo_TRAmagbo"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlexmag.Clear
            Exit Sub
        Else
            Flexitemsmag
            rs.MoveFirst
            Do Until rs.EOF
              MSFlexmag.AddItem rs!issue_no & Chr(9) & Format(rs!tdate, "dd/mm/yyyy") & Chr(9) & rs!Page & Chr(9) & rs!Space & Chr(9) & rs!Description & Chr(9) & rs!Comments & Chr(9) & rs!mat_code & Chr(9) & rs!Type & Chr(9) & rs!tra_amount
              rs.MoveNext
            Loop
        End If
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAmagbo where type='Paid'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then U = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAmagbo where type='Free'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Y = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAmagbo where type='Barter'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Z = rs1.Fields(0)
            
            
            txtGrAmountmag.Text = U
            txtfreemag.Text = Y
            txtbartermag.Text = Z
            If txtdiscmag.Text = "" Then
              txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)
            Else
              If txtadddiscountmag.Text = "" Then
               txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100)
              Else
                If txtSurcharge.Text = "" Then
                 txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text)
                Else
                 txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text) + Val(txtSurcharge)
                End If
              End If
            End If
       
             
          U = 0
          Y = 0
          Z = 0
      End If
    Else
     Exit Sub
 End If
End Sub

Private Sub Flexitemscin()
With MSFlexcin

    .Clear
    .AllowUserResizing = flexResizeColumns
    .Rows = 1
    .Cols = 9
    .Col = 0
    .Row = 0
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Sub Media"
    .ColAlignment(0) = 0
    .ColWidth(0) = 2400
    .ColWidth(1) = 1100
    .ColWidth(2) = 1125
    .ColWidth(3) = 900
    .ColWidth(4) = 600
    .ColWidth(5) = 1300
    .ColWidth(6) = 1700
    .ColWidth(7) = 800
    .ColWidth(8) = 900
    
        
    .Col = 1
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Date From"
        
    .Col = 2
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Date To"
        
    .Col = 3
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Days"
    .Col = 4
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Secs."
    .Col = 5
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Description"
    .Col = 6
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Material"
    .Col = 7
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "P_Type"
    .Col = 8
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Amount"
    .Row = 0
    .Col = 1
  
  End With
End Sub

' new start  Cinema
Private Sub MsFlexCin_dblclick()
 Dim i
 Dim j
 Dim X
 Dim Y, Z, U
 
 X = MSFlexcin.Rows
 
 If X > 1 Then
   i = MsgBox(" Are you sure .. ! You want to Remove this transaction", vbInformation + vbYesNo)
    If i = vbYes Then
     With MSFlexcin
        j = .Row
        .Col = 0
        cbosubmedia = .Text
        .Col = 1
        txtCinDateFrom.TextWithMask = .Text
        .Col = 2
        txtCinDateTo.TextWithMask = .Text
        .Col = 3
        cbodaycin = .Text
        .Col = 4
        cbolength = .Text
        .Col = 5
        txtDesccin = .Text
        .Col = 6
        cboMatCin = .Text
        .Col = 7
        cbotypecin = .Text
        .Col = 8
        txtamountcin = .Text
                            
        .RemoveItem (j)
                             
                
        Sqlqry1 = "select * from dumBo_TRAcinbo where  description ='" & Trim(txtDesccin) & "' and Day ='" & Trim(cbodaycin) & "' and Length ='" & Trim(cbolength) & "' and tra_amount =" & Val(txtamountcin) & ""
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs1.RecordCount <> 0 Then
        rs1.MoveLast
        rs1.Delete
        End If
        
         Sqlqry1 = "select sum(tra_amount) from dumBo_TRAcinbo where type='Paid'"
         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
         If IsNull(rs1.Fields(0)) = False Then U = rs1.Fields(0)
         
         Sqlqry1 = "select sum(tra_amount) from dumBo_TRAcinbo where type='Free'"
         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Y = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAcinbo where type='Barter'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Z = rs1.Fields(0)
            
            
            txtGrAmountCin.Text = U
            txtfreecin.Text = Y
            txtbarterCin.Text = Z
            If txtDisccin.Text = "" Then txtDisccin.Text = 0
            If txtAddDiscountCin.Text = "" Then txtAddDiscountCin.Text = 0
            txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100) - (Val(Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)) * txtDisccin / 100) - Val(txtAddDiscountCin.Text)
              
          U = 0
          Y = 0
          Z = 0
     
     End With
    End If
 End If
End Sub

Private Sub Flexitemstv()
With MSFlextv
    .Clear
    .AllowUserResizing = flexResizeColumns
    .Rows = 1
    .Cols = 8
    .Col = 0
    .Row = 0
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Day"
    .ColAlignment(0) = 0
    .ColWidth(0) = 600
    .ColWidth(1) = 650
    .ColWidth(2) = 3300
    .ColWidth(3) = 2800
    .ColWidth(4) = 700
    .ColWidth(5) = 600
    .ColWidth(6) = 1000
    .ColWidth(7) = 1200
    
    .Col = 1
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Time"
    .Col = 2
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Description"
    
    .Col = 3
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Material"
    .Col = 4
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "P_Type"
    .Col = 5
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Spots"
    .Col = 6
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Rate"
    .Col = 7
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Amount"
  
  End With
End Sub

' new start  Television
Private Sub MsFlexTv_dblclick()
Dim i
 Dim j
 Dim X
 Dim Y, Z, U
 
 X = MSFlextv.Rows
 
 If X > 1 Then
  i = MsgBox(" Are you sure .. ! You want to Remove this transaction", vbInformation + vbYesNo)
   If i = vbYes Then
    With MSFlextv
        j = .Row
        .Col = 0
        txtDaytv = .Text
        .Col = 1
        PVMaskTime.TextWithMask = .Text
        .Col = 2
        txtdesctv = .Text
        .Col = 3
        cboMattv = .Text
        .Col = 4
        cbotypetv = .Text
        .Col = 5
        txtSpots = .Text
        .Col = 6
        txtRatetv = .Text
        .Col = 7
        txtamounttv = .Text
                            
        .RemoveItem (j)
             
        
        Sqlqry1 = "select * from dumBo_TRAtvbo where description ='" & Trim(txtdesctv) & "' and Day ='" & Trim(txtDaytv) & "' and Time ='" & Trim(PVMaskTime.TextWithMask) & "' and tra_amount =" & Val(txtamounttv) & ""
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs1.RecordCount <> 0 Then
        rs1.MoveLast
        rs1.Delete
        End If
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAtvbo where type='Paid'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then U = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAtvbo where type='Free'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Y = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAtvbo where type='Barter'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Z = rs1.Fields(0)
            
            
            txtgramountTV.Text = U
            txtfreeTv.Text = Y
            txtbarterTv.Text = Z
            If txtdiscTV.Text = "" Then txtdiscTV.Text = 0
            If txtadddiscounttv.Text = "" Then txtadddiscounttv.Text = 0
            txtnetamountTV.Text = Val(txtgramountTV.Text) - (Val(txtgramountTV.Text) * Val(txtcompertv.Text) / 100) - (Val(Val(txtgramountTV.Text) - (Val(txtgramountTV.Text) * Val(txtcompertv.Text) / 100)) * txtdiscTV / 100) - Val(txtadddiscounttv)
                       
           
           
          U = 0
          Y = 0
          Z = 0
          
        
        
     End With
    End If
 End If
End Sub
' new Online

Private Sub Flexitemol()
With MSFlexol
    
    .Clear
    .AllowUserResizing = flexResizeColumns
    .Rows = 1
    .Cols = 8
    .Col = 0
    .Row = 0
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Dt_From"
    .ColAlignment(0) = 0
    .ColWidth(0) = 1100
    .ColWidth(1) = 1100
    .ColWidth(2) = 2500
    .ColWidth(3) = 2450
    .ColWidth(4) = 800
    .ColWidth(5) = 1000
    .ColWidth(6) = 900
    .ColWidth(7) = 1000
    
        
    .Col = 1
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Dt_To"
        
    .Col = 2
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Description"
    .Col = 3
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Material"
    .Col = 4
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "P_Type"
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Col = 5
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Impres."
    .Col = 6
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Rate"
    .Col = 7
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 10
    .CellFontBold = True
    .Text = "Amount"
    .Row = 0
    .Col = 1
  
  End With
End Sub
Private Sub msflexol_dblclick()
 
 Dim i
 Dim j
 Dim X
 Dim Y, Z, U
 
 X = MSFlexol.Rows
 
 If X > 1 Then
  i = MsgBox(" Are you sure .. ! You want to Remove this transaction", vbInformation + vbYesNo)
   If i = vbYes Then
    With MSFlexol
        j = .Row
        .Col = 0
        txtdatefrom.TextWithMask = .Text
        .Col = 1
        txtdateto.TextWithMask = .Text
        .Col = 2
        txtdescol = .Text
        .Col = 3
        cbomatol = .Text
        .Col = 4
        cbotypeol = .Text
        .Col = 5
        txtimpression = .Text
        .Col = 6
        txtrateol = .Text
        .Col = 7
        txtamountol = .Text
                            
        .RemoveItem (j)
                             
                       
        
        'Sqlqry1 = "Delete * from dumBo_TRAolbo where serial_no='" & Trim(lblserialno.Caption) & "' and description ='" & Trim(txtdescol) & "' and date_from=#" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "# and mat_code ='" & findfirstfixup(Trim(cbomatol.Text)) & "' and tra_amount =" & Val(txtamountol) & ""
        Sqlqry1 = "select * from dumBo_TRAolbo where Description ='" & Trim(txtdescol) & "' and date_from=#" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "# and mat_code ='" & findfirstfixup(Trim(cbomatol.Text)) & "' and tra_amount =" & Val(txtamountol) & ""
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs1.RecordCount <> 0 Then
         rs1.MoveLast
         rs1.Delete
        End If
        
        
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAolbo where type='Paid'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then U = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAolbo where type='Free'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Y = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from dumBo_TRAolbo where type='Barter'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Z = rs1.Fields(0)
            
            
            txtGrAmountOL.Text = U
            txtfreeol.Text = Y
            txtbarterol.Text = Z
             If txtdiscOL.Text = "" Then txtdiscOL.Text = 0
             If txtAddDiscountOL.Text = "" Then txtAddDiscountOL.Text = 0
             txtNetAmountOL.Text = Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100) - (Val(Val(txtGrAmountOL.Text) - (Val(txtGrAmountOL.Text) * Val(txtcomperOL.Text) / 100)) * txtdiscOL / 100) - Val(txtAddDiscountOL.Text)
            
          U = 0
          Y = 0
          Z = 0
          
        End With
    End If
 End If
End Sub
Private Sub txtSurcharge_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub

Private Sub txtSurcharge_LostFocus()
  If txtdiscmag.Text = "" Then txtdiscmag.Text = 0
  If txtadddiscountmag.Text = "" Then txtadddiscountmag.Text = 0
  If txtSurcharge.Text = "" Then txtSurcharge.Text = 0
  txtnetamountmag.Text = Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100) - (Val(Val(txtGrAmountmag.Text) - (Val(txtGrAmountmag.Text) * Val(txtcompermag.Text) / 100)) * txtdiscmag / 100) - Val(txtadddiscountmag.Text) + Val(txtSurcharge)
End Sub
