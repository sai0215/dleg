VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmBOCin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CinemaNew"
   ClientHeight    =   8595
   ClientLeft      =   75
   ClientTop       =   285
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1560
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
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
      Left            =   6480
      Picture         =   "frmBOCin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
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
      Left            =   5280
      Picture         =   "frmBOCin.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Add"
      DisabledPicture =   "frmBOCin.frx":0204
      DownPicture     =   "frmBOCin.frx":0736
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
      Left            =   4080
      MaskColor       =   &H008080FF&
      Picture         =   "frmBOCin.frx":0C68
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7695
      Left            =   120
      TabIndex        =   55
      Top             =   120
      Width           =   11655
      Begin VB.ComboBox CboClient 
         BackColor       =   &H80000018&
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
         Left            =   8760
         Style           =   2  'Dropdown List
         TabIndex        =   165
         Top             =   1320
         Width           =   2775
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
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   164
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
         Left            =   9600
         TabIndex        =   157
         Top             =   1800
         Width           =   1935
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
         Left            =   8760
         TabIndex        =   44
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
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cboMediatype 
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
         ForeColor       =   &H000040C0&
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   3615
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
         Left            =   6360
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
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
         Left            =   5280
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
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   4935
         Left            =   120
         TabIndex        =   84
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
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   4440
            Width           =   495
         End
         Begin VB.ComboBox cbolength 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   270
            Left            =   6120
            TabIndex        =   47
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
            Left            =   5040
            TabIndex        =   45
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
            TabIndex        =   38
            Top             =   720
            Width           =   2415
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
            TabIndex        =   140
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
            TabIndex        =   138
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
            TabIndex        =   54
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
            Locked          =   -1  'True
            TabIndex        =   121
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
            Locked          =   -1  'True
            TabIndex        =   120
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
            Left            =   8160
            TabIndex        =   49
            Top             =   720
            Width           =   1215
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
            Left            =   6960
            TabIndex        =   48
            Top             =   720
            Width           =   1095
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
            TabIndex        =   50
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
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   720
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexcin 
            Height          =   2415
            Left            =   120
            TabIndex        =   85
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
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   4440
            Width           =   495
         End
         Begin PVMaskEditLib.PVMaskEdit txtCinDateFrom 
            Height          =   255
            Left            =   2640
            TabIndex        =   159
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
            Left            =   3840
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
         Begin VB.Label Label73 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   2640
            TabIndex        =   162
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label78 
            BackColor       =   &H00FFFFFF&
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
            Left            =   3840
            TabIndex        =   161
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label63 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   146
            Top             =   4440
            Width           =   1335
         End
         Begin VB.Label Label60 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   142
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   141
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label Label58 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   139
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   125
            Top             =   4440
            Width           =   1455
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   124
            Top             =   4440
            Width           =   1455
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   123
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   122
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
            BackColor       =   &H00FFFFFF&
            Caption         =   "Days"
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
            Left            =   5040
            TabIndex        =   91
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Left            =   6960
            TabIndex        =   90
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sec."
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
            Left            =   6120
            TabIndex        =   89
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   88
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   8160
            TabIndex        =   87
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Left            =   10440
            TabIndex        =   86
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame FraMag 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   4935
         Left            =   120
         TabIndex        =   75
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
            TabIndex        =   128
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
            TabIndex        =   126
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
            TabIndex        =   107
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
            TabIndex        =   106
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
            TabIndex        =   101
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
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   163
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label72 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   158
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   154
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   153
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Label Label64 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   147
            Top             =   4320
            Width           =   975
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   129
            Top             =   3840
            Width           =   1335
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   127
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
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   105
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   104
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   103
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   102
            Top             =   4320
            Width           =   1815
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   81
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   80
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   79
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   78
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label22 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   77
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label24 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   76
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraTV 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   4935
         Left            =   120
         TabIndex        =   65
         Top             =   2640
         Width           =   11415
         Begin VB.ComboBox CboCode 
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
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   169
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtsec 
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
            Left            =   8040
            MaxLength       =   3
            TabIndex        =   166
            Top             =   480
            Width           =   495
         End
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
            TabIndex        =   32
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
            TabIndex        =   132
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
            TabIndex        =   130
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
            TabIndex        =   30
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
            TabIndex        =   33
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
            TabIndex        =   109
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
            TabIndex        =   108
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
            TabIndex        =   23
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
            Left            =   10200
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   480
            Width           =   1095
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
            Left            =   9240
            TabIndex        =   29
            Top             =   480
            Width           =   855
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
            Left            =   8640
            MaxLength       =   3
            TabIndex        =   28
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
            TabIndex        =   25
            Top             =   480
            Width           =   2055
         End
         Begin PVMaskEditLib.PVMaskEdit PVMaskTime 
            Height          =   255
            Left            =   840
            TabIndex        =   24
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
            Left            =   4800
            TabIndex        =   26
            Top             =   480
            Width           =   2055
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlextv 
            Height          =   2775
            Left            =   240
            TabIndex        =   74
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
            Left            =   6960
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label76 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Left            =   3840
            TabIndex        =   168
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label75 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sec"
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
            TabIndex        =   167
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label65 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   148
            Top             =   4320
            Width           =   1455
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   133
            Top             =   3720
            Width           =   1335
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   131
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   113
            Top             =   4320
            Width           =   1815
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   112
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   111
            Top             =   3720
            Width           =   1335
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   110
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
            BackColor       =   &H00FFFFFF&
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
            Left            =   10200
            TabIndex        =   73
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   4800
            TabIndex        =   71
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   70
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   69
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   1560
            TabIndex        =   68
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   66
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Left            =   8520
            TabIndex        =   72
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   9240
            TabIndex        =   67
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Fraemp 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   4935
         Left            =   120
         TabIndex        =   82
         Top             =   2640
         Width           =   11415
         Begin MSFlexGridLib.MSFlexGrid MSFleext 
            Height          =   4335
            Left            =   120
            TabIndex        =   83
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
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   4935
         Left            =   120
         TabIndex        =   92
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
            TabIndex        =   41
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
            TabIndex        =   136
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
            TabIndex        =   134
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
            TabIndex        =   40
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
            TabIndex        =   42
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
            TabIndex        =   115
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
            TabIndex        =   114
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
            TabIndex        =   39
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
            TabIndex        =   46
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
            TabIndex        =   36
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
            TabIndex        =   34
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
            TabIndex        =   35
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
            TabIndex        =   37
            Top             =   600
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexol 
            Height          =   2655
            Left            =   240
            TabIndex        =   93
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
            TabIndex        =   151
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
         Begin VB.Label Label67 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   150
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Top             =   4440
            Width           =   1455
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   137
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   135
            Top             =   3960
            Width           =   1215
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   119
            Top             =   4440
            Width           =   1815
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   118
            Top             =   4440
            Width           =   1455
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   117
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   116
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
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   100
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   99
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   97
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label31 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   96
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   95
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label29 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   94
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   98
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label71 
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
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   360
         TabIndex        =   156
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label70 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8760
         TabIndex        =   155
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label62 
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
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   840
         TabIndex        =   145
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblConvRate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Left            =   7560
         TabIndex        =   144
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label61 
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
         ForeColor       =   &H00004000&
         Height          =   345
         Left            =   4320
         TabIndex        =   143
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   64
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
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
         Left            =   5160
         TabIndex        =   63
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblserialno 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8760
         TabIndex        =   62
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
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
         Left            =   7800
         TabIndex        =   61
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   3960
         TabIndex        =   60
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8040
         TabIndex        =   59
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   360
         TabIndex        =   58
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   57
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   4200
         TabIndex        =   56
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmBOCin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim DT As Long
Dim m As Long
Dim Y As Long
Dim MTYPE As Long
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim sqlqry3 As String
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
Dim rs3 As Recordset
Dim ws As Workspace
Dim invdate As Date

Private Sub CboAgency_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CboClient.SetFocus
End Sub

Private Sub CboClient_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cboMediatype.SetFocus
End Sub

Private Sub CboCode_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cboMattv.SetFocus
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

Private Sub cbodaycin_GotFocus()

   If Mid(txtCinDateTo.TextWithMask, 4, 2) > 12 Then
      MsgBox "Invalid cinema Date to", vbInformation, "Invalid Entry"
      txtCinDateTo.SetFocus
      SendKeys " {Home} + {End} "
   End If
    
   If IsDate(Format(txtCinDateTo.TextWithMask, "dd/mm/yyyy")) = False Then
       MsgBox "Invalid Cinema Date To", vbInformation, "Invalid Entry"
       txtCinDateTo.SetFocus
       SendKeys " {Home} + {End} "
    End If
   
    
    If DateValue(Format(txtCinDateFrom.TextWithMask, "dd/mm/yyyy")) > DateValue(Format(txtCinDateTo.TextWithMask, "dd/mm/yyyy")) Then
      MsgBox "Date to cannot be lesser than date from", vbInformation, "Invalid entry"
      txtCinDateTo.SetFocus
      Exit Sub
    End If
End Sub

Private Sub cbolength_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtDesccin.SetFocus
End Sub

Private Sub cbolength_LostFocus()
    Dim C As String
    Dim X As Integer
    C = ""
    X = 0

 If Val(Mid(cbolength.Text, 1, 2)) = 10 Or Val(Mid(cbolength.Text, 1, 2)) = 15 Or Val(Trim(cbolength.Text)) = 30 Or Val(Trim(cbolength.Text)) = 60 Or Val(Trim(cbolength.Text)) = 90 Then
    
   If Val(Mid(cbolength.Text, 1, 2)) = 0 Then
    If Mid(cbodaycin, 1, 1) = "B" Then
       C = "BIW" & Val(Trim(cbolength.Text))
    Else
       C = "MON" & Val(Trim(cbolength.Text))
    End If
  Else
    If Mid(cbodaycin, 1, 1) = "B" Then
       C = "BIW" & Val(Mid(cbolength.Text, 1, 2))
    Else
       C = "MON" & Val(Mid(cbolength.Text, 1, 2))
    End If
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
If Mid(cboMediatype.Text, 1, 3) = "Cin" Then
   fraTV.Visible = False
   Fraol.Visible = False
   Fracin.Visible = True
   FraMag.Visible = False
   Fraemp.Visible = False
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from cinema_rates where mid(sub_media,1,2)<>'Sp' order by sub_media"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
        rs.MoveFirst
        cbosubmedia.Clear
        Do Until rs.EOF
         cbosubmedia.AddItem rs!sub_Media
         rs.MoveNext
        Loop
    End If
    
    rs.LockEdits = False
   
   cbodaycin.Clear
   cbodaycin.AddItem "Bi-Weekly"
   cbodaycin.AddItem "Monthly"
   
   cbolength.Clear
   cbolength.AddItem "10 Sl"
   cbolength.AddItem "15 FL"
   cbolength.AddItem "30"
   cbolength.AddItem "60"
   cbolength.AddItem "90"
   
 '  cbolength.AddItem "Others"
   MTYPE = 3
   txtboref.SetFocus
   
   Flexitemscin
 Else
   fraTV.Visible = False
   Fraol.Visible = False
   Fracin.Visible = True
   FraMag.Visible = False
   Fraemp.Visible = False
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from cinema_rates where mid(sub_media,1,2)='Sp' order by sub_media"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
        rs.MoveFirst
        cbosubmedia.Clear
        Do Until rs.EOF
         cbosubmedia.AddItem rs!sub_Media
         rs.MoveNext
        Loop
    End If
    
    rs.LockEdits = False
   
   cbodaycin.Clear
   cbodaycin.AddItem "Bi-Weekly"
   cbodaycin.AddItem "Monthly"
   
   cbolength.Clear
   cbolength.AddItem "10 Sl"
   cbolength.AddItem "15 FL"
   cbolength.AddItem "30"
   cbolength.AddItem "60"
   cbolength.AddItem "90"
   
   MTYPE = 3
   txtboref.SetFocus
   
   Flexitemscin
   
 End If
  
End Sub

Private Sub cbomonth_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdate.SetFocus
End Sub

Private Sub cbomonth_LostFocus()

 X = cbomonth.Text
 
 
If X = "January" Then
    DT = 31
ElseIf X = "February" Then
    DT = 28
ElseIf X = "March" Then
    DT = 31
ElseIf X = "April" Then
    DT = 30
ElseIf X = "May" Then
    DT = 31
ElseIf X = "June" Then
    DT = 30
ElseIf X = "July" Then
    DT = 31
ElseIf X = "August" Then
    DT = 31
ElseIf X = "September" Then
    DT = 30
ElseIf X = "October" Then
    DT = 31
ElseIf X = "November" Then
    DT = 30
Else
    DT = 31
End If

m = cbomonth.ListIndex + 1
Y = cboyear.Text
invdate = DT & " / " & m & " / " & Y
invdate = Format(invdate, "dd/mm/yyyy")

End Sub

Private Sub cboProduct_Click()

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from products where Product_Name='" & Trim(cboProduct.Text) & "'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount <> 0 Then
        CboAgency = rs!AGENT_NAME
        CboClient = rs!CLIENT_NAME
        
        txtComPerCin.Text = Val(rs!Discount)
        
        
    End If
    
   
   cboMatCin.Clear
   
             
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = "Select * from material where Product='" & Trim(cboProduct.Text) & "'"
    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If rs1.RecordCount <> 0 Then
        
   
            cbomatmag.Clear
            
                rs1.MoveFirst
                Do Until rs1.EOF
                   cboMatCin.AddItem rs1!Name
                   rs1.MoveNext
                Loop
                CboAgency.SetFocus
        End If
      
 End Sub
Private Sub cboProduct_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboAgency.SetFocus
End Sub
Private Sub cbosubmedia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCinDateFrom.SetFocus
End Sub
Private Sub cbotypecin_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtamountcin.SetFocus
End Sub
Private Sub cboyear_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbomonth.SetFocus
End Sub
Private Sub cboyear_LostFocus()
    Y = cboyear.Text
    m = cbomonth.ListIndex + 1
 If cbomonth.ListIndex = 0 Then
   DT = 31
 ElseIf cbomonth.ListIndex = 1 Then
   DT = 28
 ElseIf cbomonth.ListIndex = 2 Then
   DT = 31
 ElseIf cbomonth.ListIndex = 3 Then
   DT = 30
 ElseIf cbomonth.ListIndex = 4 Then
   DT = 31
 ElseIf cbomonth.ListIndex = 5 Then
   DT = 30
 ElseIf cbomonth.ListIndex = 6 Then
   DT = 31
 ElseIf cbomonth.ListIndex = 7 Then
   DT = 31
 ElseIf cbomonth.ListIndex = 8 Then
   DT = 30
 ElseIf cbomonth.ListIndex = 9 Then
   DT = 31
 ElseIf cbomonth.ListIndex = 10 Then
   DT = 30
 Else
   DT = 31
 End If
 
Y = cboyear.Text
m = cbomonth.ListIndex + 1
invdate = DT & " / " & m & " / " & Y
invdate = Format(invdate, "dd/mm/yyyy")
End Sub

Private Sub cmdadd_Click()
Dim rcount As Currency
Dim addiscpt As Currency
Dim adsurchargept As Currency
Dim agcompt
Dim adcompt

   Y = cboyear.Text
    m = cbomonth.ListIndex + 1
 If cbomonth.ListIndex = 0 Then
   DT = 31
 ElseIf cbomonth.ListIndex = 1 Then
   DT = 28
 ElseIf cbomonth.ListIndex = 2 Then
   DT = 31
 ElseIf cbomonth.ListIndex = 3 Then
   DT = 30
 ElseIf cbomonth.ListIndex = 4 Then
   DT = 31
 ElseIf cbomonth.ListIndex = 5 Then
   DT = 30
 ElseIf cbomonth.ListIndex = 6 Then
   DT = 31
 ElseIf cbomonth.ListIndex = 7 Then
   DT = 31
 ElseIf cbomonth.ListIndex = 8 Then
   DT = 30
 ElseIf cbomonth.ListIndex = 9 Then
   DT = 31
 ElseIf cbomonth.ListIndex = 10 Then
   DT = 30
 Else
   DT = 31
 End If
 
Y = cboyear.Text
m = cbomonth.ListIndex + 1
invdate = DT & " / " & m & " / " & Y
invdate = Format(invdate, "dd/mm/yyyy")

  
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "delete * from Bo_TRAcinprn"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
 
  
 
 If ValidateData = True Then
    If cboCurrency.Text = "DHS" Then
      txtConvRate.Text = 1
    End If
    
On Error GoTo xyz
                  
  If Mid(Trim(cboMediatype.Text), 1, 3) = "Cin" Then
   If MSFlexcin.Rows > 1 Then
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb", adLockOptimistic)
            Sqlqry = " Insert into Bo_Mas values('" & Val(lblserialno.Caption) & "','" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" & Trim(cboCurrency.Text) & "'," & Val(txtConvRate) & "," & Val(Trim(txtGrAmountCin.Text)) & "," & Val(Trim(txtNetAmountCin.Text)) & ",'" & cboyear.Text & "','" _
                                             & Trim(cbomonth.Text) & " '," & Val(cbomonth.ListIndex) & ",'" _
                                             & findfirstfixup(Trim(cboregion.Text)) & "','" & findfirstfixup(Trim(txtremarks.Text)) & "','','" _
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
             
        'db.Close
       
      agdisc = 0
      extdisc = 0
      adddisc = 0
      NOS = 0
      AddDiscEach = 0
       
      agdisc = Val(txtComPerCin.Text)
      extdisc = Val(txtDisccin.Text)
      adddisc = Val(txtAddDiscountCin.Text)
      
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = " select * from Dumbo_tracinbo where type ='Paid'"
'***
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If rs.RecordCount <> 0 Then
    rs.MoveFirst
    Do Until rs.EOF
      NOS = rs.RecordCount
      rs.MoveNext
    Loop
    End If
    If NOS <> "0" Then AddDiscEach = Round(adddisc / NOS, 3)
     
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
                                                    & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                                    & Trim(rs!sub_Media) & "','" _
                                                    & Trim(rs!DATEFROM) & "','" _
                                                    & Trim(rs!Dateto) & "','" _
                                                    & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                                    & Trim(rs!Day) & "','" _
                                                    & Trim(rs!Length) & "','" _
                                                    & findfirstfixup(Trim(rs!Description)) & "','" _
                                                    & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                                    & Trim(rs!Type) & "','" _
                                                    & Trim(rs!tcurrency) & "'," _
                                                    & Trim(rs!tconvertion) & "," _
                                                    & Trim(rs!tra_amount) & "," _
                                                    & Trim(rs!Amount) & "," & 0 & ")"
                           ws.BeginTrans
                           db.Execute (Sqlqry2)
                           ws.CommitTrans
                           
                           
                    Sqlqry2 = " Insert into bo_tracinprn values('" & rs!serial_no & "','" & rs!Year & "','" _
                                             & Trim(rs!Month) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                             & Trim(rs!sub_Media) & "','" _
                                             & Trim(rs!DATEFROM) & "','" _
                                             & Trim(rs!Dateto) & "','" _
                                             & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                             & Trim(rs!Day) & "','" _
                                             & Trim(rs!Length) & "','" _
                                             & findfirstfixup(Trim(rs!Description)) & "','" _
                                             & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                             & Trim(rs!Type) & "','" _
                                             & Trim(rs!tcurrency) & "'," _
                                             & Trim(rs!tconvertion) & "," & 0 & "," & 0 & ")"
                                             '& Trim(rs!tra_amount) & "," _
                                             '& Trim(rs!Amount) & ")"
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
                                                    & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                                    & Trim(rs!sub_Media) & "','" _
                                                    & Trim(rs!DATEFROM) & "','" _
                                                    & Trim(rs!Dateto) & "','" _
                                                    & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                                    & Trim(rs!Day) & "','" _
                                                    & Trim(rs!Length) & "','" _
                                                    & findfirstfixup(Trim(rs!Description)) & "','" _
                                                    & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                                    & Trim(rs!Type) & "','" _
                                                    & Trim(rs!tcurrency) & "'," _
                                                    & Trim(rs!tconvertion) & "," _
                                                    & Trim(rs!tra_amount) & "," _
                                                    & Trim(rs!Amount) & "," & Nettra & ")"
                           ws.BeginTrans
                           db.Execute (Sqlqry2)
                           ws.CommitTrans
                           
                           
                    Sqlqry2 = " Insert into bo_tracinprn values('" & rs!serial_no & "','" & rs!Year & "','" _
                                             & Trim(rs!Month) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                             & Trim(rs!sub_Media) & "','" _
                                             & Trim(rs!DATEFROM) & "','" _
                                             & Trim(rs!Dateto) & "','" _
                                             & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                             & Trim(rs!Day) & "','" _
                                             & Trim(rs!Length) & "','" _
                                             & findfirstfixup(Trim(rs!Description)) & "','" _
                                             & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                             & Trim(rs!Type) & "','" _
                                             & Trim(rs!tcurrency) & "'," _
                                             & Trim(rs!tconvertion) & "," _
                                             & Trim(rs!tra_amount) & "," _
                                             & Trim(rs!Amount) & ")"
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
            End If
            
            
            
          rs.MoveNext
         Loop
       End If
      
        
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Update docu_mas set doc_no='" & lblserialno & "' where doc_type='CIN'"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     
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
      MsgBox "No records found"
      Exit Sub
     End If
   
  textclear
 Else
 
    If MSFlexcin.Rows > 1 Then
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb", adLockOptimistic)
            Sqlqry = " Insert into Bo_Mas values('" & Val(lblserialno.Caption) & "','" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" & Trim(cboCurrency.Text) & "'," & Val(txtConvRate) & "," & Val(Trim(txtGrAmountCin.Text)) & "," & Val(Trim(txtNetAmountCin.Text)) & ",'" & cboyear.Text & "','" _
                                             & Trim(cbomonth.Text) & " '," & Val(cbomonth.ListIndex) & ",'" _
                                             & findfirstfixup(Trim(cboregion.Text)) & "','" & findfirstfixup(Trim(txtremarks.Text)) & "','','" _
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
             
        'db.Close
       
      agdisc = 0
      extdisc = 0
      adddisc = 0
      NOS = 0
      AddDiscEach = 0
       
      agdisc = Val(txtComPerCin.Text)
      extdisc = Val(txtDisccin.Text)
      adddisc = Val(txtAddDiscountCin.Text)
      
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = " select * from Dumbo_tracinbo where type ='Paid'"
'***
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If rs.RecordCount <> 0 Then
    rs.MoveFirst
    Do Until rs.EOF
      NOS = rs.RecordCount
      rs.MoveNext
    Loop
    End If
    If NOS <> "0" Then AddDiscEach = Round(adddisc / NOS, 3)
     
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
                                                    & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                                    & Trim(rs!sub_Media) & "','" _
                                                    & Trim(rs!DATEFROM) & "','" _
                                                    & Trim(rs!Dateto) & "','" _
                                                    & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                                    & Trim(rs!Day) & "','" _
                                                    & Trim(rs!Length) & "','" _
                                                    & findfirstfixup(Trim(rs!Description)) & "','" _
                                                    & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                                    & Trim(rs!Type) & "','" _
                                                    & Trim(rs!tcurrency) & "'," _
                                                    & Trim(rs!tconvertion) & "," _
                                                    & Trim(rs!tra_amount) & "," _
                                                    & Trim(rs!Amount) & "," & 0 & ")"
                           ws.BeginTrans
                           db.Execute (Sqlqry2)
                           ws.CommitTrans
                           
                           
                    Sqlqry2 = " Insert into bo_tracinprn values('" & rs!serial_no & "','" & rs!Year & "','" _
                                             & Trim(rs!Month) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                             & Trim(rs!sub_Media) & "','" _
                                             & Trim(rs!DATEFROM) & "','" _
                                             & Trim(rs!Dateto) & "','" _
                                             & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                             & Trim(rs!Day) & "','" _
                                             & Trim(rs!Length) & "','" _
                                             & findfirstfixup(Trim(rs!Description)) & "','" _
                                             & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                             & Trim(rs!Type) & "','" _
                                             & Trim(rs!tcurrency) & "'," _
                                             & Trim(rs!tconvertion) & "," & 0 & "," & 0 & ")"
                                             '& Trim(rs!tra_amount) & "," _
                                             '& Trim(rs!Amount) & ")"
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
                                                    & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                                    & Trim(rs!sub_Media) & "','" _
                                                    & Trim(rs!DATEFROM) & "','" _
                                                    & Trim(rs!Dateto) & "','" _
                                                    & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                                    & Trim(rs!Day) & "','" _
                                                    & Trim(rs!Length) & "','" _
                                                    & findfirstfixup(Trim(rs!Description)) & "','" _
                                                    & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                                    & Trim(rs!Type) & "','" _
                                                    & Trim(rs!tcurrency) & "'," _
                                                    & Trim(rs!tconvertion) & "," _
                                                    & Trim(rs!tra_amount) & "," _
                                                    & Trim(rs!Amount) & "," & Nettra & ")"
                           ws.BeginTrans
                           db.Execute (Sqlqry2)
                           ws.CommitTrans
                           
                           
                    Sqlqry2 = " Insert into bo_tracinprn values('" & rs!serial_no & "','" & rs!Year & "','" _
                                             & Trim(rs!Month) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                             & Trim(rs!sub_Media) & "','" _
                                             & Trim(rs!DATEFROM) & "','" _
                                             & Trim(rs!Dateto) & "','" _
                                             & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                             & Trim(rs!Day) & "','" _
                                             & Trim(rs!Length) & "','" _
                                             & findfirstfixup(Trim(rs!Description)) & "','" _
                                             & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                             & Trim(rs!Type) & "','" _
                                             & Trim(rs!tcurrency) & "'," _
                                             & Trim(rs!tconvertion) & "," _
                                             & Trim(rs!tra_amount) & "," _
                                             & Trim(rs!Amount) & ")"
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
            End If
            
            
            
          rs.MoveNext
         Loop
       End If
      
        
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Update docu_mas set doc_no='" & lblserialno & "' where doc_type='CIN'"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     
    lblserialno.Caption = Val(lblserialno.Caption) + 1
  
        MsgBox " Record is inserted", vbInformation, "Status"
        X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
        If X = vbYes Then
  
            With CrystalReport1
             .DataFiles(0) = App.Path & "\misov.mdb"
             .ReportFileName = App.Path & "\bocinsp.rpt"
             .WindowState = crptMaximized
             .Action = 1
            End With
        End If
     
     Else
      MsgBox "No records found"
      Exit Sub
     End If
  End If
  textclear

 End If
  Exit Sub
  
xyz:
 MsgBox "Table has been locked exclusively"
 cmdAdd.SetFocus

End Sub

Private Sub cmdBack_Click()
 Unload Me
End Sub
 Private Sub textclear()

   cboProduct.ListIndex = -1
   CboAgency.ListIndex = -1
   CboClient.ListIndex = -1
   txtboref.Text = ""
   cboMediatype.ListIndex = -1

   
 
   cbodaycin.ListIndex = -1
   cbolength.ListIndex = -1
   txtCinDateFrom.TextWithMask = Format(Now(), "dd/mm/yyyy")
   txtCinDateTo.TextWithMask = Format(Now(), "dd/mm/yyyy")
   txtDesccin.Text = ""
   cbosubmedia.ListIndex = -1
   cboMatCin.ListIndex = -1
   cbotypecin.ListIndex = -1
   txtamountcin.Text = ""
        

     txtGrAmountCin.Text = ""

     txtremarks.Text = ""
     cboregion.Text = ""
     

     txtNetAmountCin.Text = ""

     txtAddDiscountCin.Text = ""

     txtComPerCin.Text = ""

     txtDisccin.Text = ""

     txtSurcharge.Text = ""
     
     cboCurrency.ListIndex = -1
     
     lblConvRate.Visible = False
     txtConvRate.Text = ""
     txtConvRate.Visible = False
     
     
     txtfreecin.Text = ""
     
     txtbarterCin.Text = ""
     Flexitemscin
 
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from dumBo_TRAcinbo"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     
          
  End Sub
  
Private Sub cmdClear_Click()
  textclear
End Sub


Private Sub cmdPrint_Click()

End Sub

Private Sub Form_Load()
Dim fmname
Dim fmid

U = 0
Y = 0
Z = 0

MTYPE = 0

fmname = ""
fmname = Me.Caption
fmid = Me.Name

On Error GoTo xyz:

checkin


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
DT = 28
For i = 2000 To 2100
 cboyear.AddItem i
Next
X = 0

 cboyear.Text = Year(Now())
 
 X = Month(Now())
 
 
If X = 1 Then
   cbomonth.ListIndex = 0
   DT = 31
ElseIf X = 2 Then
   cbomonth.ListIndex = 1
   DT = 28
ElseIf X = 3 Then
   cbomonth.ListIndex = 2
   DT = 31
ElseIf X = 4 Then
   cbomonth.ListIndex = 3
   DT = 30
ElseIf X = 5 Then
   cbomonth.ListIndex = 4
   DT = 31
ElseIf X = 6 Then
   cbomonth.ListIndex = 5
   DT = 30
ElseIf X = 7 Then
   cbomonth.ListIndex = 6
   DT = 31
ElseIf X = 8 Then
   cbomonth.ListIndex = 7
   DT = 31
ElseIf X = 9 Then
   cbomonth.ListIndex = 8
   DT = 30
ElseIf X = 10 Then
   cbomonth.ListIndex = 9
   DT = 31
ElseIf X = 11 Then
   cbomonth.ListIndex = 10
   DT = 30
Else
   cbomonth.ListIndex = 11
   DT = 31
End If

txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")

populateproducts
populateagency
populateclient
populateMedia
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

    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumBo_TRAcinbo"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Exit Sub
    
xyz:
 
 MsgBox "Table has been locked exclusively"
    
    
       
End Sub
Private Sub checkin()
Dim uname As String
Dim compname As String
Dim objnet
Dim fmname
Dim fmid

fmname = Me.Caption
fmid = Me.Name

On Error Resume Next

Set objnet = CreateObject("WScript.Network")

If Err.number <> 0 Then
  MsgBox "Error in Getting computer name." & vbCrLf & _
   """No""If your browser warns you."
End If

uname = ""
compname = ""
uname = objnet.UserName
compname = objnet.computername

Set objnet = Nothing

   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = " Select * from formcontrol1 where form_caption='" & Trim(fmname) & "'"
   Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
  ' MsgBox Sqlqry
   If rs.RecordCount = 0 Then
     MsgBox "Form Caption is not matching"
     Exit Sub
   Else
    rs.MoveFirst
'     fmid = ""
     fmid = rs!form_name
     
    If rs!lock_status = "N" Then
       Set ws = DBEngine.Workspaces(0)
       Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       Sqlqry1 = "Update formcontrol1 set " _
                 & " U_Name='" & uname & "'," _
                 & " Comp_Name='" & compname & "'," _
                 & " Lock_status='Y' where form_caption='" & fmname & "'"
       ' MsgBox Sqlqry1
        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
    Else
       
            uname = rs!u_name
            MsgBox "Table has been locked exclusively by the user." & uname
            cmdAdd.Enabled = False
       
    End If
   End If
 
End Sub

Private Sub checkout()

Dim uname As String
Dim compname As String
Dim objnet
On Error Resume Next

fmname = Me.Caption
fmid = Me.Name

Set objnet = CreateObject("WScript.Network")

If Err.number <> 0 Then
  MsgBox "Error in Getting computer name." & vbCrLf & _
   "Do not Press""No""If your browser warns you."
End If

uname = ""
compname = ""
uname = objnet.UserName
compname = objnet.computername

Set objnet = Nothing

   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = " Select * from formcontrol1 where form_caption='" & fmname & "' and u_name='" & uname & "'"
   Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
  ' MsgBox Sqlqry
   If rs.RecordCount <> 0 Then
     rs.MoveFirst
 
     fmid = rs!form_name
        Sqlqry1 = "Update formcontrol1 set " _
                 & " U_Name=''," _
                 & " Comp_Name=''," _
                 & " Lock_status='N' where form_caption='" & fmname & "'"
        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
   End If
 
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
Private Sub populateMedia()
    cboMediatype.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Media Order by Media_Type"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        cboMediatype.AddItem "Cinema"
        cboMediatype.AddItem "Special Operations"
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
    rs.LockEdits = False
    End If
    
 End Sub

Private Sub AutoIncrementVoucher()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select DOC_no from DOCU_MAS WHERE DOC_TYPE='CIN'"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
If rs.RecordCount = 0 Then
   MsgBox "Document type 'CIN' not found"
   Exit Sub
Else
   lblserialno = Val(rs!doc_no) + 1
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
fmname = ""
fmname = Me.Caption
fmid = Me.Name
checkout
End Sub


Private Sub txtAddDiscountCin_LostFocus()
  If txtDisccin.Text = "" Then txtDisccin.Text = 0
  If txtAddDiscountCin.Text = "" Then txtAddDiscountCin.Text = 0
 txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100) - (Val(Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)) * txtDisccin / 100) - Val(txtAddDiscountCin.Text)
 cmdAdd.SetFocus
End Sub
Private Sub txtadddiscountCIN_KeyPress(KeyAscii As Integer)
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
    ' ******
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset, adlockonly)
     If rs.RecordCount = 0 Then
        Sqlqry2 = " Insert into dumBo_TRAcinbo values('" & Val(lblserialno.Caption) & "','" & cboyear.Text & "','" _
                                     & cbomonth.Text & "','" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Cinema','" _
                                     & Trim(cbosubmedia.Text) & "','" _
                                     & Format(txtCinDateFrom.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & Format(txtCinDateTo.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "','" _
                                     & Trim(cbodaycin.Text) & "','" _
                                     & Trim(cbolength.Text) & "','" _
                                     & findfirstfixup(Trim(txtDesccin.Text)) & "','" _
                                     & findfirstfixup(Trim(cboMatCin.Text)) & "','" _
                                     & Trim(cbotypecin.Text) & "','" _
                                     & Trim(cboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(Trim(txtamountcin.Text)) & ", " & Val(Trim(txtamountcin.Text)) * Val(txtConvRate.Text) & ")"

'MsgBox Sqlqry1

        ws.BeginTrans
        db.Execute (Sqlqry2)
        ws.CommitTrans
        
        
        sqlqry3 = "select * from dumBo_TRAcinbo"
       ' *******
        Set rs3 = db.OpenRecordset(sqlqry3, dbOpenDynaset)
        If rs3.RecordCount = 0 Then
            MSFlexcin.Clear
            Exit Sub
        Else
            Flexitemscin
            rs3.MoveFirst
            Do Until rs3.EOF
              MSFlexcin.AddItem rs3!sub_Media & Chr(9) & Format(rs3!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs3!Dateto, "dd/mm/yyyy") & Chr(9) & rs3!Day & Chr(9) & rs3!Length & Chr(9) & rs3!Description & Chr(9) & rs3!mat_code & Chr(9) & rs3!Type & Chr(9) & rs3!tra_amount
              rs3.MoveNext
            Loop
            rs3.LockEdits = False
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
         Sqlqry2 = " Insert into dumBo_TRAcinbo values('" & Val(lblserialno.Caption) & "','" & cboyear.Text & "','" _
                                     & cbomonth.Text & "','" _
                                     & findfirstfixup(cboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Cinema','" _
                                     & Trim(cbosubmedia) & "','" _
                                     & Format(txtCinDateFrom.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & Format(txtCinDateTo.TextWithMask, "dd/mm/yyyy") & "','" _
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
        db.Execute (Sqlqry2)
        ws.CommitTrans
       
        
        sqlqry3 = "select * from dumBo_TRAcinbo"
        ' *****
        Set rs3 = db.OpenRecordset(sqlqry3, dbOpenDynaset)
        If rs3.RecordCount = 0 Then
            MSFlexcin.Clear
            Exit Sub
        Else
            Flexitemscin
            rs3.MoveFirst
            Do Until rs3.EOF
               MSFlexcin.AddItem rs3!sub_Media & Chr(9) & Format(rs3!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs3!Dateto, "dd/mm/yyyy") & Chr(9) & rs3!Day & Chr(9) & rs3!Length & Chr(9) & rs3!Description & Chr(9) & rs3!mat_code & Chr(9) & rs3!Type & Chr(9) & rs3!tra_amount
               rs3.MoveNext
            Loop
               rs3.LockEdits = False
        End If
        
        Sqlqry2 = "select sum(tra_amount) from dumBo_TRAcinbo where type='Paid'"
        Set rs1 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then U = rs1.Fields(0)
            
        Sqlqry2 = "select sum(tra_amount) from dumBo_TRAcinbo where type='Free'"
        Set rs1 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Y = rs1.Fields(0)
            
        Sqlqry2 = "select sum(tra_amount) from dumBo_TRAcinbo where type='Barter'"
        Set rs1 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Z = rs1.Fields(0)
            
        rs1.LockEdits = False
            
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

Private Sub txtboref_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboregion.SetFocus
End Sub
Private Sub cbodaycin_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbolength.SetFocus
End Sub
Private Sub txtCinDateFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCinDateTo.SetFocus
End Sub

Private Sub txtCinDateTo_GotFocus()
    Dim X As Date
    If Mid(txtCinDateFrom.TextWithMask, 4, 2) > 12 Then
      MsgBox "Invalid cinema Date from", vbInformation, "Invalid Entry"
      txtCinDateFrom.SetFocus
      SendKeys " {Home} + {End} "
    End If
     If IsDate(X) = False Then
       MsgBox "Invalid Cinema Date From", vbInformation, "Invalid Entry"
       txtCinDateFrom.SetFocus
       SendKeys " {Home} + {End} "
    End If
End Sub

Private Sub txtCinDateTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cbodaycin.SetFocus
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
Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdateto.SetFocus
End Sub

Private Sub txtdatefrom_LostFocus()
If IsDate(txtdatefrom.TextWithMask) = False Then
   MsgBox "Invalid Date From", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdescol.SetFocus
End Sub

Private Sub txtdateto_LostFocus()
If IsDate(txtdateto.TextWithMask) = False Then
   MsgBox "Invalid Date To", vbInformation, "Invalid Entry"
   txtdateto.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub
Private Sub txtDesccin_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboMatCin.SetFocus
End Sub

Private Sub CboRegion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtremarks.SetFocus
End Sub
Private Sub txtremarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If MTYPE = 1 Then
  txtDaytv.SetFocus
 ElseIf MTYPE = 2 Then
  txtdatefrom.SetFocus
 ElseIf MTYPE = 3 Then
  cbosubmedia.SetFocus
 ElseIf MTYPE = 4 Then
  txtissueno.SetFocus
 End If
End If
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
        txtCinDateFrom.TextWithMask = Trim(.Text)
        .Col = 2
        txtCinDateTo.TextWithMask = Trim(.Text)
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
                             
                
       'Sqlqry1 = "select * from dumBo_TRAcinbo where  sub_media ='" & cbosubmedia & "' and description ='" & Trim(txtDescCin) & "' and Day ='" & Trim(cbodaycin) & "' and Length ='" & Trim(cbolength) & "' and tra_amount =" & Val(txtamountcin) & " AND TYPE='" & Trim(cbotypecin.Text) & "' AND DateFrom=#" & DateValue(Format(txtCinDateFrom.TextWithMask, "dd/MM/YYYY")) & "# AND DateTo=#" & DateValue(Format(txtCinDateTo.TextWithMask, "dd/MM/YYYY")) & "# "
        Sqlqry1 = "select * from dumBo_TRAcinbo where  sub_media ='" & cbosubmedia & "' and description ='" & Trim(txtDesccin) & "' and Day ='" & Trim(cbodaycin) & "' and Length ='" & Trim(cbolength) & "' and tra_amount =" & Val(txtamountcin) & " AND TYPE='" & Trim(cbotypecin.Text) & "' AND DateFrom=#" & DateValue(Format(txtCinDateFrom.TextWithMask, "dd/mm/yyyy")) & "# AND DateTo=#" & DateValue(Format(txtCinDateTo.TextWithMask, "dd/mm/yyyy")) & "# "
     '  MsgBox Sqlqry1
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset, adlockonly)
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

Private Sub txtSurcharge_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub

