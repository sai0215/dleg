VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmBOModCin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "BOModificationCinema"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cmdviewclear 
      Caption         =   "Clear"
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
      Left            =   9720
      TabIndex        =   188
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox CboIssue 
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
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
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
      Left            =   10800
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox CboSMedia 
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
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   3135
   End
   Begin VB.ComboBox CboSProduct 
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
      Left            =   9120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.ComboBox CboSmonth 
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
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox CboSyear 
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
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox CboSAgency 
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
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancell 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Cancelled Orders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   5160
      Picture         =   "frmBOModCin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7800
      Width           =   1335
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2160
      Top             =   7800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0FFC0&
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
      Height          =   780
      Left            =   4200
      Picture         =   "frmBOModCin.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7800
      Width           =   975
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
      Height          =   780
      Left            =   7440
      Picture         =   "frmBOModCin.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7800
      Width           =   1095
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
      Height          =   780
      Left            =   6480
      Picture         =   "frmBOModCin.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdModify 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Modify"
      DisabledPicture =   "frmBOModCin.frx":0748
      DownPicture     =   "frmBOModCin.frx":0C7A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   3240
      MaskColor       =   &H008080FF&
      Picture         =   "frmBOModCin.frx":11AC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   2655
      Left            =   240
      TabIndex        =   149
      Top             =   960
      Width           =   11655
      Begin VB.Frame Fradata 
         BackColor       =   &H00FFFFFF&
         Height          =   2655
         Left            =   0
         TabIndex        =   150
         Top             =   0
         Width           =   11655
         Begin VB.ComboBox CboAgency 
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
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   190
            Top             =   1800
            Width           =   3255
         End
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
            Left            =   7680
            Style           =   2  'Dropdown List
            TabIndex        =   189
            Top             =   1680
            Width           =   3855
         End
         Begin VB.ComboBox cboMonth 
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
            Left            =   6720
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox Cboyear 
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
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox CboProduct 
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
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   159
            Top             =   1320
            Width           =   3255
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
            Left            =   1200
            TabIndex        =   158
            Top             =   2160
            Width           =   1215
         End
         Begin VB.ComboBox cboMediaType 
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
            Left            =   7680
            Style           =   2  'Dropdown List
            TabIndex        =   157
            Top             =   1200
            Width           =   3855
         End
         Begin VB.ComboBox CboCurrency 
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
            Left            =   6720
            Style           =   2  'Dropdown List
            TabIndex        =   156
            Top             =   720
            Width           =   1695
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
            Height          =   345
            Left            =   10080
            TabIndex        =   155
            Top             =   750
            Width           =   735
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Height          =   1695
            Left            =   240
            TabIndex        =   154
            Top             =   240
            Width           =   2175
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
               Height          =   1500
               Left            =   120
               TabIndex        =   7
               Top             =   120
               Width           =   1935
            End
         End
         Begin VB.CheckBox optcancel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cancelled"
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
            Left            =   9120
            TabIndex        =   153
            Top             =   240
            Width           =   1575
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
            Left            =   3720
            TabIndex        =   152
            Top             =   2280
            Width           =   2415
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
            Height          =   330
            Left            =   7680
            TabIndex        =   151
            Top             =   2205
            Width           =   3855
         End
         Begin PVMaskEditLib.PVMaskEdit txtdate 
            Height          =   255
            Left            =   3720
            TabIndex        =   160
            Top             =   840
            Width           =   1335
            _Version        =   65541
            _ExtentX        =   2355
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
            Height          =   255
            Left            =   5520
            TabIndex        =   172
            Top             =   360
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
            Height          =   255
            Left            =   2880
            TabIndex        =   171
            Top             =   360
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
            Height          =   255
            Left            =   2640
            TabIndex        =   170
            Top             =   1320
            Width           =   975
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
            Height          =   255
            Left            =   7080
            TabIndex        =   169
            Top             =   1800
            Width           =   615
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
            Height          =   375
            Left            =   2640
            TabIndex        =   168
            Top             =   1800
            Width           =   975
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
            Height          =   255
            Left            =   240
            TabIndex        =   167
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Media"
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
            Left            =   6960
            TabIndex        =   166
            Top             =   1320
            Width           =   735
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
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   5640
            TabIndex        =   165
            Top             =   840
            Width           =   930
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
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   8820
            TabIndex        =   164
            Top             =   840
            Width           =   1125
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
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   3120
            TabIndex        =   163
            Top             =   840
            Width           =   510
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
            Left            =   6720
            TabIndex        =   162
            Top             =   2280
            Width           =   975
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
            Left            =   2880
            TabIndex        =   161
            Top             =   2280
            Width           =   855
         End
      End
      Begin VB.Frame FraView 
         BackColor       =   &H00FFFFFF&
         Height          =   2655
         Left            =   0
         TabIndex        =   173
         Top             =   0
         Width           =   11655
         Begin MSFlexGridLib.MSFlexGrid MSFlexview 
            Height          =   2295
            Left            =   120
            TabIndex        =   180
            Top             =   120
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   4048
            _Version        =   393216
            Cols            =   8
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
      End
   End
   Begin VB.Frame Framain 
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   240
      TabIndex        =   15
      Top             =   3720
      Width           =   11655
      Begin VB.Frame Fracin 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   3735
         Left            =   120
         TabIndex        =   117
         Top             =   120
         Width           =   11415
         Begin VB.TextBox txtDescCin 
            Height          =   285
            Left            =   6840
            TabIndex        =   124
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtamountcin 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10440
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   600
            Width           =   855
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
            TabIndex        =   126
            Top             =   600
            Width           =   855
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
            TabIndex        =   125
            Top             =   600
            Width           =   1215
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
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1335
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
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1335
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
            TabIndex        =   131
            Top             =   3240
            Width           =   1215
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
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   3240
            Width           =   495
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
            TabIndex        =   129
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1215
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
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1215
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
            TabIndex        =   118
            Top             =   600
            Width           =   2655
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
            TabIndex        =   122
            Text            =   "cbodaycin"
            Top             =   600
            Width           =   855
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
            Left            =   6000
            TabIndex        =   123
            Top             =   600
            Width           =   735
         End
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
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   3240
            Width           =   495
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexcin 
            Height          =   1695
            Left            =   120
            TabIndex        =   134
            Top             =   960
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   2990
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
         Begin PVMaskEditLib.PVMaskEdit txtCinDateFrom 
            Height          =   255
            Left            =   2880
            TabIndex        =   119
            Top             =   600
            Width           =   975
            _Version        =   65541
            _ExtentX        =   1720
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
            Left            =   3960
            TabIndex        =   120
            Top             =   600
            Width           =   975
            _Version        =   65541
            _ExtentX        =   1720
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
         Begin VB.Label Label79 
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
            Left            =   6840
            TabIndex        =   187
            Top             =   360
            Width           =   1095
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
            Left            =   3960
            TabIndex        =   186
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label18 
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
            Left            =   10320
            TabIndex        =   148
            Top             =   360
            Width           =   975
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
            TabIndex        =   147
            Top             =   360
            Width           =   735
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
            Left            =   9360
            TabIndex        =   146
            Top             =   360
            Width           =   495
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
            Left            =   6000
            TabIndex        =   145
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label27 
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
            Left            =   2760
            TabIndex        =   144
            Top             =   360
            Width           =   975
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
            TabIndex        =   143
            Top             =   360
            Width           =   615
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00004080&
            X1              =   0
            X2              =   11400
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agency Commission"
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
            TabIndex        =   142
            Top             =   3240
            Width           =   2295
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
            TabIndex        =   141
            Top             =   2760
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
            TabIndex        =   140
            Top             =   3240
            Width           =   1455
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
            TabIndex        =   139
            Top             =   3240
            Width           =   1455
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
            TabIndex        =   137
            Top             =   2760
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
            TabIndex        =   136
            Top             =   360
            Width           =   1335
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
            TabIndex        =   135
            Top             =   3240
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
            TabIndex        =   138
            Top             =   2760
            Width           =   1335
         End
      End
      Begin VB.Frame Fraol 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   3735
         Left            =   120
         TabIndex        =   85
         Top             =   120
         Width           =   11415
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
            TabIndex        =   98
            Top             =   600
            Width           =   975
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
            TabIndex        =   97
            Top             =   600
            Width           =   2295
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
            TabIndex        =   96
            Top             =   600
            Width           =   1695
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
            TabIndex        =   95
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
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   600
            Width           =   1095
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
            TabIndex        =   93
            Top             =   600
            Width           =   975
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
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1335
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
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1335
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
            TabIndex        =   90
            Top             =   3240
            Width           =   1095
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
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   3240
            Width           =   495
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
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1215
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
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1215
         End
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
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   3240
            Width           =   495
         End
         Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
            Height          =   255
            Left            =   120
            TabIndex        =   99
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
            TabIndex        =   100
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
         Begin MSFlexGridLib.MSFlexGrid MSFlexol 
            Height          =   1455
            Left            =   240
            TabIndex        =   101
            Top             =   1080
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   2566
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
            TabIndex        =   116
            Top             =   240
            Width           =   975
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
            TabIndex        =   114
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
            TabIndex        =   113
            Top             =   240
            Width           =   735
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
            TabIndex        =   112
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Amount "
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
            TabIndex        =   111
            Top             =   240
            Width           =   1095
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
            TabIndex        =   110
            Top             =   120
            Width           =   615
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00004080&
            X1              =   0
            X2              =   11400
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agency Commission"
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
            Left            =   4920
            TabIndex        =   109
            Top             =   3240
            Width           =   2175
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
            TabIndex        =   108
            Top             =   2760
            Width           =   1455
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
            TabIndex        =   107
            Top             =   3240
            Width           =   1455
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
            TabIndex        =   106
            Top             =   3240
            Width           =   1815
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
            TabIndex        =   105
            Top             =   2760
            Width           =   1215
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
            TabIndex        =   104
            Top             =   2760
            Width           =   1335
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
            TabIndex        =   103
            Top             =   3240
            Width           =   1455
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
            TabIndex        =   102
            Top             =   240
            Width           =   855
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
            TabIndex        =   115
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame FraMag 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   3735
         Left            =   120
         TabIndex        =   16
         Top             =   120
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
            Height          =   255
            Left            =   6360
            TabIndex        =   19
            Top             =   480
            Width           =   1215
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
            Height          =   240
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   32
            Top             =   480
            Width           =   375
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
            Left            =   7680
            TabIndex        =   20
            Top             =   480
            Width           =   1695
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
            Height          =   255
            Left            =   4200
            TabIndex        =   18
            Top             =   480
            Width           =   2055
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
            TabIndex        =   31
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtamountmag 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10560
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   480
            Width           =   735
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
            Height          =   240
            Left            =   120
            MaxLength       =   5
            TabIndex        =   29
            Top             =   480
            Width           =   375
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
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   3240
            Width           =   495
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
            TabIndex        =   27
            Top             =   3240
            Width           =   975
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
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1335
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
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1335
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
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1215
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
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1215
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
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   3240
            Width           =   495
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
            TabIndex        =   21
            Top             =   3240
            Width           =   975
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
            TabIndex        =   17
            Top             =   480
            Width           =   1815
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexmag 
            Height          =   1815
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   3201
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
            TabIndex        =   34
            Top             =   480
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
            HighlightColor  =   12632256
            Alignment       =   1
         End
         Begin VB.Label Label77 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mat. Stat."
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
            Left            =   6360
            TabIndex        =   185
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label22 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Position"
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
            Left            =   4200
            TabIndex        =   50
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label24 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Iss #"
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
            Left            =   120
            TabIndex        =   49
            Top             =   240
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
            TabIndex        =   48
            Top             =   240
            Width           =   495
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
            TabIndex        =   47
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label19 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mat. Copy"
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
            Left            =   7680
            TabIndex        =   46
            Top             =   240
            Width           =   975
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
            TabIndex        =   45
            Top             =   240
            Width           =   1215
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
            TabIndex        =   44
            Top             =   3240
            Width           =   1815
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
            TabIndex        =   43
            Top             =   3240
            Width           =   1575
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
            TabIndex        =   42
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Commission"
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
            Left            =   3960
            TabIndex        =   41
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00004080&
            X1              =   0
            X2              =   11400
            Y1              =   3120
            Y2              =   3120
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
            TabIndex        =   40
            Top             =   2760
            Width           =   1215
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
            TabIndex        =   39
            Top             =   2760
            Width           =   1335
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
            TabIndex        =   38
            Top             =   3240
            Width           =   975
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
            TabIndex        =   37
            Top             =   3240
            Width           =   1095
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
            TabIndex        =   36
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   720
            TabIndex        =   35
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.Frame fraTV 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   3735
         Left            =   120
         TabIndex        =   51
         Top             =   120
         Width           =   11415
         Begin VB.TextBox txtSEc 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7800
            MaxLength       =   3
            TabIndex        =   64
            Top             =   720
            Width           =   495
         End
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
            Left            =   3240
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox cbotypetv 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   360
            Left            =   6600
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   720
            Width           =   1095
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
            TabIndex        =   62
            Top             =   720
            Width           =   1695
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
            Height          =   405
            Left            =   1440
            TabIndex        =   60
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtSpots 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8520
            MaxLength       =   3
            TabIndex        =   66
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtRatetv 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9120
            TabIndex        =   68
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtamounttv 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10200
            Locked          =   -1  'True
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtDaytv 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MaxLength       =   5
            TabIndex        =   59
            Top             =   720
            Width           =   495
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
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   3360
            Width           =   1335
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
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1335
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
            TabIndex        =   56
            Top             =   3360
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
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   3360
            Width           =   495
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
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1215
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
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1215
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
            Left            =   4200
            MaxLength       =   2
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   3360
            Width           =   495
         End
         Begin PVMaskEditLib.PVMaskEdit PVMaskTime 
            Height          =   375
            Left            =   720
            TabIndex        =   65
            Top             =   720
            Width           =   615
            _Version        =   65541
            _ExtentX        =   1085
            _ExtentY        =   661
            _StockProps     =   244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
         Begin MSFlexGridLib.MSFlexGrid MSFlextv 
            Height          =   1455
            Left            =   240
            TabIndex        =   67
            Top             =   1200
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   2566
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
         Begin VB.Label Label82 
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
            Left            =   7680
            TabIndex        =   193
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label81 
            Alignment       =   1  'Right Justify
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
            Height          =   375
            Left            =   3240
            TabIndex        =   192
            Top             =   480
            Width           =   615
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00004080&
            X1              =   0
            X2              =   11400
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agency Commission"
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
            TabIndex        =   76
            Top             =   3360
            Width           =   2175
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
            TabIndex        =   75
            Top             =   2760
            Width           =   1335
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
            TabIndex        =   74
            Top             =   3360
            Width           =   1095
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
            TabIndex        =   73
            Top             =   3360
            Width           =   1815
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
            TabIndex        =   72
            Top             =   2760
            Width           =   1215
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
            TabIndex        =   71
            Top             =   2760
            Width           =   1335
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
            Left            =   2640
            TabIndex        =   69
            Top             =   3360
            Width           =   1455
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
            Left            =   8400
            TabIndex        =   84
            Top             =   480
            Width           =   735
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
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   480
            Width           =   495
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
            TabIndex        =   82
            Top             =   480
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
            Height          =   375
            Left            =   1440
            TabIndex        =   81
            Top             =   480
            Width           =   1335
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
            Left            =   720
            TabIndex        =   80
            Top             =   480
            Width           =   615
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
            Left            =   6600
            TabIndex        =   79
            Top             =   480
            Width           =   615
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
            Height          =   375
            Left            =   4920
            TabIndex        =   78
            Top             =   480
            Width           =   855
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
            Height          =   375
            Left            =   10080
            TabIndex        =   77
            Top             =   480
            Width           =   975
         End
         Begin VB.Label LblCode 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   3960
            TabIndex        =   191
            Top             =   120
            Width           =   7095
         End
      End
      Begin VB.Frame Fraemp 
         BackColor       =   &H00FFFFFF&
         Height          =   3735
         Left            =   120
         TabIndex        =   183
         Top             =   120
         Width           =   11415
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   3495
            Left            =   0
            TabIndex        =   184
            Top             =   120
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   6165
            _Version        =   393216
            Rows            =   15
            Cols            =   12
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
      End
   End
   Begin VB.Label LblviewSubmedia 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6960
      TabIndex        =   182
      Top             =   600
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblviewMedia 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6720
      TabIndex        =   181
      Top             =   600
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label LblIss 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Issue #"
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
      Left            =   4080
      TabIndex        =   179
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label76 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Media"
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
      Left            =   0
      TabIndex        =   178
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label75 
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
      Height          =   255
      Left            =   1800
      TabIndex        =   177
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label74 
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
      Height          =   255
      Left            =   120
      TabIndex        =   176
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label73 
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
      Height          =   255
      Left            =   8160
      TabIndex        =   175
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label72 
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
      Height          =   255
      Left            =   4200
      TabIndex        =   174
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmBOModCin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public thirtysecrate As Currency
Dim i As Integer
Dim DT As Long
Dim mm As Long
Dim Y As Long
Dim MTYPE
Dim med As String
Dim SNo
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim X, Z As Integer
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim ws As Workspace
Dim invdate As Date
Dim acdatefrom As Date
Dim acdateto As Date
Dim xxx
Dim l, o, p As String
Dim n, m As String
Dim agdisc As Currency
Dim extdisc As Currency
Dim adddisc As Currency
Dim AddDiscEach As Currency
Dim Nettra As Currency
Dim NOS As Integer
Private Sub PopulateVoucher()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select * from Bo_Mas where status='N' and Media='Cinema' ORDER BY val(Serial_NO)"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
lstVoucNo.Clear
If rs.RecordCount <> 0 Then
   rs.MoveFirst
   Do Until rs.EOF
       lstVoucNo.AddItem rs!serial_no
       rs.MoveNext
   Loop
End If
End Sub
Private Sub CboAgency_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboClient.SetFocus
End Sub
Private Sub CboClient_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtboref.SetFocus
End Sub
Private Sub CboCode_Change()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from cnnrates where Code='" & Trim(CboCode.Text) & "'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    LblCode.Caption = rs!region & " : " & rs!wtype & " : " & rs!sp_Prog & " : " & rs!ttime & " : " & rs!quarter & " : " & rs!Rate
End Sub
Private Sub CboCode_Click()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from cnnrates where Code='" & Trim(CboCode.Text) & "'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    LblCode.Caption = rs!region & " : " & rs!wtype & " : " & rs!sp_Prog & " : " & rs!ttime & " : " & rs!quarter & " : " & rs!Rate
End Sub
Private Sub CboCode_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cboMattv.SetFocus
End Sub
Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     If CboCurrency.Text = "USD" Then
      lblConvRate.Visible = True
      txtConvRate.Visible = True
      txtConvRate.Text = ""
      txtConvRate.SetFocus
     Else
      lblConvRate.Visible = False
      txtConvRate.Visible = False
      txtConvRate.Text = 1
      CboProduct.SetFocus
     End If
    End If
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
     txtConvRate.TabIndex = 38
 End If
End Sub
Private Sub cbodaycin_GotFocus()
    If Mid(txtCinDateTo.TextWithMask, 4, 2) > 12 Then
          MsgBox "Invalid cinema Date to", vbInformation, "Invalid Entry"
          txtCinDateTo.SetFocus
          SendKeys " {Home} + {End} "
    End If

End Sub
Private Sub CboIssue_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdView.SetFocus
End Sub

Private Sub cbolength_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtDescCin.SetFocus
End Sub

Private Sub cbolength_LostFocus()
 Dim C As String
 Dim X As Integer
 Dim Y As Currency
 C = ""
 X = 0
 Y = txtamountcin.Text
 
  
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
 If Mid(cboMediaType.Text, 1, 3) = "Tel" Then
   MTYPE = 1
 ElseIf Mid(cboMediaType.Text, 1, 3) = "Onl" Then
   MTYPE = 2
 ElseIf Mid(cboMediaType.Text, 1, 3) = "Cin" Then
   MTYPE = 3
 ElseIf Mid(cboMediaType.Text, 1, 3) = "Spe" Then
   MTYPE = 3
 ElseIf Mid(cboMediaType.Text, 1, 3) = "Mag" Then
   MTYPE = 4
Else
   fraTV.Visible = False
   Fraol.Visible = False
   Fracin.Visible = False
   FraMag.Visible = False
   Fraemp.Visible = True
End If
End Sub
Private Sub cboMediaType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboAgency.SetFocus
End Sub
Private Sub cboMediatype_LostFocus()
If Mid(cboMediaType.Text, 1, 3) = "Cin" Then
   fraTV.Visible = False
   Fraol.Visible = False
   Fracin.Visible = True
   FraMag.Visible = False
   Fraemp.Visible = False
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from cinema_rates where mid(sub_media,1,2) <>'Sp' order by sub_media"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
        rs.MoveFirst
        Do Until rs.EOF
         cbosubmedia.AddItem rs!sub_Media
         rs.MoveNext
        Loop
    End If
   
   cbosubmedia.SetFocus
   cbodaycin.AddItem "Bi-Weekly"
   cbodaycin.AddItem "Monthly"
   
   cbolength.AddItem "10 Sl"
   cbolength.AddItem "15 FL"
   cbolength.AddItem "30"
   cbolength.AddItem "60"
   cbolength.AddItem "90"
   
   MTYPE = 3
   txtboref.SetFocus
   
   Flexitemscin

  ''''
 ElseIf Mid(cboMediaType.Text, 1, 3) = "Spe" Then
   fraTV.Visible = False
   Fraol.Visible = False
   Fracin.Visible = True
   FraMag.Visible = False
   Fraemp.Visible = False
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from cinema_rates where mid(sub_media,1,2) ='Sp' order by sub_media"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
        rs.MoveFirst
        Do Until rs.EOF
         cbosubmedia.AddItem rs!sub_Media
         rs.MoveNext
        Loop
    End If
   
   cbosubmedia.SetFocus
   cbodaycin.AddItem "Bi-Weekly"
   cbodaycin.AddItem "Monthly"
   
   cbolength.AddItem "10 Sl"
   cbolength.AddItem "15 FL"
   cbolength.AddItem "30"
   cbolength.AddItem "60"
   cbolength.AddItem "90"
   
   MTYPE = 3
   txtboref.SetFocus
   
   Flexitemscin

Else
   fraTV.Visible = False
   Fraol.Visible = False
   Fracin.Visible = False
   FraMag.Visible = False
   Fraemp.Visible = True
 End If
End Sub
Private Sub cbomonth_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then optcancel.SetFocus
End Sub

Private Sub cbomonth_LostFocus()

 X = cboMonth.Text
 
 
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

mm = cboMonth.ListIndex + 1
Y = Cboyear.Text
invdate = DT & " / " & mm & " / " & Y
invdate = Format(invdate, "dd/mm/yyyy")
End Sub

Private Sub cboProduct_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboMediaType.SetFocus
End Sub

Private Sub CboProduct_LostFocus()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from products where Product_Name='" & Trim(CboProduct.Text) & "'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount <> 0 Then
        CboAgency = rs!AGENT_NAME
        CboClient = rs!CLIENT_NAME
        txtComPerCin.Text = Val(rs!Discount)
    End If
   cboMatCin.Clear
             
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = "Select * from material where Product='" & Trim(CboProduct.Text) & "'"
    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If rs1.RecordCount <> 0 Then
        
          
            cboMatCin.Clear
                      
                rs1.MoveFirst
                
                Do Until rs1.EOF
                  cboMatCin.AddItem rs1!Name
                 rs1.MoveNext
                Loop
              
    End If
End Sub
Private Sub CboRegion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtremarks.SetFocus
End Sub
Private Sub CboSAgency_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CboSProduct.SetFocus
End Sub
Private Sub CboSmedia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdView.SetFocus
End Sub
Private Sub CboSmedia_LostFocus()
 lblviewMedia.Caption = ""
 LblviewSubmedia.Caption = ""
If CboSMedia.Text = "Cinema" Then
   lblviewMedia.Caption = "Cinema"
   LblviewSubmedia.Caption = ""
ElseIf Mid(CboSMedia, 1, 3) = "Cin" Then
   lblviewMedia.Caption = "Cinema"
   LblviewSubmedia.Caption = Trim(Mid(CboSMedia, 8, 30))
End If
End Sub
Private Sub CboSMonth_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboSAgency.SetFocus
End Sub
Private Sub cbospace_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdescmag.SetFocus
End Sub
Private Sub CboSProduct_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboSMedia.SetFocus
End Sub
Private Sub cbosubmedia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtCinDateFrom.SetFocus
End Sub
Private Sub CboSYear_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboSmonth.SetFocus
End Sub
Private Sub cbotypecin_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtamountcin.SetFocus
End Sub
Private Sub cboyear_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboMonth.SetFocus
End Sub

Private Sub cboyear_LostFocus()
Y = Cboyear.Text
mm = cboMonth.ListIndex + 1
 If cboMonth.ListIndex = 0 Then
   DT = 31
 ElseIf cboMonth.ListIndex = 1 Then
   DT = 28
 ElseIf cboMonth.ListIndex = 2 Then
   DT = 31
 ElseIf cboMonth.ListIndex = 3 Then
   DT = 30
 ElseIf cboMonth.ListIndex = 4 Then
   DT = 31
 ElseIf cboMonth.ListIndex = 5 Then
   DT = 30
 ElseIf cboMonth.ListIndex = 6 Then
   DT = 31
 ElseIf cboMonth.ListIndex = 7 Then
   DT = 31
 ElseIf cboMonth.ListIndex = 8 Then
   DT = 30
 ElseIf cboMonth.ListIndex = 9 Then
   DT = 31
 ElseIf cboMonth.ListIndex = 10 Then
   DT = 30
 Else
   DT = 31
 End If
 
Y = Cboyear.Text
mm = cboMonth.ListIndex + 1
invdate = DT & " / " & mm & " / " & Y
invdate = Format(invdate, "dd/mm/yyyy")
End Sub

Private Sub cmdCancell_Click()
   With CrystalReport1
          .DataFiles(0) = App.Path & "\misov.mdb"
          .ReportFileName = App.Path & "\bocancell.rpt"
          .WindowState = crptMaximized
          .Action = 1
    End With
End Sub

Private Sub CmdPrint_Click()
  SNo = 0
  SNo = Val(lstVoucNo)
  
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "delete * from Bo_TRAcinprn"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
  
 If Mid(Trim(cboMediaType.Text), 1, 3) = "Cin" Then
         
    Sqlqry1 = "Select * from Bo_TRAcin where serial_no='" & Val(lstVoucNo) & "'"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount <> 0 Then
         
         rs.MoveFirst
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           
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
          rs.MoveNext
         Loop
       End If
         
         With CrystalReport1
          .DataFiles(0) = App.Path & "\misov.mdb"
          .ReportFileName = App.Path & "\bocin.rpt"
          .WindowState = crptMaximized
          .Action = 1
         End With
   Else
          
    Sqlqry1 = "Select * from Bo_TRAcin where serial_no='" & Val(lstVoucNo) & "'"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount <> 0 Then
         
         rs.MoveFirst
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           
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
          rs.MoveNext
         Loop
       End If
         
         With CrystalReport1
          .DataFiles(0) = App.Path & "\misov.mdb"
          .ReportFileName = App.Path & "\bocinsp.rpt"
          .WindowState = crptMaximized
          .Action = 1
         End With
  End If
  
End Sub
Private Sub cmdBack_Click()
 Unload Me
End Sub
Private Sub textclear()
    
   CboProduct.ListIndex = -1
   optcancel.Value = 0
   CboAgency.ListIndex = -1
   CboClient.ListIndex = -1
   txtboref.Text = ""
   cboMediaType.ListIndex = -1
 
 
   cbodaycin.ListIndex = -1
   cbolength.ListIndex = -1
   txtDescCin.Text = ""
   txtCinDateFrom.TextWithMask = Format(Now, "DD/MM/YYYY")
   txtCinDateTo.TextWithMask = Format(Now, "DD/MM/YYYY")
   cbosubmedia.ListIndex = -1
   cboMatCin.ListIndex = -1
   cbotypecin.ListIndex = -1
   txtamountcin.Text = ""
       
   
   txtremarks.Text = ""
   cboregion.Text = ""
         
     txtGrAmountCin.Text = ""
     txtNetAmountCin.Text = ""
     txtAddDiscountCin.Text = ""
     txtComPerCin.Text = ""
     txtDisccin.Text = ""
 
     CboCurrency.ListIndex = -1
     lblConvRate.Visible = False
     txtConvRate.Text = ""
     txtConvRate.Visible = False
     
 
     txtfreecin.Text = ""
     txtbarterCin.Text = ""
     Flexitemscin
 
    
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from Dumbo_tracinbomod"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
End Sub
Private Sub textclear1()
   CboProduct.ListIndex = -1
   optcancel.Value = 0
   txtboref.Text = ""
   cboMediaType.ListIndex = -1
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
   txtDescCin.Text = ""
   cbosubmedia.ListIndex = -1
   cboMatCin.ListIndex = -1
   cbotypecin.ListIndex = -1
   txtamountcin.Text = ""
        
   txtissueno.Text = ""
   txtPage.Text = ""
   txtdescmag.Text = ""
   cbomatmag.ListIndex = -1
   cbotypemag.ListIndex = -1
   txtamountmag.Text = ""
   
     
     txtGrAmountCin.Text = ""
     txtNetAmountCin.Text = ""
     txtAddDiscountCin.Text = ""
     txtComPerCin.Text = ""
     txtDisccin.Text = ""
     
     txtSurcharge.Text = ""
     CboCurrency.ListIndex = -1
     lblConvRate.Visible = False
     txtConvRate.Text = ""
     txtConvRate.Visible = False
     
     
     txtfreecin.Text = ""
     txtbarterCin.Text = ""
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from Dumbo_tracinbomod"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
End Sub
Private Sub cmdClear_Click()
  textclear
End Sub
Private Sub Cmdmodify_Click()
 Dim a
 Dim B
 Dim C
 Dim X
 Dim Y
 Dim m
 Dim optcan As String
 Dim acday
 Dim actime
 Dim acdesc
 Dim acmat
 Dim acptype
 Dim acspots
 Dim acrate
 Dim acamount
 Dim acissue
 Dim acpage
 Dim acspace
 Dim acsubmedia
 Dim acsec
 Dim accode
 Dim acnewsec
 Dim acdtfrom As Date
 Dim acdtto  As Date
 Dim acimpressions
 
 Dim rcount As Currency
 Dim addiscpt As Currency
 Dim adsurcharge As Currency
 Dim agcompt
 Dim adcompt
 
  If CboCurrency.Text = "DHS" Then
      txtConvRate.Text = 1
  End If
  
optcan = ""

    Y = Cboyear.Text
    m = cboMonth.ListIndex + 1
 If cboMonth.ListIndex = 0 Then
   DT = 31
 ElseIf cboMonth.ListIndex = 1 Then
   DT = 28
 ElseIf cboMonth.ListIndex = 2 Then
   DT = 31
 ElseIf cboMonth.ListIndex = 3 Then
   DT = 30
 ElseIf cboMonth.ListIndex = 4 Then
   DT = 31
 ElseIf cboMonth.ListIndex = 5 Then
   DT = 30
 ElseIf cboMonth.ListIndex = 6 Then
   DT = 31
 ElseIf cboMonth.ListIndex = 7 Then
   DT = 31
 ElseIf cboMonth.ListIndex = 8 Then
   DT = 30
 ElseIf cboMonth.ListIndex = 9 Then
   DT = 31
 ElseIf cboMonth.ListIndex = 10 Then
   DT = 30
 Else
   DT = 31
 End If
 
Y = Cboyear.Text
m = cboMonth.ListIndex + 1
invdate = DT & " / " & m & " / " & Y
invdate = Format(invdate, "dd/mm/yyyy")
 
If lstVoucNo.SelCount = 0 Then
  MsgBox "Select Booking Order Serial No. to Modify"
  Exit Sub
End If
 X = MsgBox("Do You Want to Modify Booking Order Serial #" & Val(lstVoucNo), vbInformation + vbYesNo, "Confirm")
 If X = vbNo Then Exit Sub

 If optcancel.Value = 1 Then
       optcan = "Y"
 Else
       optcan = "N"
 End If
 
 If ValidateData = True Then
  
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    
 
  If Mid(cboMediaType, 1, 3) = "Cin" Then
                      If MSFlexcin.Rows = 1 Then
                         MsgBox "Transactions are empty, Cannot Modify"
                         Exit Sub
                      End If
                      
                       If Val(txtNetAmountCin.Text) = 0 Then
                         Sqlqry1 = "Select * from crdt_mas where Ref_no='" & lstVoucNo & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                           If rs1.RecordCount <> 0 Then
                                 MsgBox " Credit Note" & rs1!vouc_no & "is existing, You cannot modify it"
                                 Exit Sub
                           End If
                      End If
                    
                       Sqlqry = " Delete * from bo_tracinprn"
                        ws.BeginTrans
                        db.Execute Sqlqry
                        ws.CommitTrans
                                         
                      If txtGrAmountCin = "" Then
                          txtGrAmountCin = 0
                      End If
                      Sqlqry = "Update Bo_Mas set TDATE=#" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "#," & _
                                            " Tcurrency ='" & Trim(CboCurrency) & "'," & _
                                            " Tconvertion =" & txtConvRate & "," & _
                                            " Tra_gamount =" & txtGrAmountCin & "," & _
                                            " Tra_namount =" & txtNetAmountCin & "," & _
                                            " year ='" & Trim(Cboyear) & "'," & _
                                            " Month ='" & Trim(cboMonth) & "'," & _
                                            " Monthind=" & Val(cboMonth.ListIndex) & ", " & _
                                            " Region ='" & Trim(cboregion) & "'," & _
                                            " boremarks ='" & Trim(txtremarks) & "'," & _
                                            " Product ='" & findfirstfixup(Trim(CboProduct.Text)) & "'," & _
                                            " client='" & findfirstfixup(Trim(CboClient)) & "'," & _
                                            " Agency='" & findfirstfixup(Trim(CboAgency)) & "'," & _
                                            " Media ='Cinema'," & _
                                            " Sub_Media='Cinema'," & _
                                            " Bo_ref ='" & findfirstfixup(Trim(txtboref)) & "', " & _
                                            " Gross_Amount = " & Val(txtGrAmountCin) * Val(txtConvRate) & ", " & _
                                            " Tot_free=" & Val(txtfreecin) & "," & _
                                            " Tot_barter=" & Val(txtbarterCin) & "," & _
                                            " disc_Percentage='" & Val(Trim(txtDisccin)) & "'," & _
                                            " disc_rate='" & Val(Trim(txtComPerCin)) & "'," & _
                                            " add_discount=" & Val(Trim(txtAddDiscountCin)) & "," & _
                                            " Invoice_date=#" & Format(invdate, "dd/mm/yyyy") & "#," & _
                                            " cancell ='" & optcan & "' Where serial_NO = '" & Val(lstVoucNo.Text) & "'"
                              ws.BeginTrans
                              db.Execute Sqlqry
                              ws.CommitTrans
                            
                   Sqlqry1 = "Delete * from Bo_tracin where serial_no='" & Val(lstVoucNo) & "' "
                   ws.BeginTrans
                   db.Execute Sqlqry1
                   ws.CommitTrans
                   
                      agdisc = 0
                      extdisc = 0
                      adddisc = 0
                      NOS = 0
                      AddDiscEach = 0
                       
                      agdisc = Val(txtComPerCin.Text)
                      extdisc = Val(txtDisccin.Text)
                      adddisc = Val(txtAddDiscountCin.Text)
                      
    
                     
                With MSFlexcin
                      a = .Rows
                     For B = 1 To a - 1
                      .Row = B
                      .Col = 0
                        acsubmedia = .Text
                      .Col = 1
                        'acdatefrom = format(.Text, "dd/mm/yyyy")
                        acdatefrom = .Text
                      .Col = 2
                        'acdateto = format(.Text, "dd/mm/yyyy")
                        acdateto = .Text
                      .Col = 3
                        acday = .Text
                      .Col = 4
                        acsec = .Text
                      .Col = 5
                        acdesc = .Text
                      .Col = 6
                        acmat = .Text
                      .Col = 7
                        acptype = .Text
                           If acptype = "Paid" Then
                             NOS = NOS + 1
                           End If
                      .Col = 8
                           acamount = .Text
                          
                      
                           Sqlqry2 = " Insert into Bo_TRAcinprn values('" & Val(lstVoucNo.Text) & "','" & Cboyear.Text & "','" _
                                                     & cboMonth.Text & "','" _
                                                     & findfirstfixup(CboProduct.Text) & "','" _
                                                     & findfirstfixup(CboClient) & "','" _
                                                     & findfirstfixup(CboAgency) & "','Cinema','" _
                                                     & acsubmedia & "','" _
                                                     & Format(acdatefrom, "DD/MM/YYYY") & "','" _
                                                     & Format(acdateto, "DD/MM/YYYY") & "','" _
                                                     & findfirstfixup(Trim(txtboref.Text)) & "','" _
                                                     & acday & "','" _
                                                     & acsec & "','" _
                                                     & findfirstfixup(Trim(acdesc)) & "','" _
                                                     & findfirstfixup(acmat) & "','" _
                                                     & acptype & "','" _
                                                     & Trim(CboCurrency.Text) & "'," _
                                                     & Val(txtConvRate.Text) & "," _
                                                     & Val(acamount) & ", " & Val(acamount) * Val(txtConvRate.Text) & ")"
       '  MsgBox Sqlqry2
                           ws.BeginTrans
                           db.Execute (Sqlqry2)
                           ws.CommitTrans
                           
                           
                       Next
                     
                End With
                If adddisc = 0 Then
                 AddDiscEach = 0
                Else
                 AddDiscEach = Round(adddisc / NOS, 3)
                End If
                   
             
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = "Select * from Bo_TRAcinprn"
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
                           
                           
            End If
            
            
            
          rs.MoveNext
         Loop
       End If
       
       Set ws = DBEngine.Workspaces(0)
       Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       Sqlqry1 = "update Bo_TRAcinprn  set tra_amount=" & 0 & ", " & _
                " amount=" & 0 & "  where Type<>'Paid'"
       
                    
       ws.BeginTrans
       db.Execute (Sqlqry1)
       ws.CommitTrans
                   
   
   
  MsgBox " Booking Order Serial No.  & '" & Val(lstVoucNo.Text) & "' &  is modified"
  
 End If
  textclear
  lstVoucNo.SetFocus
End If
 
End Sub

Private Sub cmdView_Click()
    
  If CboSyear.Text = "" Then
    MsgBox "Invalid year", vbInformation, "Invalid Entry"
    CboSyear.SetFocus
    SendKeys " {Home} + {End} "
    Exit Sub
   End If
   
   If CboSmonth.Text = "" Then
    MsgBox "Invalid Month", vbInformation, "Invalid Entry"
    CboSmonth.SetFocus
    SendKeys " {Home} + {End} "
    Exit Sub
   End If
      
   If CboSAgency.Text = "" Then
    MsgBox "Invalid Agency", vbInformation, "Invalid Entry"
    CboSAgency.SetFocus
    SendKeys " {Home} + {End} "
    Exit Sub
   End If
   
   If CboSProduct.Text = "" Then
    MsgBox "Invalid Product", vbInformation, "Invalid Entry"
    CboSProduct.SetFocus
    SendKeys " {Home} + {End} "
    Exit Sub
   End If
     
   If CboSMedia.Text = "" Then
    MsgBox "Invalid Media", vbInformation, "Invalid Entry"
    CboSMedia.SetFocus
    SendKeys " {Home} + {End} "
    Exit Sub
   End If
   
   Fradata.Visible = False
   FraView.Visible = True
   
   n = Trim(lblviewMedia.Caption)
   m = Trim(LblviewSubmedia.Caption)
 
If CboSAgency.Text <> "All" Then o = CboSAgency.Text
If CboSProduct.Text <> "All" Then p = CboSProduct.Text
If CboSmonth.Text <> "All" Then

          If CboSMedia.Text = "Cinema" Then
             Flexitemsviewcin
             If CboSAgency.Text = "All" And CboSProduct.Text = "All" Then
                Sqlqry1 = "Select * from bo_TraCin where Media='Cinema' and year='" & Val(CboSyear) & "' and month='" & Trim(CboSmonth) & "' order by sub_media"
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                     If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                            Set ws = DBEngine.Workspaces(0)
                            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                            MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & Format(rs1!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs1!Dateto, "dd/mm/yyyy") & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                         Loop
                     End If
                ElseIf CboSAgency.Text = o And CboSProduct.Text = "All" Then
                    Sqlqry1 = "Select * from bo_TraCin where Media='Cinema' and Agency='" & o & "' and year='" & Val(CboSyear) & "' and month='" & Trim(CboSmonth) & "' order by sub_media"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & Format(rs1!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs1!Dateto, "dd/mm/yyyy") & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                ElseIf CboSAgency.Text = "All" And CboSProduct.Text = p Then
                    Sqlqry1 = "Select * from bo_TraCin where Media='Cinema' and Product='" & p & "' and year='" & Val(CboSyear) & "' and month='" & Trim(CboSmonth) & "' order by sub_media"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & Format(rs1!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs1!Dateto, "dd/mm/yyyy") & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                ElseIf CboSAgency.Text = o And CboSProduct.Text = p Then
                    Sqlqry1 = "Select * from bo_TraCin where Media='Cinema' and Product='" & p & "' and agency='" & o & "' and year='" & Val(CboSyear) & "' and month='" & Trim(CboSmonth) & "' order by sub_media"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                         Set ws = DBEngine.Workspaces(0)
                         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & Format(rs1!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs1!Dateto, "dd/mm/yyyy") & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                         rs1.MoveNext
                        Loop
                     End If
                End If
          
                  
          ElseIf Mid(CboSMedia.Text, 1, 3) = "Cin" Then
             Flexitemsviewcin
             If CboSAgency.Text = "All" And CboSProduct.Text = "All" Then
                Sqlqry1 = "Select * from bo_TraCin where Media='" & n & "' and sub_media='" & m & "' and year='" & Val(CboSyear) & "' and month='" & Trim(CboSmonth) & "' order by sub_media"
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                     If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                        MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!DATEFROM & Chr(9) & rs1!Dateto & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                ElseIf CboSAgency.Text = o And CboSProduct.Text = "All" Then
                    Sqlqry1 = "Select * from bo_TraCin where Media='" & n & "'  and sub_media='" & m & "' and Agency='" & o & "' and year='" & Val(CboSyear) & "' and month='" & Trim(CboSmonth) & "' order by sub_media"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                           MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!DATEFROM & Chr(9) & rs1!Dateto & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                ElseIf CboSAgency.Text = "All" And CboSProduct.Text = p Then
                    Sqlqry1 = "Select * from bo_TraCin where Media='" & n & "' and sub_media='" & m & "' and Product='" & p & "' and year='" & Val(CboSyear) & "' and month='" & Trim(CboSmonth) & "' order by sub_media"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                           MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!DATEFROM & Chr(9) & rs1!Dateto & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                ElseIf CboSAgency.Text = o And CboSProduct.Text = p Then
                    Sqlqry1 = "Select * from bo_TraCin where Media='" & n & "' and sub_media='" & m & "' and Product='" & p & "' and agency='" & o & "' and year='" & Val(CboSyear) & "' and month='" & Trim(CboSmonth) & "' order by sub_media"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                           MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!DATEFROM & Chr(9) & rs1!Dateto & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                End If
          End If
  Else
    populatecmdviewmonth
  End If
End Sub

Private Sub populatecmdviewmonth()

n = Trim(lblviewMedia.Caption)
m = Trim(LblviewSubmedia.Caption)
 

          
          If CboSMedia.Text = "Cinema" Then
             Flexitemsviewcin
             If CboSAgency.Text = "All" And CboSProduct.Text = "All" Then
                Sqlqry1 = "Select * from bo_TraCin where Media='Cinema' and year='" & Val(CboSyear) & "' order by sub_media"
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                     If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & Format(rs1!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs1!Dateto, "dd/mm/yyyy") & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                ElseIf CboSAgency.Text = o And CboSProduct.Text = "All" Then
                    Sqlqry1 = "Select * from bo_TraCin where Media='Cinema' and Agency='" & o & "' and year='" & Val(CboSyear) & "' order by sub_media"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                        If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & Format(rs1!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs1!Dateto, "dd/mm/yyyy") & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                ElseIf CboSAgency.Text = "All" And CboSProduct.Text = p Then
                    Sqlqry1 = "Select * from bo_TraCin where Media='Cinema' and Product='" & p & "' and year='" & Val(CboSyear) & "' order by sub_media "
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & Format(rs1!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs1!Dateto, "dd/mm/yyyy") & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                ElseIf CboSAgency.Text = o And CboSProduct.Text = p Then
                    Sqlqry1 = "Select * from bo_TraCin where Media='Cinema' and Product='" & p & "' and agency='" & o & "' and year='" & Val(CboSyear) & "' order by sub_media"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                         Set ws = DBEngine.Workspaces(0)
                         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & Format(rs1!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs1!Dateto, "dd/mm/yyyy") & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                         rs1.MoveNext
                        Loop
                     End If
                End If
          ElseIf Mid(CboSMedia.Text, 1, 3) = "Cin" Then
             Flexitemsviewcin
             If CboSAgency.Text = "All" And CboSProduct.Text = "All" Then
                Sqlqry1 = "Select * from bo_TraCin where Media='" & n & "' and sub_media='" & m & "' and year='" & Val(CboSyear) & "' order by sub_media"
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                     If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                        MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!DATEFROM & Chr(9) & rs1!Dateto & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                ElseIf CboSAgency.Text = o And CboSProduct.Text = "All" Then
                    Sqlqry1 = "Select * from bo_TraCin where Media='" & n & "'  and sub_media='" & m & "' and Agency='" & o & "' and year='" & Val(CboSyear) & "' order by sub_media "
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                           MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!DATEFROM & Chr(9) & rs1!Dateto & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                ElseIf CboSAgency.Text = "All" And CboSProduct.Text = p Then
                    Sqlqry1 = "Select * from bo_TraCin where Media='" & n & "' and sub_media='" & m & "' and Product='" & p & "' and year='" & Val(CboSyear) & "' order by sub_media "
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                           MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!DATEFROM & Chr(9) & rs1!Dateto & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                ElseIf CboSAgency.Text = o And CboSProduct.Text = p Then
                    Sqlqry1 = "Select * from bo_TraCin where Media='" & n & "' and sub_media='" & m & "' and Product='" & p & "' and agency='" & o & "' and year='" & Val(CboSyear) & "' order by sub_media"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If rs1.RecordCount <> 0 Then
                        rs1.MoveFirst
                        Do Until rs1.EOF
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                           MSFlexview.AddItem rs1!serial_no & Chr(9) & rs1!Agency & Chr(9) & rs1!Product & Chr(9) & rs1!sub_Media & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!DATEFROM & Chr(9) & rs1!Dateto & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                           rs1.MoveNext
                           Loop
                     End If
                End If
          
                End If
         
End Sub
Private Sub Cmdviewclear_Click()
    FraView.Visible = False
    Fradata.Visible = True
End Sub

Private Sub Form_Load()
Dim fmname
Dim fmid

fmname = ""
fmname = Me.Caption
fmid = Me.Name

On Error GoTo xyz:

 checkin

X = 0
U = 0
Y = 0
Z = 0
MTYPE = 0


lblConvRate.Visible = False
txtConvRate.Visible = False

CboCurrency.AddItem "DHS"
CboCurrency.AddItem "USD"

cboMonth.AddItem "January"
cboMonth.AddItem "February"
cboMonth.AddItem "March"
cboMonth.AddItem "April"
cboMonth.AddItem "May"
cboMonth.AddItem "June"
cboMonth.AddItem "July"
cboMonth.AddItem "August"
cboMonth.AddItem "September"
cboMonth.AddItem "October"
cboMonth.AddItem "November"
cboMonth.AddItem "December"

CboSmonth.AddItem "All"
CboSmonth.AddItem "January"
CboSmonth.AddItem "February"
CboSmonth.AddItem "March"
CboSmonth.AddItem "April"
CboSmonth.AddItem "May"
CboSmonth.AddItem "June"
CboSmonth.AddItem "July"
CboSmonth.AddItem "August"
CboSmonth.AddItem "September"
CboSmonth.AddItem "October"
CboSmonth.AddItem "November"
CboSmonth.AddItem "December"


i = 2000
DT = 28
For i = 2000 To 2100
 Cboyear.AddItem i
 CboSyear.AddItem i
Next
X = 0

 Cboyear.Text = Year(Now())
 CboSyear.Text = Year(Now())
 
 X = Month(Now())
 
 
If X = 1 Then
   cboMonth.ListIndex = 0
   CboSmonth.ListIndex = 1
   DT = 31
ElseIf X = 2 Then
   cboMonth.ListIndex = 1
   CboSmonth.ListIndex = 2
   DT = 28
ElseIf X = 3 Then
   cboMonth.ListIndex = 2
   CboSmonth.ListIndex = 3
   DT = 31
ElseIf X = 4 Then
   cboMonth.ListIndex = 3
   CboSmonth.ListIndex = 4
   DT = 30
ElseIf X = 5 Then
   cboMonth.ListIndex = 4
   CboSmonth.ListIndex = 5
   DT = 31
ElseIf X = 6 Then
   cboMonth.ListIndex = 5
   CboSmonth.ListIndex = 6
   DT = 30
ElseIf X = 7 Then
   cboMonth.ListIndex = 6
   CboSmonth.ListIndex = 7
   DT = 31
ElseIf X = 8 Then
   cboMonth.ListIndex = 7
   CboSmonth.ListIndex = 8
   DT = 31
ElseIf X = 9 Then
   cboMonth.ListIndex = 8
   CboSmonth.ListIndex = 9
   DT = 30
ElseIf X = 10 Then
   cboMonth.ListIndex = 9
   CboSmonth.ListIndex = 10
   DT = 31
ElseIf X = 11 Then
   cboMonth.ListIndex = 10
   CboSmonth.ListIndex = 11
   DT = 30
Else
   cboMonth.ListIndex = 11
   CboSmonth.ListIndex = 12
   DT = 31
End If

txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
populateproducts
populateagency
populateclient
populateMedia
PopulateVoucher
PopulateCodes
PopulateAgencycodes

  FraView.Visible = False
  Fradata.Visible = True
  
  LblIss.Visible = False
  CboIssue.Visible = False
 
   fraTV.Visible = False
   Fraol.Visible = False
   Fracin.Visible = False
   FraMag.Visible = False
   Fraemp.Visible = True
 
optcancel.Value = 0
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
 Sqlqry = "delete * from Dumbo_tracinbomod"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
 Exit Sub
 
xyz:
 MsgBox "Table has been locked exclusively"
 
    
       
End Sub
 Private Sub PopulateCodes()
    CboCode.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select code from cnnrates"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
       
        rs.MoveFirst
       Do Until rs.EOF
         CboCode.AddItem rs!code
         rs.MoveNext
       Loop
    End If
 End Sub

Private Sub populateproducts()
    CboProduct.Clear
    CboSProduct.Clear
    CboSProduct.AddItem "All"
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from products Order by Product_Name"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        rs.MoveFirst
            Do Until rs.EOF
              CboProduct.AddItem Trim(rs!product_name)
              CboSProduct.AddItem Trim(rs!product_name)
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
    
    cboMediaType.Clear
    cboMediaType.AddItem "Cinema"
    cboMediaType.AddItem "Special Operation"
    
    CboSMedia.Clear
    CboSMedia.AddItem "Cinema"

'    Set ws = DBEngine.Workspaces(0)
'    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
'    Sqlqry = "Select * from Media Order by Media_Type"
'    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
'    If rs.RecordCount = 0 Then
'         Exit Sub
'    Else
'        cboMediaType.AddItem "Cinema"
'        rs.MoveFirst
'            Do Until rs.EOF
'              If Mid(rs!Media_Type, 1, 3) = "Cin" Then
'               rs.MoveNext
'              Else
'                cboMediaType.AddItem rs!Media_Type & "  " & Trim(rs!sub_Media)
'              rs.MoveNext
'              End If
'       Loop
'    End If
    
    
    
'    Set ws = DBEngine.Workspaces(0)
'    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
'    Sqlqry = "Select * from Media where Media_type='Cinema' Order by Media_Type"
'    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
'    If rs.RecordCount = 0 Then
'         Exit Sub
'    Else
'
'        rs.MoveFirst
'            Do Until rs.EOF
'                CboSMedia.AddItem rs!Media_Type & "  " & Trim(rs!sub_Media)
'                rs.MoveNext
'            Loop
'    End If
 End Sub

Private Sub Form_Unload(Cancel As Integer)
 fmname = ""
 fmname = Me.Caption
 fmid = Me.Name
 checkout
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
            cmdModify.Enabled = False
       
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



Private Sub lstVoucNo_Click()
Dim i
Dim X
Dim Y
Dim Z
Dim U
    textclear1
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Val(lstVoucNo.Text)
        
        Sqlqry2 = " Select * from Bo_Mas Where Serial_no= '" & i & "'"
        Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
         If rs2.RecordCount <> 0 Then
          
           txtdate.TextWithMask = Format(rs2!tDate, "dd/mm/yyyy")
           Cboyear.Text = rs2!Year
           Y = Trim(rs2!Month)
            
           If Y = "December" Then
             cboMonth.ListIndex = 11
           ElseIf Y = "November" Then
             cboMonth.ListIndex = 10
           ElseIf Y = "October" Then
             cboMonth.ListIndex = 9
           ElseIf Y = "September" Then
             cboMonth.ListIndex = 8
           ElseIf Y = "August" Then
             cboMonth.ListIndex = 7
           ElseIf Y = "July" Then
             cboMonth.ListIndex = 6
           ElseIf Y = "June" Then
             cboMonth.ListIndex = 5
           ElseIf Y = "May" Then
             cboMonth.ListIndex = 4
           ElseIf Y = "April" Then
             cboMonth.ListIndex = 3
           ElseIf Y = "March" Then
             cboMonth.ListIndex = 2
           ElseIf Y = "February" Then
             cboMonth.ListIndex = 1
           Else
             cboMonth.ListIndex = 0
           End If
            
           
           
           If IsNull(rs2!bo_ref) = True Then
             txtboref.Text = ""
           Else
             txtboref.Text = rs2!bo_ref
           End If
           
           If IsNull(rs2!boremarks) = True Then
             txtremarks.Text = ""
           Else
             txtremarks.Text = rs2!boremarks
           End If
           
           If IsNull(rs2!region) = True Then
             cboregion.Text = ""
           Else
             cboregion.Text = rs2!region
           End If
           
           
           CboAgency = rs2!Agency
           CboClient = rs2!client
           CboProduct = Trim(rs2!Product)
           
                                                 
           If rs2!tcurrency = "USD" Then
             CboCurrency.ListIndex = 1
             lblConvRate.Visible = True
             txtConvRate.Visible = True
             txtConvRate.Text = rs2!tconvertion
             txtConvRate.TabIndex = 4
           Else
             CboCurrency.ListIndex = 0
             lblConvRate.Visible = False
             txtConvRate.Visible = False
             txtConvRate.Text = rs2!tconvertion
           End If
             
           If rs2!cancell = "N" Then
             optcancel.Value = 0
           Else
             optcancel.Value = 1
           End If
                      
         If rs2!Media = "Cinema" Then
               fraTV.Visible = False
               Fraol.Visible = False
               Fracin.Visible = True
               FraMag.Visible = False
               Fraemp.Visible = False
               cboMediaType.Text = "Cinema"
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 Sqlqry1 = "Select * from cinema_rates order by sub_media"
                 Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                 If rs1.RecordCount <> 0 Then
                    rs1.MoveFirst
                    cbosubmedia.Clear
                    Do Until rs1.EOF
                     cbosubmedia.AddItem rs1!sub_Media
                     rs1.MoveNext
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
               MSFlexcin.Clear
               Flexitemscin
               Sqlqry1 = "Select * from Bo_tracin where Serial_no= '" & i & " '"
               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                 
                   'MSFlexcin.Clear
                   rs1.MoveFirst
                   Do Until rs1.EOF
                   Set ws = DBEngine.Workspaces(0)
                   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                   
                     Sqlqry2 = " Insert into Dumbo_tracinbomod values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                     & Trim(rs1!Month) & "','" _
                                     & findfirstfixup(rs1!Product) & "','" _
                                     & findfirstfixup(rs1!client) & "','" _
                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                     & Trim(rs1!sub_Media) & "','" _
                                     & Trim(rs1!DATEFROM) & "','" _
                                     & Trim(rs1!Dateto) & "','" _
                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                     & Trim(rs1!Day) & "','" _
                                     & Trim(rs1!Length) & "','" _
                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                     & Trim(rs1!Type) & "','" _
                                     & Trim(rs1!tcurrency) & "'," _
                                     & Trim(rs1!tconvertion) & "," _
                                     & Trim(rs1!tra_amount) & "," _
                                     & Trim(rs1!Amount) & ")"
                            ws.BeginTrans
                            db.Execute (Sqlqry2)
                            ws.CommitTrans
                   
                  
                    MSFlexcin.AddItem rs1!sub_Media & Chr(9) & Format(rs1!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs1!Dateto, "dd/mm/yyyy") & Chr(9) & rs1!Day & Chr(9) & rs1!Length & Chr(9) & rs1!Description & Chr(9) & rs1!mat_code & Chr(9) & rs1!Type & Chr(9) & rs1!tra_amount
                    rs1.MoveNext
                   Loop
                   
                End If
                
               txtGrAmountCin.Text = rs2!tra_gamount
               txtNetAmountCin.Text = rs2!tra_namount
               txtfreecin.Text = rs2!Tot_free
               txtbarterCin.Text = rs2!Tot_barter
               txtComPerCin.Text = rs2!disc_rate
               txtDisccin.Text = rs2!disc_percentage
               txtAddDiscountCin.Text = rs2!add_discount
               
            Else
               fraTV.Visible = False
               Fraol.Visible = False
               Fracin.Visible = False
               FraMag.Visible = False
               Fraemp.Visible = True
             End If
             
          
                      
       End If
    
End Sub
Private Sub MSFlexview_DblClick()
 Dim i
 Dim j
 Dim X
 Dim Y, Z, U
 Dim ref As String
 
 X = MSFlexview.Rows
 
 If X > 1 Then
   i = MsgBox(" Are you sure .. ! You want to Modify this transaction (id #) ", vbInformation + vbYesNo)
    If i = vbYes Then
     With MSFlexview
        j = .Row
        .Col = 0
        ref = .Text
      End With
      lstVoucNo.Text = Val(ref)
     End If
 End If
 
   FraView.Visible = False
   Fradata.Visible = True
End Sub
Private Sub OptCancel_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdate.SetFocus
End Sub
Private Sub OptCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If optcancel.Value = 0 Then
  optcancel.Value = 1
 ElseIf optcancel.Value = 1 Then
  optcancel.Value = 0
 End If
 txtdate.SetFocus
End Sub
Private Sub txtAddDiscountCin_LostFocus()
  If txtDisccin.Text = "" Then txtDisccin.Text = 0
  If txtAddDiscountCin.Text = "" Then txtAddDiscountCin.Text = 0
 txtNetAmountCin.Text = Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100) - (Val(Val(txtGrAmountCin.Text) - (Val(txtGrAmountCin.Text) * Val(txtComPerCin.Text) / 100)) * txtDisccin / 100) - Val(txtAddDiscountCin.Text)
 cmdModify.SetFocus
End Sub
Private Sub txtadddiscountCIN_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdModify.SetFocus
End Sub
Private Sub txtamountcin_GotFocus()
 SendKeys "{Home} + {End}"
End Sub
Private Sub txtamountcin_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cbosubmedia.SetFocus
End Sub

Private Sub txtamountcin_LostFocus()
  
  If ValidateData = True Then
      
   If IsNumeric(txtamountcin.Text) = False Or IsNull(txtamountcin.Text) = True Then
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
    Sqlqry1 = " select * from Dumbo_tracinbomod"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
     If rs.RecordCount = 0 Then
       Sqlqry1 = " Insert into Dumbo_tracinbomod values('" & Val(lstVoucNo) & "','" & Cboyear.Text & "','" _
                                     & cboMonth.Text & "','" _
                                     & findfirstfixup(CboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Cinema','" _
                                     & Trim(cbosubmedia.Text) & "','" _
                                     & Format(txtCinDateFrom.TextWithMask, "dd/mm/YYYY") & "','" _
                                     & Format(txtCinDateTo.TextWithMask, "dd/mm/YYYY") & "','" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "','" _
                                     & Trim(cbodaycin.Text) & "','" _
                                     & Trim(cbolength.Text) & "','" _
                                     & findfirstfixup(Trim(txtDescCin.Text)) & "','" _
                                     & findfirstfixup(Trim(cboMatCin.Text)) & "','" _
                                     & Trim(cbotypecin.Text) & "','" _
                                     & Trim(CboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(Trim(txtamountcin.Text)) & ", " & Val(Trim(txtamountcin.Text)) * Val(txtConvRate.Text) & ")"

        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
        
        Sqlqry1 = "select * from Dumbo_tracinbomod"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlexcin.Clear
            Exit Sub
        Else
            Flexitemscin
            rs.MoveFirst
            Do Until rs.EOF
              'MSFlexcin.AddItem rs!sub_media & Chr(9) & Trim(rs!DATEFROM) & Chr(9) & Trim(rs!DATETO) & Chr(9) & rs!Day & Chr(9) & rs!Length & Chr(9) & rs!Description & Chr(9) & rs!mat_code & Chr(9) & rs!Type & Chr(9) & rs!tra_amount
              MSFlexcin.AddItem rs!sub_Media & Chr(9) & Format(rs!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs!Dateto, "dd/mm/yyyy") & Chr(9) & rs!Day & Chr(9) & rs!Length & Chr(9) & rs!Description & Chr(9) & rs!mat_code & Chr(9) & rs!Type & Chr(9) & rs!tra_amount
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
         Sqlqry1 = " Insert into Dumbo_tracinbomod values('" & Val(lstVoucNo) & "','" & Cboyear.Text & "','" _
                                     & cboMonth.Text & "','" _
                                     & findfirstfixup(CboProduct.Text) & "','" _
                                     & findfirstfixup(CboClient) & "','" _
                                     & findfirstfixup(CboAgency) & "','Cinema','" _
                                     & Trim(cbosubmedia) & "','" _
                                     & Format(txtCinDateFrom.TextWithMask, "dd/mm/YYYY") & "','" _
                                     & Format(txtCinDateTo.TextWithMask, "dd/mm/YYYY") & "','" _
                                     & findfirstfixup(Trim(txtboref.Text)) & "','" _
                                     & Trim(cbodaycin.Text) & "','" _
                                     & Trim(cbolength.Text) & "','" _
                                     & findfirstfixup(Trim(txtDescCin.Text)) & "','" _
                                     & findfirstfixup(Trim(cboMatCin.Text)) & "','" _
                                     & Trim(cbotypecin.Text) & "','" _
                                     & Trim(CboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(Trim(txtamountcin.Text)) & ", " & Val(Trim(txtamountcin.Text)) * Val(txtConvRate.Text) & ")"


        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
       
        
        Sqlqry1 = "select * from Dumbo_tracinbomod"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlexcin.Clear
            Exit Sub
        Else
            Flexitemscin
            rs.MoveFirst
            Do Until rs.EOF
             MSFlexcin.AddItem rs!sub_Media & Chr(9) & Format(rs!DATEFROM, "dd/mm/yyyy") & Chr(9) & Format(rs!Dateto, "dd/mm/yyyy") & Chr(9) & rs!Day & Chr(9) & rs!Length & Chr(9) & rs!Description & Chr(9) & rs!mat_code & Chr(9) & rs!Type & Chr(9) & rs!tra_amount
             ' MSFlexcin.AddItem rs!sub_media & Chr(9) & Trim(rs!DATEFROM) & Chr(9) & Trim(rs!DATETO) & Chr(9) & rs!Day & Chr(9) & rs!Length & Chr(9) & rs!Description & Chr(9) & rs!mat_code & Chr(9) & rs!Type & Chr(9) & rs!tra_amount
              rs.MoveNext
            Loop
        End If
        
        Sqlqry1 = "select sum(tra_amount) from Dumbo_tracinbomod where type='Paid'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then U = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from Dumbo_tracinbomod where type='Free'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Y = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from Dumbo_tracinbomod where type='Barter'"
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
Private Sub txtboref_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboregion.SetFocus
End Sub
Private Sub cbodaycin_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbolength.SetFocus
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


Private Sub txtCinDateTo_GotFocus()
    If Mid(txtCinDateFrom.TextWithMask, 4, 2) > 12 Then
          MsgBox "Invalid cinema Date from", vbInformation, "Invalid Entry"
          txtCinDateFrom.SetFocus
          SendKeys " {Home} + {End} "
    End If
End Sub

Private Sub txtCinDateTo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbodaycin.SetFocus
End Sub


Private Sub txtCinDateTo_LostFocus()
If IsDate(Format(txtCinDateTo.TextWithMask, "dd/mm/yyyy")) = False Then
       MsgBox "Invalid Cinema Date To", vbInformation, "Invalid Entry"
       txtCinDateTo.SetFocus
       SendKeys " {Home} + {End} "
End If
    
If DateValue(Format(txtCinDateFrom.TextWithMask, "dd/mm/yyyy")) > DateValue(Format(txtCinDateTo.TextWithMask, "dd/mm/yyyy")) Then
    MsgBox "Date to cannot be lesser than date from", vbInformation, "Invalid entry"
      ' txtCinDateTo.SetFocus
      Exit Sub
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

Private Sub txtConvRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboProduct.SetFocus
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboCurrency.SetFocus
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
   MsgBox "Invalid Date from", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdescol.SetFocus
End Sub

Private Sub txtdateto_LostFocus()
If IsDate(txtdateto.TextWithMask) = False Then
   MsgBox "Invalid Date to", vbInformation, "Invalid Entry"
   txtdateto.SetFocus
   SendKeys " {Home} + {End} "
End If

If DateValue(txtdateto.TextWithMask) > DateValue(txtdatefrom.TextWithMask) Then
 MsgBox " Date To Cannot be greater than Date From"
 txtdateto.SetFocus
 Exit Sub
End If

End Sub
Private Sub txtDesccin_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboMatCin.SetFocus
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

Private Sub txtSEc_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtSpots.SetFocus
End Sub

Private Sub txtSEc_LostFocus()
 thirtysecrate = 0
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select rate from cnnrates where code='" & Trim(CboCode) & "'Order by code  "
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount <> 0 Then
       thirtysecrate = rs!Rate
       txtRatetv = Round((thirtysecrate / 30) * Val(txtSEc), 2)
    End If
End Sub

Private Function ValidateData()
ValidateData = False
If Cboyear.Text = "" Then
  MsgBox "Invalid Year ", vbInformation, "Invalid Entry"
  Cboyear.SetFocus
  SendKeys "{Home} + {End}"
  Exit Function
ElseIf cboMonth.Text = "" Then
  MsgBox "Invalid Month", vbInformation, "Invalid Entry"
  cboMonth.SetFocus
  Exit Function
ElseIf CboProduct.Text = "" Then
  MsgBox "Invalid Product", vbInformation, "Invalid Entry"
  CboProduct.SetFocus
  Exit Function
ElseIf cboMediaType.Text = "" Then
  MsgBox "Invalid Media Type", vbInformation, "Invalid Entry"
  cboMediaType.SetFocus
  Exit Function
ElseIf CboCurrency.Text = "" Then
  MsgBox "Select Currency Type", vbInformation, "Invalid Entry"
  CboCurrency.SetFocus
  Exit Function
ElseIf txtConvRate.Text = "" Then
  MsgBox "Enter Convertion Rate - - cannot be zero", vbInformation, "Invalid Entry"
  txtConvRate.SetFocus
  Exit Function
  
Else
  ValidateData = True
End If
End Function

Private Sub PopulateAgencycodes()
    CboSAgency.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from agndtls Order by AgentName"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
      rs.MoveFirst
        CboSAgency.Clear
         CboSAgency.AddItem "All"
        Do Until rs.EOF
            CboSAgency.AddItem rs!agentname
            rs.MoveNext
        Loop
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

Private Sub Flexitemsviewcin()
With MSFlexview

    .Clear
    .AllowUserResizing = flexResizeColumns
    .Rows = 1
    .Cols = 11
    .Col = 0
    .Row = 0
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Id #"
    .ColAlignment(0) = 0
    .ColWidth(0) = 750
    .ColWidth(1) = 1275
    .ColWidth(2) = 1200
    .ColWidth(3) = 1300
    .ColWidth(4) = 1075
    .ColWidth(5) = 1075
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(7) = 1100
    .ColWidth(8) = 800
    .ColWidth(9) = 800
    
    .Col = 1
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Agency"
    .Col = 2
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Product"
    
    .Col = 3
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Sub Media"
    
    .Col = 4
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "DateFrom"
    
    .Col = 5
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Date To"
    
    .Col = 6
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Days"
    .Col = 7
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Secs."
    
        
    .Col = 8
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "Material"
    .Col = 9
    .CellTextStyle = flexTextInset
    .CellBackColor = RGB(180, 180, 180)
    .CellFontSize = 8
    .CellFontBold = True
    .Text = "P_Type"
    .Col = 10
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
        txtCinDateFrom.TextWithMask = ""
        txtCinDateTo.TextWithMask = ""
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
        txtDescCin = .Text
        .Col = 6
        cboMatCin = .Text
        .Col = 7
        cbotypecin = .Text
        .Col = 8
        txtamountcin = .Text
                            
        .RemoveItem (j)
                             
      ' Sqlqry1 = "select * from Dumbo_tracinbomod where  description ='" & Trim(txtDesccin) & "' and Day ='" & Trim(cbodaycin) & "' and Length ='" & Trim(cbolength) & "' and tra_amount =" & Val(txtamountcin) & " AND TYPE='" & Trim(cbotypecin.Text) & "' AND DATEFROM=#" & Format(txtCinDateFrom.TextWithMask, "MM/DD/YYYY") & "# AND DATETO=#" & Format(txtCinDateTo.TextWithMask, "MM/DD/YYYY") & "# "
        Sqlqry1 = "select * from Dumbo_tracinbomod where  sub_media ='" & cbosubmedia & "' and description ='" & Trim(txtDescCin) & "' and Day ='" & Trim(cbodaycin) & "' and Length ='" & Trim(cbolength) & "' and tra_amount =" & Val(txtamountcin) & " AND TYPE='" & Trim(cbotypecin.Text) & "' AND DATEFROM=#" & DateValue(Format(txtCinDateFrom.TextWithMask, "dd/mm/YYYY")) & "# AND DATETO=#" & DateValue(Format(txtCinDateTo.TextWithMask, "DD/MM/YYYY")) & "# "
     '  MsgBox Sqlqry1
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs1.RecordCount <> 0 Then
        rs1.MoveLast
        rs1.Delete
        End If
        
         Sqlqry1 = "select sum(tra_amount) from Dumbo_tracinbomod where type='Paid'"
         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
         If IsNull(rs1.Fields(0)) = False Then U = rs1.Fields(0)
         
         Sqlqry1 = "select sum(tra_amount) from Dumbo_tracinbomod where type='Free'"
         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Y = rs1.Fields(0)
            
        Sqlqry1 = "select sum(tra_amount) from Dumbo_tracinbomod where type='Barter'"
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
