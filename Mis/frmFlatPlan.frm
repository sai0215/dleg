VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmFlatPlan 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Flat Plan"
   ClientHeight    =   8490
   ClientLeft      =   -15
   ClientTop       =   315
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2160
      TabIndex        =   29
      Top             =   0
      Width           =   6135
      Begin VB.OptionButton OptDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date Wise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   720
         TabIndex        =   31
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton OptMonth 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Month Wise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   3360
         TabIndex        =   30
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame FraMain 
      BackColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   10935
      Begin VB.CommandButton cmdExcel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Excel"
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
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   6720
         Width           =   855
      End
      Begin VB.Frame Fradate 
         BackColor       =   &H00FFFFFF&
         Height          =   6135
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   10695
         Begin VB.Frame Frasort 
            BackColor       =   &H8000000E&
            Caption         =   "Sort"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1095
            Left            =   9360
            TabIndex        =   42
            Top             =   4800
            Width           =   1215
            Begin VB.OptionButton OptIssue 
               BackColor       =   &H8000000E&
               Caption         =   "Issue #"
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
               TabIndex        =   44
               Top             =   720
               Width           =   975
            End
            Begin VB.OptionButton OptAgency 
               BackColor       =   &H8000000E&
               Caption         =   "Agency"
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
               TabIndex        =   43
               Top             =   360
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.CommandButton cmddtl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Selected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4200
            Picture         =   "frmFlatPlan.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   3960
            Width           =   1095
         End
         Begin VB.CommandButton CmdDtll 
            BackColor       =   &H00C0C0C0&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4200
            Picture         =   "frmFlatPlan.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   4800
            Width           =   1095
         End
         Begin VB.CommandButton cmddtg 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Selected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4200
            Picture         =   "frmFlatPlan.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   2280
            Width           =   1095
         End
         Begin VB.ListBox lstdtissuesel 
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
            Height          =   3420
            Left            =   6120
            MultiSelect     =   1  'Simple
            TabIndex        =   4
            Top             =   2280
            Width           =   3015
         End
         Begin VB.ListBox lstdtissue 
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
            Height          =   3420
            Left            =   480
            MultiSelect     =   1  'Simple
            TabIndex        =   3
            Top             =   2280
            Width           =   2895
         End
         Begin VB.CommandButton cmddtgg 
            BackColor       =   &H00C0C0C0&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4200
            Picture         =   "frmFlatPlan.frx":0CC6
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   3120
            Width           =   1095
         End
         Begin VB.ComboBox cbodtsubmedia 
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
            ForeColor       =   &H000040C0&
            Height          =   360
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1320
            Width           =   5415
         End
         Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
            Height          =   255
            Left            =   3120
            TabIndex        =   0
            Top             =   720
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
         Begin PVMaskEditLib.PVMaskEdit txtdateto 
            Height          =   255
            Left            =   7080
            TabIndex        =   1
            Top             =   720
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
         Begin VB.Line Line2 
            BorderColor     =   &H008080FF&
            BorderWidth     =   2
            X1              =   0
            X2              =   10680
            Y1              =   6120
            Y2              =   6120
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Issue Date From"
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
            Height          =   255
            Left            =   960
            TabIndex        =   28
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Issue Date To"
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
            Height          =   255
            Left            =   5160
            TabIndex        =   27
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   " Sub Media"
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
            Height          =   255
            Left            =   1080
            TabIndex        =   26
            Top             =   1440
            Width           =   1935
         End
      End
      Begin VB.Frame FraMonth 
         BackColor       =   &H00FFFFFF&
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
         Height          =   6135
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   10695
         Begin VB.CommandButton cmdL 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Selected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4560
            Picture         =   "frmFlatPlan.frx":1108
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   4200
            Width           =   1095
         End
         Begin VB.CommandButton cmdLL 
            BackColor       =   &H00C0C0C0&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4560
            Picture         =   "frmFlatPlan.frx":154A
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   5040
            Width           =   1095
         End
         Begin VB.CommandButton Cmdg 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Selected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4560
            Picture         =   "frmFlatPlan.frx":198C
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2280
            Width           =   1095
         End
         Begin VB.ListBox lstissuesel 
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
            Height          =   3420
            Left            =   6480
            MultiSelect     =   1  'Simple
            TabIndex        =   14
            Top             =   2280
            Width           =   3015
         End
         Begin VB.ListBox lstissue 
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
            Height          =   3420
            Left            =   840
            MultiSelect     =   1  'Simple
            TabIndex        =   13
            Top             =   2280
            Width           =   2895
         End
         Begin VB.CommandButton cmdGG 
            BackColor       =   &H00C0C0C0&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4560
            Picture         =   "frmFlatPlan.frx":1DCE
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3240
            Width           =   1095
         End
         Begin VB.ComboBox CbomonthTo 
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
            ForeColor       =   &H000040C0&
            Height          =   420
            Left            =   6600
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   960
            Width           =   2895
         End
         Begin VB.ComboBox Cbomonthfrom 
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
            ForeColor       =   &H000040C0&
            Height          =   420
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   960
            Width           =   2895
         End
         Begin VB.ComboBox Cboyear 
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
            ForeColor       =   &H000040C0&
            Height          =   420
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   2895
         End
         Begin VB.ComboBox Cbosubmedia 
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
            ForeColor       =   &H000040C0&
            Height          =   420
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1560
            Width           =   2895
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   120
            Top             =   4440
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   262150
         End
         Begin VB.Line Line1 
            BorderColor     =   &H008080FF&
            BorderWidth     =   2
            X1              =   0
            X2              =   10680
            Y1              =   6120
            Y2              =   6120
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Month To"
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
            Height          =   255
            Left            =   5280
            TabIndex        =   21
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Month From"
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
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "  Year"
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
            Height          =   255
            Left            =   3000
            TabIndex        =   19
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   " Sub Media"
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
            Height          =   255
            Left            =   2400
            TabIndex        =   18
            Top             =   1680
            Width           =   1575
         End
      End
      Begin VB.Frame Fradatesel 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   2280
         TabIndex        =   32
         Top             =   6240
         Width           =   6735
         Begin VB.CommandButton cmddtprintwoamt 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Preview without Amount"
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
            Left            =   240
            Picture         =   "frmFlatPlan.frx":2210
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmddtprint 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Preview with Amount"
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
            Left            =   1920
            Picture         =   "frmFlatPlan.frx":2652
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton CmdDtback 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Back"
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
            Left            =   4800
            Picture         =   "frmFlatPlan.frx":2A94
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmddtClear 
            BackColor       =   &H00C0E0FF&
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
            Left            =   3480
            Picture         =   "frmFlatPlan.frx":2ED6
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame framonthsel 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   2280
         TabIndex        =   36
         Top             =   6240
         Width           =   6735
         Begin VB.CommandButton cmdwithoutmoney 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Without Money"
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
            Left            =   1800
            Picture         =   "frmFlatPlan.frx":3318
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdwithmoney 
            BackColor       =   &H00C0E0FF&
            Caption         =   "With &Money"
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
            Left            =   240
            Picture         =   "frmFlatPlan.frx":375A
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdBack 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Back"
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
            Left            =   4920
            Picture         =   "frmFlatPlan.frx":3B9C
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H00C0E0FF&
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
            Left            =   3360
            Picture         =   "frmFlatPlan.frx":3FDE
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   240
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "frmFlatPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim i
Dim f
Dim Z
Dim X As Integer
Dim sum As Integer
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim agcomnull
Dim adpernull
Dim addiscnull
Dim surchargenull
Dim crdtamt As Currency
Dim crdtamteach As Currency
Private Sub cbodtsubmedia_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then lstdtissue.SetFocus
End Sub
Private Sub cbodtsubmedia_LostFocus()
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   If cbodtsubmedia = "All" Then
   lstdtissue.Clear
   lstdtissuesel.Clear
        
                    Sqlqry1 = "Select distinct(issue_no),sub_media from bo_tramag where tdate>=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and  tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and issue_no<>'' "
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                    If rs1.RecordCount <> 0 Then
                    rs1.MoveFirst
                    Do Until rs1.EOF
                       lstdtissue.AddItem rs1!issue_no & "    :   " & rs1!sub_Media
                       rs1.MoveNext
                    Loop
                    End If
       
  Else
  lstdtissue.Clear
  lstdtissuesel.Clear
   
                    Sqlqry1 = "Select distinct(issue_no),sub_media from bo_tramag where tdate>=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and  tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and sub_media='" & Trim(cbodtsubmedia.Text) & "' and issue_no<>''"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                    If rs1.RecordCount <> 0 Then
                      rs1.MoveFirst
                    Do Until rs1.EOF
                       lstdtissue.AddItem rs1!issue_no & "    :   " & rs1!sub_Media
                       rs1.MoveNext
                    Loop
                    End If
  
End If
End Sub
       

Private Sub cbomonthfrom_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CbomonthTo.SetFocus
End Sub
Private Sub cbomonthTo_GotFocus()
 CbomonthTo.Clear
 If Cbomonthfrom.ListIndex = 0 Then
    CbomonthTo.AddItem "January"
    CbomonthTo.AddItem "February"
    CbomonthTo.AddItem "March"
    CbomonthTo.AddItem "April"
    CbomonthTo.AddItem "May"
    CbomonthTo.AddItem "June"
    CbomonthTo.AddItem "July"
    CbomonthTo.AddItem "August"
    CbomonthTo.AddItem "September"
    CbomonthTo.AddItem "October"
    CbomonthTo.AddItem "November"
    CbomonthTo.AddItem "December"
    CbomonthTo.SetFocus
 ElseIf Cbomonthfrom.ListIndex = 1 Then
    CbomonthTo.AddItem "February"
    CbomonthTo.AddItem "March"
    CbomonthTo.AddItem "April"
    CbomonthTo.AddItem "May"
    CbomonthTo.AddItem "June"
    CbomonthTo.AddItem "July"
    CbomonthTo.AddItem "August"
    CbomonthTo.AddItem "September"
    CbomonthTo.AddItem "October"
    CbomonthTo.AddItem "November"
    CbomonthTo.AddItem "December"
    CbomonthTo.SetFocus
ElseIf Cbomonthfrom.ListIndex = 2 Then
    CbomonthTo.AddItem "March"
    CbomonthTo.AddItem "April"
    CbomonthTo.AddItem "May"
    CbomonthTo.AddItem "June"
    CbomonthTo.AddItem "July"
    CbomonthTo.AddItem "August"
    CbomonthTo.AddItem "September"
    CbomonthTo.AddItem "October"
    CbomonthTo.AddItem "November"
    CbomonthTo.AddItem "December"
    CbomonthTo.SetFocus
ElseIf Cbomonthfrom.ListIndex = 3 Then
    CbomonthTo.AddItem "April"
    CbomonthTo.AddItem "May"
    CbomonthTo.AddItem "June"
    CbomonthTo.AddItem "July"
    CbomonthTo.AddItem "August"
    CbomonthTo.AddItem "September"
    CbomonthTo.AddItem "October"
    CbomonthTo.AddItem "November"
    CbomonthTo.AddItem "December"
    CbomonthTo.SetFocus
ElseIf Cbomonthfrom.ListIndex = 4 Then
    CbomonthTo.AddItem "May"
    CbomonthTo.AddItem "June"
    CbomonthTo.AddItem "July"
    CbomonthTo.AddItem "August"
    CbomonthTo.AddItem "September"
    CbomonthTo.AddItem "October"
    CbomonthTo.AddItem "November"
    CbomonthTo.AddItem "December"
    CbomonthTo.SetFocus
ElseIf Cbomonthfrom.ListIndex = 5 Then
    CbomonthTo.AddItem "June"
    CbomonthTo.AddItem "July"
    CbomonthTo.AddItem "August"
    CbomonthTo.AddItem "September"
    CbomonthTo.AddItem "October"
    CbomonthTo.AddItem "November"
    CbomonthTo.AddItem "December"
    CbomonthTo.SetFocus
ElseIf Cbomonthfrom.ListIndex = 6 Then
    CbomonthTo.AddItem "July"
    CbomonthTo.AddItem "August"
    CbomonthTo.AddItem "September"
    CbomonthTo.AddItem "October"
    CbomonthTo.AddItem "November"
    CbomonthTo.AddItem "December"
    CbomonthTo.SetFocus
ElseIf Cbomonthfrom.ListIndex = 7 Then
    CbomonthTo.AddItem "August"
    CbomonthTo.AddItem "September"
    CbomonthTo.AddItem "October"
    CbomonthTo.AddItem "November"
    CbomonthTo.AddItem "December"
    CbomonthTo.SetFocus
ElseIf Cbomonthfrom.ListIndex = 8 Then
    CbomonthTo.AddItem "September"
    CbomonthTo.AddItem "October"
    CbomonthTo.AddItem "November"
    CbomonthTo.AddItem "December"
    CbomonthTo.SetFocus
ElseIf Cbomonthfrom.ListIndex = 9 Then
    CbomonthTo.AddItem "October"
    CbomonthTo.AddItem "November"
    CbomonthTo.AddItem "December"
    CbomonthTo.SetFocus
ElseIf Cbomonthfrom.ListIndex = 10 Then
    CbomonthTo.AddItem "November"
    CbomonthTo.AddItem "December"
    CbomonthTo.SetFocus
ElseIf Cbomonthfrom.ListIndex = 11 Then
    CbomonthTo.AddItem "December"
    CbomonthTo.SetFocus
Else
    CbomonthTo.SetFocus
End If
End Sub
Private Sub cbomonthTo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Cbosubmedia.SetFocus
End Sub
Private Sub cbosubmedia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then lstissue.SetFocus
End Sub
Private Sub cbosubmedia_LostFocus()
' Month
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from flatplanMonthTra"
 ws.BeginTrans
 db.Execute (Sqlqry)
 ws.CommitTrans
 
 
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        If Cbosubmedia = "All" Then
         Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(Cbomonthfrom.ListIndex) & " AND monthind<= " & Val(CbomonthTo.ListIndex) + Val(Cbomonthfrom.ListIndex) & " and media='Magazine' and cancell='N'"
        Else
         Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(Cbomonthfrom.ListIndex) & " AND monthind<= " & Val(CbomonthTo.ListIndex) + Val(Cbomonthfrom.ListIndex) & " and media='Magazine' and sub_media='" & Trim(Cbosubmedia) & "' and cancell='N'"
        End If
        
      Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
      If rs.RecordCount <> 0 Then
      lstissue.Clear
      lstissuesel.Clear
      rs.MoveFirst
      Do Until rs.EOF
                    Sqlqry1 = "Select * from bo_tramag where serial_no='" & Trim(rs!serial_no) & "' order by val(page),serial_no"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                    If rs1.RecordCount <> 0 Then
                      rs1.MoveFirst
                    '  Do Until rs1.EOF
                      
                     X = 0
                     X = rs1.RecordCount
                         crdtamt = 0
                         crdtamteach = 0
                         
                         Sqlqry2 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                         If IsNull(rs2.Fields(0)) = False Then crdtamt = rs2.Fields(0)
                          crdtamteach = Round(crdtamt / X, 2)
                          
                    
                    Do Until rs1.EOF
                         agcomnull = 0
                         adpernull = 0
                         surchargenull = 0
                         addiscnull = 0
                         
                         If IsNull(rs1!agcom) = True Then
                             agcomnull = 0
                         Else
                            agcomnull = rs1!agcom
                         End If
                         
                         If IsNull(rs1!adper) = True Then
                             adpernull = 0
                         Else
                            adpernull = rs1!adper
                         End If
                         
                         If IsNull(rs1!addisc) = True Then
                             addiscnull = 0
                         Else
                            addiscnull = rs1!addisc
                         End If
                         
                         If IsNull(rs1!surcharge) = True Then
                             surchargenull = 0
                         Else
                            surchargenull = rs1!surcharge
                         End If
                         
                         
                         Sqlqry2 = " Insert into flatplanMonthTra values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                     & Trim(rs1!Month) & "'," & rs1!monthind & ",'" _
                                     & findfirstfixup(rs1!Product) & "','" _
                                     & findfirstfixup(rs1!client) & "','" _
                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                     & Trim(rs1!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                     & Trim(rs1!issue_no) & "','" _
                                     & Trim(rs1!tDate) & "','" _
                                     & Trim(rs1!Page) & "','" _
                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                     & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                     & Trim(rs1!Space) & "','" _
                                     & Trim(rs1!Type) & "','" _
                                     & Trim(rs1!tcurrency) & "'," _
                                     & Trim(rs1!tconvertion) & "," _
                                     & Trim(rs1!tra_amount) & "," _
                                     & Trim(rs1!Amount) & ",'" _
                                     & adpernull & "','" _
                                     & agcomnull & "'," _
                                     & addiscnull + crdtamteach & "," _
                                     & surchargenull & ")"
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
                   rs1.MoveNext
                    Loop
                    End If
       rs.MoveNext
       Loop
     End If
     
     
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
  Sqlqry = "Select distinct(issue_no),sub_media from flatplanmonthtra where issue_no<>''"
  Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
      If rs.RecordCount <> 0 Then
            lstissue.Clear
            lstissuesel.Clear
            rs.MoveFirst
               Do Until rs.EOF
                  lstissue.AddItem rs!issue_no & "    :   " & rs!sub_Media
                  rs.MoveNext
               Loop
      End If
       
End Sub
Private Sub cboyear_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Cbomonthfrom.SetFocus
End Sub
Private Sub cmdBack_Click()
 Unload Me
End Sub
Private Sub CmdDtBack_Click()
 Unload Me
End Sub
Private Sub CmdDtClear_Click()
    txtdatefrom.TextWithMask = Now()
    txtdateto.TextWithMask = Now()
    cbodtsubmedia.Text = ""
    lstdtissue.Clear
    lstdtissuesel.Clear
End Sub

Private Sub CmdDtprint_Click()
    Dim i
    Dim a, B, C
    Dim X, Y, Z

   If cbodtsubmedia.Text = "" Then
    MsgBox "Invalid submedia", vbInformation, "Invalid Entry"
    cbodtsubmedia.SetFocus
    SendKeys " {Home} + {End} "
    Exit Sub
   End If
   
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = "Delete * from flatplandatetra2"
   ws.BeginTrans
   db.Execute (Sqlqry)
   ws.CommitTrans
      
   
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = "Delete * from flatplandate"
   ws.BeginTrans
   db.Execute (Sqlqry)
   ws.CommitTrans
   
                            
   If cbodtsubmedia.Text = "All" Then
        
        f = lstdtissuesel.ListIndex
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
               
            For f = 0 To lstdtissuesel.ListCount - 1
            Sqlqry = "Select * from bo_tramag where tdate>=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and  tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#  and issue_no='" & Trim(Mid(lstdtissuesel.List(f), 1, 6)) & "' order by val(page),serial_no"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            If rs.RecordCount <> 0 Then
              rs.MoveFirst
               Do Until rs.EOF
                  Sqlqry1 = "Select * from bo_tramag where tdate>=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and  tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#  and issue_no='" & Trim(Mid(lstdtissuesel.List(f), 1, 6)) & "' and serial_no='" & rs!serial_no & "' order by val(page),serial_no"
                  Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                  X = rs1.RecordCount
                         crdtamt = 0
                         crdtamteach = 0
                         
                         Sqlqry2 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                         If IsNull(rs2.Fields(0)) = False Then crdtamt = rs2.Fields(0)
                          crdtamteach = Round(crdtamt / X, 2)
                          
              
                         agcomnull = 0
                         adpernull = 0
                         surchargenull = 0
                         addiscnull = 0
                         
                         If IsNull(rs!agcom) = True Then
                             agcomnull = 0
                         Else
                            agcomnull = rs!agcom
                         End If
                         
                         If IsNull(rs!adper) = True Then
                             adpernull = 0
                         Else
                            adpernull = rs!adper
                         End If
                         
                         If IsNull(rs!addisc) = True Then
                             addiscnull = 0
                         Else
                            addiscnull = rs!addisc
                         End If
                         
                         If IsNull(rs!surcharge) = True Then
                             surchargenull = 0
                         Else
                            surchargenull = rs!surcharge
                         End If
                         
                 
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 Sqlqry1 = " Insert into flatplandatetra2 values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Format(rs!tDate, "DD/MM/YYYY") & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Val(rs!tra_amount) & "," _
                                     & Trim(rs!Amount) & ",'" _
                                     & agcomnull & "','" _
                                     & adpernull & "'," _
                                     & addiscnull + crdtamteach & "," _
                                     & surchargenull & ")"
                     ws.BeginTrans
                     db.Execute (Sqlqry1)
                     ws.CommitTrans
                
                    rs.MoveNext
                   Loop
              End If
            
        Next
     Else
        f = lstdtissuesel.ListIndex
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        For f = 0 To lstdtissuesel.ListCount - 1
            
             Sqlqry = "Select * from bo_tramag where tdate>=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and  tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#  and issue_no='" & Trim(Mid(lstdtissuesel.List(f), 1, 6)) & "' and sub_media ='" & Trim(cbodtsubmedia) & "' order by val(page),serial_no"
             Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
             If rs.RecordCount <> 0 Then
              rs.MoveFirst
              Do Until rs.EOF
                     Sqlqry1 = "Select * from bo_tramag where tdate>=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and  tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#  and issue_no='" & Trim(Mid(lstdtissuesel.List(f), 1, 6)) & "' and sub_media='" & Trim(cbodtsubmedia) & "' and serial_no='" & rs!serial_no & "' order by val(page),serial_no"
                     Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                     X = rs1.RecordCount
                        X = 0
                        X = rs1.RecordCount
                         crdtamt = 0
                         crdtamteach = 0
                         
                         Sqlqry2 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                         If IsNull(rs2.Fields(0)) = False Then crdtamt = rs2.Fields(0)
                         crdtamteach = Round(crdtamt / X, 2)
                          
                
                                    
                         agcomnull = 0
                         adpernull = 0
                         surchargenull = 0
                         addiscnull = 0
                         
                         If IsNull(rs!agcom) = True Then
                             agcomnull = 0
                         Else
                            agcomnull = rs!agcom
                         End If
                         
                         If IsNull(rs!adper) = True Then
                             adpernull = 0
                         Else
                            adpernull = rs!adper
                         End If
                         
                         If IsNull(rs!addisc) = True Then
                             addiscnull = 0
                         Else
                            addiscnull = rs!addisc
                         End If
                         
                         If IsNull(rs!surcharge) = True Then
                             surchargenull = 0
                         Else
                            surchargenull = rs!surcharge
                         End If
                         
              
                 Sqlqry1 = " Insert into flatplandatetra2 values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!Amount) & ",'" _
                                     & agcomnull & "','" _
                                     & adpernull & "'," _
                                     & addiscnull + crdtamteach & "," _
                                     & surchargenull & ")"
                     ws.BeginTrans
                     db.Execute (Sqlqry1)
                     ws.CommitTrans
                    rs.MoveNext
                   Loop
              End If
            Next
        End If
        
        Sqlqry = "Select * from bO_MAS WHERE CANCELL='Y'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
          rs.MoveFirst
             Do Until rs.EOF
                  Sqlqry1 = "DELETE * FROM flatplandatetra2 WHERE SERIAL_NO='" & rs!serial_no & "' "
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
              
               rs.MoveNext
              Loop
        End If
        
        
        Sqlqry = "Select * from flatplandatetra2"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
          rs.MoveFirst
             Do Until rs.EOF
                 If rs!Type = "Paid" Then
                    Sqlqry1 = " Insert into flatplandate values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!NET_Amount) & "," & 0 & ",'" _
                                     & rs!agcom & "','" _
                                     & rs!adper & "'," _
                                     & rs!addisc & "," _
                                     & rs!surcharge & ")"
                     ws.BeginTrans
                     db.Execute (Sqlqry1)
                     ws.CommitTrans
                 
                 Else
                    Sqlqry1 = " Insert into flatplandate values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," & 0 & "," & 0 & "," _
                                     & Trim(rs!tra_amount) & ",'" _
                                     & rs!agcom & "','" _
                                     & rs!adper & "'," _
                                     & rs!addisc & "," _
                                     & rs!surcharge & ")"
                     ws.BeginTrans
                     db.Execute (Sqlqry1)
                     ws.CommitTrans
                End If
                 
               rs.MoveNext
              Loop
        End If
        
 If OptAgency.Value = True Then
    
    With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\fpissuedate.rpt"
     .Formulas(0) = "zzz='" & " Issue Date From: " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
     .Formulas(1) = "yyy='" & Trim(cbodtsubmedia) & "'"
     .WindowState = crptMaximized
     .Action = 1
    End With
 Else
    
    With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\fpissuedateiss.rpt"
     .Formulas(0) = "zzz='" & " Issue Date From: " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
     .Formulas(1) = "yyy='" & Trim(cbodtsubmedia) & "'"
     .WindowState = crptMaximized
     .Action = 1
    End With
 End If
    
End Sub

Private Sub cmddtprintwoamt_Click()
    Dim i
    Dim a, B, C

   If cbodtsubmedia.Text = "" Then
    MsgBox "Invalid submedia", vbInformation, "Invalid Entry"
    cbodtsubmedia.SetFocus
    SendKeys " {Home} + {End} "
    Exit Sub
   End If
   
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = "Delete * from flatplandatetra"
   ws.BeginTrans
   db.Execute (Sqlqry)
   ws.CommitTrans
      
   
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = "Delete * from flatplandate"
   ws.BeginTrans
   db.Execute (Sqlqry)
   ws.CommitTrans
   
                            
 If cbodtsubmedia.Text = "All" Then
        
        f = lstdtissuesel.ListIndex
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
               
            For f = 0 To lstdtissuesel.ListCount - 1
            Sqlqry = "Select * from bo_tramag where tdate>=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and  tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#  and issue_no='" & Trim(Mid(lstdtissuesel.List(f), 1, 6)) & "' order by val(page),serial_no"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            If rs.RecordCount <> 0 Then
              rs.MoveFirst
                       X = 0
                       X = rs.RecordCount
                         crdtamt = 0
                         crdtamteach = 0
                         
                         Sqlqry2 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                         If IsNull(rs2.Fields(0)) = False Then crdtamt = rs2.Fields(0)
                          crdtamteach = Round(crdtamt / X, 2)
                          
              Do Until rs.EOF
                         agcomnull = 0
                         adpernull = 0
                         surchargenull = 0
                         addiscnull = 0
                         
                         If IsNull(rs!agcom) = True Then
                             agcomnull = 0
                         Else
                            agcomnull = rs!agcom
                         End If
                         
                         If IsNull(rs!adper) = True Then
                             adpernull = 0
                         Else
                            adpernull = rs!adper
                         End If
                         
                         If IsNull(rs!addisc) = True Then
                             addiscnull = 0
                         Else
                            addiscnull = rs!addisc
                         End If
                         
                         If IsNull(rs!surcharge) = True Then
                             surchargenull = 0
                         Else
                            surchargenull = rs!surcharge
                         End If
                         
                 
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 Sqlqry1 = " Insert into flatplandatetra values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Format(rs!tDate, "DD/MM/YYYY") & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Val(rs!tra_amount) & "," _
                                     & Trim(rs!Amount) & ",'" _
                                     & agcomnull & "','" _
                                     & adpernull & "'," _
                                     & addiscnull + crdtamteach & "," _
                                     & surchargenull & ")"
                     ws.BeginTrans
                     db.Execute (Sqlqry1)
                     ws.CommitTrans
                
                    rs.MoveNext
                   Loop
              End If
            
        Next
     Else
        f = lstdtissuesel.ListIndex
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        For f = 0 To lstdtissuesel.ListCount - 1
            
             Sqlqry = "Select * from bo_tramag where tdate>=#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#  and  tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#  and issue_no='" & Trim(Mid(lstdtissuesel.List(f), 1, 6)) & "' and sub_media ='" & Trim(cbodtsubmedia) & "' order by val(page),serial_no"
             Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
             If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  X = 0
                        X = rs.RecordCount
                         crdtamt = 0
                         crdtamteach = 0
                         
                         Sqlqry2 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                         If IsNull(rs2.Fields(0)) = False Then crdtamt = rs2.Fields(0)
                         crdtamteach = Round(crdtamt / X, 2)
                          
              Do Until rs.EOF
                                      
                         agcomnull = 0
                         adpernull = 0
                         surchargenull = 0
                         addiscnull = 0
                         
                         If IsNull(rs!agcom) = True Then
                             agcomnull = 0
                         Else
                            agcomnull = rs!agcom
                         End If
                         
                         If IsNull(rs!adper) = True Then
                             adpernull = 0
                         Else
                            adpernull = rs!adper
                         End If
                         
                         If IsNull(rs!addisc) = True Then
                             addiscnull = 0
                         Else
                            addiscnull = rs!addisc
                         End If
                         
                         If IsNull(rs!surcharge) = True Then
                             surchargenull = 0
                         Else
                            surchargenull = rs!surcharge
                         End If
                         
              
                 Sqlqry1 = " Insert into flatplandatetra values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!Amount) & ",'" _
                                     & agcomnull & "','" _
                                     & adpernull & "'," _
                                     & addiscnull + crdtamteach & "," _
                                     & surchargenull & ")"
                     ws.BeginTrans
                     db.Execute (Sqlqry1)
                     ws.CommitTrans
                    rs.MoveNext
                   Loop
              End If
            Next
        End If
        
        Sqlqry = "Select * from bO_MAS WHERE CANCELL='Y'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
          rs.MoveFirst
             Do Until rs.EOF
                  Sqlqry1 = "DELETE * FROM flatplandatetra WHERE SERIAL_NO='" & rs!serial_no & "' "
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
              
               rs.MoveNext
              Loop
        End If
        
        
        Sqlqry = "Select * from flatplandatetra"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
          rs.MoveFirst
             Do Until rs.EOF
                 If rs!Type = "Paid" Then
                    Sqlqry1 = " Insert into flatplandate values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!NET_Amount) & "," & 0 & ",'" _
                                     & rs!agcom & "','" _
                                     & rs!adper & "'," _
                                     & rs!addisc & "," _
                                     & rs!surcharge & ")"
                     ws.BeginTrans
                     db.Execute (Sqlqry1)
                     ws.CommitTrans
                 
                 Else
                    i = 0
                    Sqlqry1 = " Insert into flatplandate values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," & i & "," & i & "," _
                                     & Trim(rs!tra_amount) & ",'" _
                                     & rs!agcom & "','" _
                                     & rs!adper & "'," _
                                     & rs!addisc & "," _
                                     & rs!surcharge & ")"
                     ws.BeginTrans
                     db.Execute (Sqlqry1)
                     ws.CommitTrans
                End If
               rs.MoveNext
              Loop
        End If
        
 With CrystalReport1
  .DataFiles(0) = App.Path & "\misov.mdb"
  .ReportFileName = App.Path & "\fpissuedatewoamt.rpt"
  .Formulas(0) = "zzz='" & " Issue Date From: " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
  .Formulas(1) = "yyy='" & Trim(cbodtsubmedia) & "'"
  .WindowState = crptMaximized
  .Action = 1
 End With
End Sub

Private Sub Cmdexcel_Click()
 Dim objxl As Object
 Dim result As Variant
If OptDate.Value = True Then
   
    Set objxl = CreateObject("Excel.application")
    objxl.Workbooks.Open FileName:=App.Path & "\text1.xls"
    objxl.Visible = True
    objxl.Run "fplandt"
Else
 
    Set objxl = CreateObject("Excel.application")
    objxl.Workbooks.Open FileName:=App.Path & "\text2.xls"
    objxl.Visible = True
    objxl.Run "fplanmonth"

End If
End Sub
Private Sub cmdwithmoney_Click()
Dim i
Dim a, B, C

i = lstissuesel.ListIndex
If lstissuesel.ListCount = 0 Then
    MsgBox "Transactions are not found"
    lstissue.SetFocus
    Exit Sub
End If

 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from flatplan"
 ws.BeginTrans
 db.Execute (Sqlqry)
 ws.CommitTrans
            
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from flatplanMonth"
 ws.BeginTrans
 db.Execute (Sqlqry)
 ws.CommitTrans
            
            
 If Cbosubmedia.Text = "All" Then
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        For f = lstissuesel.ListCount - 1 To 0 Step -1
           Sqlqry = "Select * from flatplanmonthtra where year=" & Val(Cboyear.Text) & "' AND monthind >=" & Val(Cbomonthfrom.ListIndex) & " AND monthind<= " & Val(CbomonthTo.ListIndex) + Val(Cbomonthfrom.ListIndex) & " and issue_no='" & Trim(Mid(lstissuesel.List(f), 1, 6)) & "' order by val(page),serial_no"
           Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
             If rs.RecordCount <> 0 Then
              rs.MoveFirst
              Do Until rs.EOF
                          
              
                  agcomnull = 0
                         adpernull = 0
                         surchargenull = 0
                         addiscnull = 0
                         
                         If IsNull(rs!agcom) = True Then
                             agcomnull = 0
                         Else
                            agcomnull = rs!agcom
                         End If
                         
                         If IsNull(rs!adper) = True Then
                             adpernull = 0
                         Else
                            adpernull = rs!adper
                         End If
                         
                         If IsNull(rs!addisc) = True Then
                             addiscnull = 0
                         Else
                            addiscnull = rs!addisc
                         End If
                         
                         If IsNull(rs!surcharge) = True Then
                             surchargenull = 0
                         Else
                            surchargenull = rs!surcharge
                         End If
                         
                 
                 Sqlqry1 = " Insert into flatplanmonth values('" & rs!serial_no & "','" & rs!Year & "','" _
                                            & UCase(Trim(rs!Month)) & "','" _
                                            & findfirstfixup(UCase(rs!Product)) & "','" _
                                            & findfirstfixup(UCase(rs!client)) & "','" _
                                            & findfirstfixup(UCase(rs!Agency)) & "','" & UCase(Trim(rs!Media)) & "','" _
                                            & UCase(Trim(rs!sub_Media)) & "','" _
                                            & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                            & Trim(rs!issue_no) & "','" _
                                            & Trim(rs!Page) & "','" _
                                            & findfirstfixup(Trim(rs!Description)) & "','" _
                                            & findfirstfixup(Trim(rs!Comments)) & "','" _
                                            & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                            & UCase(Trim(rs!Space)) & "','" _
                                            & UCase(Trim(rs!Type)) & "','" _
                                            & UCase(Trim(rs!tcurrency)) & "'," _
                                            & Trim(rs!tconvertion) & "," _
                                            & Trim(rs!tra_amount) & "," _
                                            & Trim(rs!NET_Amount) & ",'" _
                                            & agcomnull & "','" _
                                            & adpernull & "'," _
                                            & addiscnull & "," _
                                            & surchargenull & ")"
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
              
                    rs.MoveNext
                   Loop
              End If
            
        Next
     Else
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
             
         For f = lstissuesel.ListCount - 1 To 0 Step -1
            
             Sqlqry = "Select * from flatplanmonthtra where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(Cbomonthfrom.ListIndex) & " AND monthind<= " & Val(CbomonthTo.ListIndex) + Val(Cbomonthfrom.ListIndex) & " and issue_no='" & Trim(Mid(lstissuesel.List(f), 1, 6)) & "' and sub_media ='" & Trim(Cbosubmedia) & "' order by val(page),serial_no"
             Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
             If rs.RecordCount <> 0 Then
              rs.MoveFirst
                       ' x = 0
                       ' x = rs.RecordCount
                       '  crdtamt = 0
                       '  crdtamteach = 0
                         
                       '  Sqlqry2 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                       '  Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                       '  If IsNull(rs2.Fields(0)) = False Then crdtamt = rs2.Fields(0)
                       '  crdtamteach = Round(crdtamt / x, 2)
 
 
                
              Do Until rs.EOF
              
                agcomnull = 0
                         adpernull = 0
                         surchargenull = 0
                         addiscnull = 0
                         
                         If IsNull(rs!agcom) = True Then
                             agcomnull = 0
                         Else
                            agcomnull = rs!agcom
                         End If
                         
                         If IsNull(rs!adper) = True Then
                             adpernull = 0
                         Else
                            adpernull = rs!adper
                         End If
                         
                         If IsNull(rs!addisc) = True Then
                             addiscnull = 0
                         Else
                            addiscnull = rs!addisc
                         End If
                         
                         If IsNull(rs!surcharge) = True Then
                             surchargenull = 0
                         Else
                            surchargenull = rs!surcharge
                         End If
                         
                 
                 
                 Sqlqry1 = " Insert into flatplanMonth values('" & rs!serial_no & "','" & rs!Year & "','" _
                                            & Trim(rs!Month) & "','" _
                                            & findfirstfixup(UCase(rs!Product)) & "','" _
                                            & findfirstfixup(UCase(rs!client)) & "','" _
                                            & findfirstfixup(UCase(rs!Agency)) & "','" & Trim(UCase(rs!Media)) & "','" _
                                            & UCase(Trim(rs!sub_Media)) & "','" _
                                            & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                            & Trim(rs!issue_no) & "','" _
                                            & Trim(rs!Page) & "','" _
                                            & findfirstfixup(Trim(rs!Description)) & "','" _
                                            & findfirstfixup(Trim(rs!Comments)) & "','" _
                                            & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                            & UCase(Trim(rs!Space)) & "','" _
                                            & UCase(Trim(rs!Type)) & "','" _
                                            & UCase(Trim(rs!tcurrency)) & "'," _
                                            & Trim(rs!tconvertion) & "," _
                                            & Trim(rs!tra_amount) & "," _
                                            & Trim(rs!NET_Amount) & ",'" _
                                            & agcomnull & "','" _
                                            & adpernull & "'," _
                                            & addiscnull & "," _
                                            & surchargenull & ")"
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
              
                    rs.MoveNext
                   Loop
              End If
            
            Next
        End If
        
        
        Sqlqry = "Select * from bO_MAS WHERE CANCELL='Y'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
          rs.MoveFirst
             Do Until rs.EOF
                  Sqlqry1 = "DELETE * FROM flatplanMonth WHERE SERIAL_NO='" & rs!serial_no & "' "
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
              
               rs.MoveNext
              Loop
        End If
        
        Sqlqry = "Select * from flatplanmonth"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
          rs.MoveFirst
          Do Until rs.EOF
                If rs!Type = "PAID" Then
                  Sqlqry1 = " Insert into flatplan values('" & rs!serial_no & "','" & rs!Year & "','" _
                                            & Trim(rs!Month) & "','" _
                                            & findfirstfixup(UCase(rs!Product)) & "','" _
                                            & findfirstfixup(UCase(rs!client)) & "','" _
                                            & findfirstfixup(UCase(rs!Agency)) & "','" & Trim(UCase(rs!Media)) & "','" _
                                            & UCase(Trim(rs!sub_Media)) & "','" _
                                            & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                            & Trim(rs!issue_no) & "','" _
                                            & Trim(rs!Page) & "','" _
                                            & findfirstfixup(Trim(rs!Description)) & "','" _
                                            & findfirstfixup(Trim(rs!Comments)) & "','" _
                                            & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                            & UCase(Trim(rs!Space)) & "','" _
                                            & UCase(Trim(rs!Type)) & "','" _
                                            & UCase(Trim(rs!tcurrency)) & "'," _
                                            & Trim(rs!tconvertion) & "," _
                                            & Trim(rs!tra_amount) & "," _
                                            & Trim(rs!Amount) & "," & 0 & ",'" _
                                            & rs!agcom & "','" _
                                            & rs!adper & "'," _
                                            & rs!addisc & "," _
                                            & rs!surcharge & ")"
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
              
                 
                
                Else
                  Sqlqry1 = " Insert into flatplan values('" & rs!serial_no & "','" & rs!Year & "','" _
                                            & Trim(rs!Month) & "','" _
                                            & findfirstfixup(UCase(rs!Product)) & "','" _
                                            & findfirstfixup(UCase(rs!client)) & "','" _
                                            & findfirstfixup(UCase(rs!Agency)) & "','" & Trim(UCase(rs!Media)) & "','" _
                                            & UCase(Trim(rs!sub_Media)) & "','" _
                                            & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                            & Trim(rs!issue_no) & "','" _
                                            & Trim(rs!Page) & "','" _
                                            & findfirstfixup(Trim(rs!Description)) & "','" _
                                            & findfirstfixup(Trim(rs!Comments)) & "','" _
                                            & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                            & UCase(Trim(rs!Space)) & "','" _
                                            & UCase(Trim(rs!Type)) & "','" _
                                            & UCase(Trim(rs!tcurrency)) & "'," _
                                            & Trim(rs!tconvertion) & "," & 0 & "," & 0 & "," _
                                            & Trim(rs!tra_amount) & ",'" _
                                            & rs!agcom & "','" _
                                            & rs!adper & "'," _
                                            & rs!addisc & "," _
                                            & rs!surcharge & ")"
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
              
                
                End If
                
               rs.MoveNext
              Loop
        End If
         
   
    
               
 With CrystalReport1
  .DataFiles(0) = App.Path & "\misov.mdb"
  .ReportFileName = App.Path & "\fpwithamount.rpt"
  .Formulas(0) = "zzz='" & "Month From " & Trim(Cbomonthfrom.Text) & " " & Trim(Cboyear) & " To " & Trim(CbomonthTo.Text) & " " & Trim(Cboyear) & "'"
  .Formulas(1) = "yyy='" & Trim(Cbosubmedia) & "'"
  .WindowState = crptMaximized
  .Action = 1
 End With
 
End Sub

Private Sub cmdwithoutmoney_Click()
Dim i
Dim a, B, C

i = lstissuesel.ListIndex
If lstissuesel.ListCount = 0 Then
    MsgBox "Transactions are not found"
    lstissue.SetFocus
    Exit Sub
End If

 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from flatplan"
 ws.BeginTrans
 db.Execute (Sqlqry)
 ws.CommitTrans
            
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from flatplanMonth"
 ws.BeginTrans
 db.Execute (Sqlqry)
 ws.CommitTrans
            
            
 If Cbosubmedia.Text = "All" Then
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        For f = lstissuesel.ListCount - 1 To 0 Step -1
           Sqlqry = "Select * from flatplanmonthtra where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(Cbomonthfrom.ListIndex) & " AND monthind<= " & Val(CbomonthTo.ListIndex) + Val(Cbomonthfrom.ListIndex) & " and issue_no='" & Trim(Mid(lstissuesel.List(f), 1, 6)) & "' order by val(page),serial_no"
           Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
             If rs.RecordCount <> 0 Then
              rs.MoveFirst
                      '  x = 0
                      '  x = rs.RecordCount
                      '   crdtamt = 0
                      '   crdtamteach = 0
                         
                      '   Sqlqry2 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                      '   Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                      '   If IsNull(rs2.Fields(0)) = False Then crdtamt = rs2.Fields(0)
                      '   crdtamteach = Round(crdtamt / x, 2)
                          
                
              
              Do Until rs.EOF
                  agcomnull = 0
                         adpernull = 0
                         surchargenull = 0
                         addiscnull = 0
                         
                         If IsNull(rs!agcom) = True Then
                             agcomnull = 0
                         Else
                            agcomnull = rs!agcom
                         End If
                         
                         If IsNull(rs!adper) = True Then
                             adpernull = 0
                         Else
                            adpernull = rs!adper
                         End If
                         
                         If IsNull(rs!addisc) = True Then
                             addiscnull = 0
                         Else
                            addiscnull = rs!addisc
                         End If
                         
                         If IsNull(rs!surcharge) = True Then
                             surchargenull = 0
                         Else
                            surchargenull = rs!surcharge
                         End If
                         
                 
                 Sqlqry1 = " Insert into flatplanmonth values('" & rs!serial_no & "','" & rs!Year & "','" _
                                            & UCase(Trim(rs!Month)) & "','" _
                                            & findfirstfixup(UCase(rs!Product)) & "','" _
                                            & findfirstfixup(UCase(rs!client)) & "','" _
                                            & findfirstfixup(UCase(rs!Agency)) & "','" & UCase(Trim(rs!Media)) & "','" _
                                            & UCase(Trim(rs!sub_Media)) & "','" _
                                            & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                            & Trim(rs!issue_no) & "','" _
                                            & Trim(rs!Page) & "','" _
                                            & findfirstfixup(Trim(rs!Description)) & "','" _
                                            & findfirstfixup(Trim(rs!Comments)) & "','" _
                                            & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                            & UCase(Trim(rs!Space)) & "','" _
                                            & UCase(Trim(rs!Type)) & "','" _
                                            & UCase(Trim(rs!tcurrency)) & "'," _
                                            & Trim(rs!tconvertion) & "," _
                                            & Trim(rs!tra_amount) & "," _
                                            & Trim(rs!NET_Amount) & ",'" _
                                            & agcomnull & "','" _
                                            & adpernull & "'," _
                                            & addiscnull & "," _
                                            & surchargenull & ")"
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
              
                    rs.MoveNext
                   Loop
              End If
            
        Next
     Else
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
             
         For f = lstissuesel.ListCount - 1 To 0 Step -1
            
             Sqlqry = "Select * from flatplanmonthtra where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(Cbomonthfrom.ListIndex) & " AND monthind<= " & Val(CbomonthTo.ListIndex) + Val(Cbomonthfrom.ListIndex) & " and issue_no='" & Trim(Mid(lstissuesel.List(f), 1, 6)) & "' and sub_media ='" & Trim(Cbosubmedia) & "' order by val(page),serial_no"
             Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
             If rs.RecordCount <> 0 Then
              rs.MoveFirst
                      '  x = 0
                      '  x = rs.RecordCount
                      '   crdtamt = 0
                      '   crdtamteach = 0
                         
                      '   Sqlqry2 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                      '   Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                      '   If IsNull(rs2.Fields(0)) = False Then crdtamt = rs2.Fields(0)
                      '   crdtamteach = Round(crdtamt / x, 2)
 
 
                
              Do Until rs.EOF
              
                agcomnull = 0
                         adpernull = 0
                         surchargenull = 0
                         addiscnull = 0
                         
                         If IsNull(rs!agcom) = True Then
                             agcomnull = 0
                         Else
                            agcomnull = rs!agcom
                         End If
                         
                         If IsNull(rs!adper) = True Then
                             adpernull = 0
                         Else
                            adpernull = rs!adper
                         End If
                         
                         If IsNull(rs!addisc) = True Then
                             addiscnull = 0
                         Else
                            addiscnull = rs!addisc
                         End If
                         
                         If IsNull(rs!surcharge) = True Then
                             surchargenull = 0
                         Else
                            surchargenull = rs!surcharge
                         End If
                         
                 
                 
                 Sqlqry1 = " Insert into flatplanMonth values('" & rs!serial_no & "','" & rs!Year & "','" _
                                            & Trim(rs!Month) & "','" _
                                            & findfirstfixup(UCase(rs!Product)) & "','" _
                                            & findfirstfixup(UCase(rs!client)) & "','" _
                                            & findfirstfixup(UCase(rs!Agency)) & "','" & Trim(UCase(rs!Media)) & "','" _
                                            & UCase(Trim(rs!sub_Media)) & "','" _
                                            & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                            & Trim(rs!issue_no) & "','" _
                                            & Trim(rs!Page) & "','" _
                                            & findfirstfixup(Trim(rs!Description)) & "','" _
                                            & findfirstfixup(Trim(rs!Comments)) & "','" _
                                            & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                            & UCase(Trim(rs!Space)) & "','" _
                                            & UCase(Trim(rs!Type)) & "','" _
                                            & UCase(Trim(rs!tcurrency)) & "'," _
                                            & Trim(rs!tconvertion) & "," _
                                            & Trim(rs!tra_amount) & "," _
                                            & Trim(rs!NET_Amount) & ",'" _
                                            & agcomnull & "','" _
                                            & adpernull & "'," _
                                            & addiscnull & "," _
                                            & surchargenull & ")"
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
              
                    rs.MoveNext
                   Loop
              End If
            
            Next
        End If
        
        
        Sqlqry = "Select * from bO_MAS WHERE CANCELL='Y'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
          rs.MoveFirst
             Do Until rs.EOF
                  Sqlqry1 = "DELETE * FROM flatplanMonth WHERE SERIAL_NO='" & rs!serial_no & "' "
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
              
               rs.MoveNext
              Loop
        End If
        
        Sqlqry = "Select * from flatplanmonth"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
          rs.MoveFirst
          Do Until rs.EOF
                If rs!Type = "Paid" Then
                  Sqlqry1 = " Insert into flatplan values('" & rs!serial_no & "','" & rs!Year & "','" _
                                            & Trim(rs!Month) & "','" _
                                            & findfirstfixup(UCase(rs!Product)) & "','" _
                                            & findfirstfixup(UCase(rs!client)) & "','" _
                                            & findfirstfixup(UCase(rs!Agency)) & "','" & Trim(UCase(rs!Media)) & "','" _
                                            & UCase(Trim(rs!sub_Media)) & "','" _
                                            & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                            & Trim(rs!issue_no) & "','" _
                                            & Trim(rs!Page) & "','" _
                                            & findfirstfixup(Trim(rs!Description)) & "','" _
                                            & findfirstfixup(Trim(rs!Comments)) & "','" _
                                            & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                            & UCase(Trim(rs!Space)) & "','" _
                                            & UCase(Trim(rs!Type)) & "','" _
                                            & UCase(Trim(rs!tcurrency)) & "'," _
                                            & Trim(rs!tconvertion) & "," _
                                            & Trim(rs!tra_amount) & "," _
                                            & Trim(rs!NET_Amount) & "," & 0 & ",'" _
                                            & rs!agcom & "','" _
                                            & rs!adper & "'," _
                                            & rs!addisc & "," _
                                            & rs!surcharge & ")"
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
              
                 
                
                Else
                  Sqlqry1 = " Insert into flatplan values('" & rs!serial_no & "','" & rs!Year & "','" _
                                            & Trim(rs!Month) & "','" _
                                            & findfirstfixup(UCase(rs!Product)) & "','" _
                                            & findfirstfixup(UCase(rs!client)) & "','" _
                                            & findfirstfixup(UCase(rs!Agency)) & "','" & Trim(UCase(rs!Media)) & "','" _
                                            & UCase(Trim(rs!sub_Media)) & "','" _
                                            & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                            & Trim(rs!issue_no) & "','" _
                                            & Trim(rs!Page) & "','" _
                                            & findfirstfixup(Trim(rs!Description)) & "','" _
                                            & findfirstfixup(Trim(rs!Comments)) & "','" _
                                            & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                            & UCase(Trim(rs!Space)) & "','" _
                                            & UCase(Trim(rs!Type)) & "','" _
                                            & UCase(Trim(rs!tcurrency)) & "'," _
                                            & Trim(rs!tconvertion) & "," & 0 & "," & 0 & "," _
                                            & Trim(rs!tra_amount) & ",'" _
                                            & rs!agcom & "','" _
                                            & rs!adper & "'," _
                                            & rs!addisc & "," _
                                            & rs!surcharge & ")"
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
              
                
                End If
                
               rs.MoveNext
              Loop
        End If
         

 With CrystalReport1
  .DataFiles(0) = App.Path & "\misov.mdb"
  .ReportFileName = App.Path & "\fpwithoutamount.rpt"
  .Formulas(0) = "zzz='" & "Month From " & Trim(Cbomonthfrom.Text) & " " & Trim(Cboyear) & " To " & Trim(CbomonthTo.Text) & " " & Trim(Cboyear) & "'"
  .Formulas(1) = "yyy='" & Trim(Cbosubmedia) & "'"
  .WindowState = crptMaximized
  .Action = 1
 End With
 
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub cmdClear_Click()
textclear
End Sub
Private Sub Form_Load()
Dim X

Cbomonthfrom.AddItem "January"
Cbomonthfrom.AddItem "February"
Cbomonthfrom.AddItem "March"
Cbomonthfrom.AddItem "April"
Cbomonthfrom.AddItem "May"
Cbomonthfrom.AddItem "June"
Cbomonthfrom.AddItem "July"
Cbomonthfrom.AddItem "August"
Cbomonthfrom.AddItem "September"
Cbomonthfrom.AddItem "October"
Cbomonthfrom.AddItem "November"
Cbomonthfrom.AddItem "December"

i = 2000

For i = 2000 To 2100
 Cboyear.AddItem i
Next
X = 0

 Cboyear.Text = Year(Now())
 
 X = Month(Now())
  
If X = 1 Then
   Cbomonthfrom.ListIndex = 0
ElseIf X = 2 Then
   Cbomonthfrom.ListIndex = 1
ElseIf X = 3 Then
   Cbomonthfrom.ListIndex = 2
ElseIf X = 4 Then
   Cbomonthfrom.ListIndex = 3
ElseIf X = 5 Then
   Cbomonthfrom.ListIndex = 4
ElseIf X = 6 Then
   Cbomonthfrom.ListIndex = 5
ElseIf X = 7 Then
   Cbomonthfrom.ListIndex = 6
ElseIf X = 8 Then
   Cbomonthfrom.ListIndex = 7
ElseIf X = 9 Then
   Cbomonthfrom.ListIndex = 8
ElseIf X = 10 Then
   Cbomonthfrom.ListIndex = 9
ElseIf X = 11 Then
   Cbomonthfrom.ListIndex = 10
Else
   Cbomonthfrom.ListIndex = 11
End If

populateMedia

OptDate.Value = True
  
txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
txtdateto.TextWithMask = Format(Now, "dd/mm/yyyy")

End Sub
Private Sub populateMedia()
 Cbosubmedia.Clear
 cbodtsubmedia.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select sub_media from Media where media_type='Magazine' order by sub_media"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        rs.MoveFirst
           Cbosubmedia.AddItem "All"
           cbodtsubmedia.AddItem "All"
          Do Until rs.EOF
           Cbosubmedia.AddItem Trim(rs!sub_Media)
           cbodtsubmedia.AddItem Trim(rs!sub_Media)
           rs.MoveNext
          Loop
    End If
End Sub
Private Sub Cmdg_Click()
 For i = lstissue.ListCount - 1 To 0 Step -1
    If lstissue.Selected(i) Then
       lstissuesel.AddItem lstissue.List(i)
       lstissue.RemoveItem (i)
    End If
 Next
End Sub

Private Sub Cmdgg_Click()
  For i = lstissue.ListCount - 1 To 0 Step -1
         lstissuesel.AddItem lstissue.List(i)
         lstissue.RemoveItem (i)
  Next i

End Sub

Private Sub Cmdl_Click()
 For f = lstissuesel.ListCount - 1 To 0 Step -1
    
    If lstissuesel.Selected(f) Then
       lstissue.AddItem lstissuesel.Text
       lstissuesel.RemoveItem (f)
    End If
 Next
End Sub

Private Sub Cmdll_Click()
 For i = lstissuesel.ListCount - 1 To 0 Step -1
         lstissue.AddItem lstissuesel.List(i)
         lstissuesel.RemoveItem (i)
 Next i
End Sub

Private Sub Cmddtg_Click()
 For i = lstdtissue.ListCount - 1 To 0 Step -1
    If lstdtissue.Selected(i) Then
       lstdtissuesel.AddItem lstdtissue.List(i)
       lstdtissue.RemoveItem (i)
    End If
 Next
End Sub

Private Sub Cmddtgg_Click()
  For i = lstdtissue.ListCount - 1 To 0 Step -1
         lstdtissuesel.AddItem lstdtissue.List(i)
         lstdtissue.RemoveItem (i)
  Next i

End Sub

Private Sub Cmddtl_Click()
 For f = lstdtissuesel.ListCount - 1 To 0 Step -1
    
    If lstdtissuesel.Selected(f) Then
       lstdtissue.AddItem lstdtissuesel.Text
       lstdtissuesel.RemoveItem (f)
    End If
 Next
End Sub

Private Sub Cmddtll_Click()
 For i = lstdtissuesel.ListCount - 1 To 0 Step -1
         lstdtissue.AddItem lstdtissuesel.List(i)
         lstdtissuesel.RemoveItem (i)
 Next i
End Sub
Private Sub textclear()
    lstissue.Clear
    lstissuesel.Clear
End Sub

Private Sub lstdtissue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstdtissuesel.SetFocus
End Sub

Private Sub lstdtissuesel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmddtprint.SetFocus
End Sub
Private Sub OptDate_Click()
    FraMonth.Visible = False
    Fradate.Visible = True
    framonthsel.Visible = False
    Fradatesel.Visible = True
End Sub

Private Sub OptMonth_Click()
    FraMonth.Visible = True
    Fradate.Visible = False
    framonthsel.Visible = True
    Fradatesel.Visible = False
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdateto.SetFocus
End Sub

Private Sub txtdatefrom_LostFocus()
If IsDate(txtdatefrom.TextWithMask) = False Then
      MsgBox "Invalid Date from ", vbInformation, "Invalid Entry"
      txtdatefrom.SetFocus
      SendKeys "{Home} + {End}"
End If
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbodtsubmedia.SetFocus
End Sub

Private Sub txtdateto_LostFocus()
If IsDate(txtdateto.TextWithMask) = False Then
      MsgBox "Invalid Date to ", vbInformation, "Invalid Entry"
      txtdateto.SetFocus
      SendKeys "{Home} + {End}"
End If
End Sub
