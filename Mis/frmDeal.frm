VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmDeal 
   BackColor       =   &H80000004&
   ClientHeight    =   8595
   ClientLeft      =   15
   ClientTop       =   345
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "                                           Deal /Agency                                    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   8415
      Left            =   240
      TabIndex        =   36
      Top             =   120
      Width           =   11295
      Begin VB.ComboBox CboCurrency 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   390
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Frame Framedia 
         BackColor       =   &H80000004&
         Caption         =   "      Media Allocation   "
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
         Height          =   2535
         Left            =   3360
         TabIndex        =   64
         Top             =   1920
         Width           =   3015
         Begin VB.TextBox txtCinema 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   1560
            TabIndex        =   5
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtMagazine 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   1560
            TabIndex        =   6
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtTelevision 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   1560
            TabIndex        =   7
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtOnline 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   1560
            TabIndex        =   8
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "Cinema"
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
            TabIndex        =   68
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "Magazine"
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
            Height          =   375
            Left            =   240
            TabIndex        =   67
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "Television"
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
            TabIndex        =   66
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "Online"
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
            TabIndex        =   65
            Top             =   1560
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000004&
         Caption         =   "Last  2 years Deal Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   1695
         Left            =   240
         TabIndex        =   34
         Top             =   5400
         Width           =   10815
         Begin VB.TextBox txtyear1budget 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   3480
            TabIndex        =   20
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtyear1actgr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   5160
            TabIndex        =   21
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtyear1actnet 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   6600
            TabIndex        =   22
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtyear1freeallowed 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   8040
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtyear2budget 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   3480
            TabIndex        =   27
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtyear2actgr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   5160
            TabIndex        =   28
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtyear2actnet 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   6600
            TabIndex        =   29
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtyear2freeallowed 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   8040
            TabIndex        =   30
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtyear2freetaken 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   9480
            TabIndex        =   31
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtyear1freetaken 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   9480
            TabIndex        =   24
            Top             =   720
            Width           =   1215
         End
         Begin PVMaskEditLib.PVMaskEdit txtyear2datefrom 
            Height          =   195
            Left            =   480
            TabIndex        =   25
            Top             =   1200
            Width           =   1335
            _Version        =   65541
            _ExtentX        =   2355
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
         Begin PVMaskEditLib.PVMaskEdit txtyear1datefrom 
            Height          =   195
            Left            =   480
            TabIndex        =   18
            Top             =   720
            Width           =   1335
            _Version        =   65541
            _ExtentX        =   2355
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
         Begin PVMaskEditLib.PVMaskEdit txtyear2dateto 
            Height          =   255
            Left            =   2040
            TabIndex        =   26
            Top             =   1200
            Width           =   1335
            _Version        =   65541
            _ExtentX        =   2355
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
         Begin PVMaskEditLib.PVMaskEdit txtyear1dateto 
            Height          =   255
            Left            =   2040
            TabIndex        =   19
            Top             =   720
            Width           =   1335
            _Version        =   65541
            _ExtentX        =   2355
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
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "2"
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
            Left            =   120
            TabIndex        =   62
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "1"
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
            Left            =   120
            TabIndex        =   61
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label14 
            BackColor       =   &H80000004&
            Caption         =   "Free Allow."
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
            Height          =   255
            Left            =   8040
            TabIndex        =   60
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000004&
            Caption         =   "Budget (Gr.)"
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
            Height          =   255
            Left            =   3480
            TabIndex        =   59
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000004&
            Caption         =   "Actual (Net)"
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
            Height          =   375
            Left            =   6600
            TabIndex        =   58
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label17 
            BackColor       =   &H80000004&
            Caption         =   "Actual (Gr.)"
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
            Height          =   255
            Left            =   5160
            TabIndex        =   57
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label20 
            BackColor       =   &H80000004&
            Caption         =   "Free Taken"
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
            Height          =   255
            Left            =   9480
            TabIndex        =   56
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "DateFrom"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   600
            TabIndex        =   55
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "DateTo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2040
            TabIndex        =   54
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   3600
         Picture         =   "frmDeal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   7320
         Width           =   1335
      End
      Begin VB.TextBox txtRemarks 
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
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   4560
         Width           =   9375
      End
      Begin VB.TextBox txtdealname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   720
         Width           =   7455
      End
      Begin VB.TextBox txtBudget 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   1560
         TabIndex        =   4
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000004&
         Caption         =   "              Volume Rebate             "
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
         Height          =   2535
         Left            =   6720
         TabIndex        =   42
         Top             =   1920
         Width           =   4215
         Begin VB.TextBox txtvol4disc 
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
            Left            =   3480
            MaxLength       =   2
            TabIndex        =   16
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtvol3disc 
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
            Left            =   3480
            MaxLength       =   2
            TabIndex        =   14
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtvol2disc 
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
            Left            =   3480
            MaxLength       =   2
            TabIndex        =   12
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtvol1disc 
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
            Left            =   3480
            MaxLength       =   2
            TabIndex        =   10
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtVol4 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            TabIndex        =   15
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtVol3 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            TabIndex        =   13
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtvol2 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            TabIndex        =   11
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtvol1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            TabIndex        =   9
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "%"
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
            Left            =   2880
            TabIndex        =   50
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "%"
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
            Left            =   2880
            TabIndex        =   49
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "%"
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
            Left            =   2880
            TabIndex        =   48
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "%"
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
            Left            =   2880
            TabIndex        =   47
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "Volume 4"
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
            Left            =   120
            TabIndex        =   46
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "Volume 3"
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
            Left            =   120
            TabIndex        =   45
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "Volume 2"
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
            Left            =   120
            TabIndex        =   44
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "Volume 1"
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
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.ComboBox CboAgency 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   390
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1320
         Width           =   5775
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
         Height          =   945
         Left            =   6240
         Picture         =   "frmDeal.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   7320
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
         Height          =   945
         Left            =   4920
         Picture         =   "frmDeal.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   7320
         Width           =   1335
      End
      Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   2280
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
      Begin PVMaskEditLib.PVMaskEdit txtdateto 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   3000
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
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Currency"
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
         Height          =   375
         Left            =   8160
         TabIndex        =   70
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblvoucherno 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
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
         Left            =   9480
         TabIndex        =   63
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Remarks"
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
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Deal Name"
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
         Left            =   120
         TabIndex        =   52
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Budget"
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
         Height          =   375
         Left            =   240
         TabIndex        =   51
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Agency"
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
         Height          =   375
         Left            =   360
         TabIndex        =   41
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "DateFrom"
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
         Left            =   120
         TabIndex        =   40
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "DateTo"
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
         TabIndex        =   39
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   11280
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Label lblMedianame 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   6840
         TabIndex        =   38
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblSubMediaName 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   6840
         TabIndex        =   37
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmDeal.frx":0CC6
      Left            =   840
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
   End
End
Attribute VB_Name = "frmDeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim i As Integer
Dim X, Y, Z As Integer
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Private Sub CboAgency_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboCurrency.SetFocus
End Sub
Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdatefrom.SetFocus
End Sub
Private Sub CmdBack_Click()
 Unload Me
End Sub
Private Sub cmdClear_Click()
 textclear
End Sub
Private Sub CmdSave_Click()

If ValidateData = True Then

   If Val(txtBudget) <> Val(txtCinema) + Val(txtMagazine) + Val(txtOnline) + Val(txtTelevision) Then
      MsgBox " Budget is Not allocated properly to all media"
      txtCinema.SetFocus
      Exit Sub
   End If
  
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Select * from deal where Name='" & findfirstfixup(Trim(UCase(txtdealname))) & "' "
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount <> 0 Then
      MsgBox " Deal Already Exists"
      Exit Sub
    Else
     Sqlqry1 = " Insert into deal values('" & Val(lblvoucherno) & "','" & findfirstfixup(Trim(UCase(txtdealname))) & "','" _
              & Trim(CboAgency) & "','" & Trim(CboCurrency) & "','" _
              & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','" _
              & Format(txtdateto.TextWithMask, "dd/mm/yyyy") & "'," _
              & Val(Trim(txtBudget)) & "," _
              & Val(Trim(txtCinema)) & "," _
              & Val(Trim(txtMagazine)) & "," _
              & Val(Trim(txtOnline)) & "," _
              & Val(Trim(txtTelevision)) & "," _
              & Trim(Val(txtvol1)) & "," & Trim(Val(txtvol1disc)) & "," _
              & Trim(Val(txtvol2)) & "," & Trim(Val(txtvol2disc)) & "," _
              & Trim(Val(txtVol3)) & "," & Trim(Val(txtvol3disc)) & "," _
              & Trim(Val(txtVol4)) & "," & Trim(Val(txtvol4disc)) & ",'" _
              & findfirstfixup(Trim(txtRemarks)) & "','" _
              & Format(txtyear1datefrom.TextWithMask, "dd/mm/yyyy") & "','" _
              & Format(txtyear1dateto.TextWithMask, "dd/mm/yyyy") & "'," _
              & Trim(Val(txtyear1budget)) & "," & Trim(Val(txtyear1actgr)) & "," & Trim(Val(txtyear1actnet)) & "," & Trim(Val(txtyear1freeallowed)) & "," & Trim(Val(txtyear1freetaken)) & ",'" _
              & Format(txtyear2datefrom.TextWithMask, "dd/mm/yyyy") & "','" _
              & Format(txtyear2dateto.TextWithMask, "dd/mm/yyyy") & "'," _
              & Trim(Val(txtyear2budget)) & "," & Trim(Val(txtyear2actgr)) & "," & Trim(Val(txtyear2actnet)) & "," & Trim(Val(txtyear2freeallowed)) & "," & Trim(Val(txtyear2freetaken)) & ")"
                ws.BeginTrans
                db.Execute (Sqlqry1)
                ws.CommitTrans
                
                 MsgBox "Record is inserted", vbDefaultButton3, "Status"
                      
                Set ws = DBEngine.Workspaces(0)
                Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                Sqlqry = "Update docu_mas set doc_no='" & lblvoucherno & "' where doc_type='DEA'"
                ws.BeginTrans
                db.Execute (Sqlqry)
                ws.CommitTrans
                
                 lblvoucherno = lblvoucherno + 1
                 textclear
                 PopulateAgencycodes
                Exit Sub
            End If
        Else
          MsgBox "Information not properly keyned", vbDefaultButton1, "Improper data"
     Exit Sub
  End If
  
End Sub
Private Sub Form_Load()
  PopulateAgencycodes
  AutoIncrementnumber
  CboCurrency.AddItem "DHS"
  CboCurrency.AddItem "USD"
End Sub
Private Sub PopulateAgencycodes()
    CboAgency.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from agndtls Order by AgentName"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
      rs.MoveFirst
        CboAgency.Clear
          Do Until rs.EOF
            CboAgency.AddItem rs!agentname
            rs.MoveNext
        Loop
    End If
        
End Sub

Private Function ValidateData()

ValidateData = False
If txtdealname.Text = "" Then
   MsgBox "Invalid Deal Name", vbInformation, "Invalid Entry"
   txtdealname.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
   
ElseIf CboCurrency.Text = "" Then
   MsgBox "Select Currency", vbInformation, "Invalid Entry"
   CboCurrency.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf IsDate(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsDate(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid To Date", vbInformation, "Invalid Entry"
   txtdateto.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf Val(txtBudget.Text) = 0 Then
   MsgBox "Enter Budget Amount", vbInformation, "Invalid Entry"
   txtBudget.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf CboAgency.Text = "" Then
   MsgBox "Select Agency", vbInformation, "Invalid Entry"
   CboAgency.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
   
Else
  ValidateData = True
End If
End Function

Private Sub textclear()
txtdealname.Text = ""
CboAgency.ListIndex = -1
CboCurrency.ListIndex = -1
txtdatefrom.TextWithMask = ""
txtdateto.TextWithMask = ""
txtyear1datefrom.TextWithMask = ""
txtyear1dateto.TextWithMask = ""
txtyear2datefrom.TextWithMask = ""
txtyear2dateto.TextWithMask = ""
txtBudget.Text = ""
txtRemarks.Text = ""
txtvol1.Text = ""
txtvol1disc.Text = ""
txtvol2.Text = ""
txtvol2disc.Text = ""
txtVol3.Text = ""
txtvol3disc.Text = ""
txtVol4.Text = ""
txtvol4disc.Text = ""
txtyear1budget.Text = ""
txtyear1actgr.Text = ""
txtyear1actnet.Text = ""
txtyear1freeallowed.Text = ""
txtyear1freetaken.Text = ""
txtyear2budget.Text = ""
txtyear2actgr.Text = ""
txtyear2actnet.Text = ""
txtyear2freeallowed.Text = ""
txtyear2freetaken.Text = ""
txtCinema.Text = ""
txtMagazine.Text = ""
txtTelevision.Text = ""
txtOnline.Text = ""
End Sub
Private Sub AutoIncrementnumber()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select DOC_no from DOCU_MAS WHERE DOC_TYPE='DEA'"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
If rs.RecordCount = 0 Then
   MsgBox "Document type 'DEA' not found"
   Exit Sub
Else
   rs.MoveLast
   lblvoucherno = Val(rs!doc_no) + 1
End If
End Sub

Private Sub txtBudget_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCinema.SetFocus
End Sub

Private Sub txtCinema_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtMagazine.SetFocus
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdateto.SetFocus
End Sub

Private Sub txtdatefrom_LostFocus()
If IsDate(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtBudget.SetFocus
End Sub

Private Sub txtdateto_LostFocus()
    If IsDate(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) = False Then
       MsgBox "Invalid To Date", vbInformation, "Invalid Entry"
       txtdateto.SetFocus
       SendKeys " {Home} + {End} "
    End If
End Sub

Private Sub txtdealname_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CboAgency.SetFocus
End Sub

Private Sub txtMagazine_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtOnline.SetFocus
End Sub

Private Sub txtOnline_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtTelevision.SetFocus
End Sub

Private Sub txtremarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear1datefrom.SetFocus
End Sub

Private Sub txtTelevision_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtvol1.SetFocus
End Sub

Private Sub txtvol1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtvol1disc.SetFocus
End Sub

Private Sub txtvol1disc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtvol2.SetFocus
End Sub

Private Sub txtvol2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtvol2disc.SetFocus
End Sub

Private Sub txtvol2disc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtVol3.SetFocus
End Sub

Private Sub txtVol3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtvol3disc.SetFocus
End Sub

Private Sub txtvol3disc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtVol4.SetFocus
End Sub

Private Sub txtVol4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtvol4disc.SetFocus
End Sub

Private Sub txtvol4disc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub

Private Sub txtyear1actgr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear1actnet.SetFocus
End Sub

Private Sub txtyear1actnet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear1freeallowed.SetFocus
End Sub

Private Sub txtyear1budget_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear1actgr.SetFocus
End Sub

Private Sub txtyear1datefrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear1dateto.SetFocus
End Sub

Private Sub txtyear1datefrom_LostFocus()
If IsDate(Format(txtyear1datefrom.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid Year 1 From Date", vbInformation, "Invalid Entry"
   txtyear1datefrom.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtyear1dateto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear1budget.SetFocus
End Sub

Private Sub txtyear1dateto_LostFocus()
If IsDate(Format(txtyear1dateto.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid Year 1 to Date", vbInformation, "Invalid Entry"
   txtyear1dateto.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtyear1freeallowed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear1freetaken.SetFocus
End Sub

Private Sub txtyear1freetaken_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear2datefrom.SetFocus
End Sub

Private Sub txtyear2actgr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear2actnet.SetFocus
End Sub

Private Sub txtyear2actnet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear2freeallowed.SetFocus
End Sub

Private Sub txtyear2budget_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear2actgr.SetFocus
End Sub

Private Sub txtyear2datefrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear2dateto.SetFocus
End Sub

Private Sub txtyear2datefrom_LostFocus()
If IsDate(Format(txtyear2datefrom.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid Year 2 From Date", vbInformation, "Invalid Entry"
   txtyear2datefrom.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtyear2dateto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear2budget.SetFocus
End Sub

Private Sub txtyear2dateto_LostFocus()
    If IsDate(Format(txtyear1dateto.TextWithMask, "dd/mm/yyyy")) = False Then
       MsgBox "Invalid Year 1 to Date", vbInformation, "Invalid Entry"
       txtyear1dateto.SetFocus
       SendKeys " {Home} + {End} "
    End If
End Sub

Private Sub txtyear2freeallowed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtyear2freetaken.SetFocus
End Sub

Private Sub txtyear2freetaken_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdSave.SetFocus
End Sub
