VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmAgencyClientPrdTo 
   BackColor       =   &H80000005&
   Caption         =   "AgencyClientProductTO"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmAgencyClientProdTO.frx":0000
      Left            =   11040
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "                                 Turnover - Agency / Client / Product                                     "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8535
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton CmdRegMed 
         BackColor       =   &H000000FF&
         Caption         =   ">"
         Height          =   195
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   840
         Width           =   255
      End
      Begin VB.Frame FraProduct 
         BackColor       =   &H80000009&
         Caption         =   "     Product"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1695
         Left            =   120
         TabIndex        =   49
         Top             =   6720
         Width           =   7575
         Begin VB.CommandButton CmdClPr 
            BackColor       =   &H000000FF&
            Caption         =   ">"
            Height          =   195
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   0
            Width           =   255
         End
         Begin VB.ListBox LstProductFrom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   55
            Top             =   240
            Width           =   3375
         End
         Begin VB.ListBox LstProductTo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   4200
            MultiSelect     =   2  'Extended
            TabIndex        =   54
            Top             =   240
            Width           =   3255
         End
         Begin VB.CommandButton CmdPRDG 
            Caption         =   ">"
            Height          =   195
            Left            =   3600
            TabIndex        =   53
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton CmdPrdGG 
            Caption         =   ">>"
            Height          =   195
            Left            =   3600
            TabIndex        =   52
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton CmdPrdLL 
            Caption         =   "<<"
            Height          =   195
            Left            =   3600
            TabIndex        =   51
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton CmdPRDL 
            Caption         =   "<"
            Height          =   195
            Left            =   3600
            TabIndex        =   50
            Top             =   1320
            Width           =   375
         End
      End
      Begin VB.ComboBox cboyear 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame FraMedia 
         BackColor       =   &H80000009&
         Caption         =   "     Media"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1335
         Left            =   5160
         TabIndex        =   40
         Top             =   840
         Width           =   6375
         Begin VB.ListBox LstMediaFrom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   46
            Top             =   240
            Width           =   2895
         End
         Begin VB.ListBox LstMediaTo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   3600
            MultiSelect     =   2  'Extended
            TabIndex        =   45
            Top             =   240
            Width           =   2655
         End
         Begin VB.CommandButton CmdMediaG 
            Caption         =   ">"
            Height          =   195
            Left            =   3120
            TabIndex        =   44
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton CmdMediaGG 
            Caption         =   ">>"
            Height          =   195
            Left            =   3120
            TabIndex        =   43
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton CmdMediaLL 
            Caption         =   "<<"
            Height          =   195
            Left            =   3120
            TabIndex        =   42
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton CmdMediaL 
            Caption         =   "<"
            Height          =   195
            Left            =   3120
            TabIndex        =   41
            Top             =   960
            Width           =   375
         End
      End
      Begin VB.Frame FraClient 
         BackColor       =   &H80000009&
         Caption         =   "     Client"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2175
         Left            =   120
         TabIndex        =   33
         Top             =   4440
         Width           =   7575
         Begin VB.CommandButton CmdAgCl 
            BackColor       =   &H000000FF&
            Caption         =   ">"
            Height          =   195
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton CmdClL 
            Caption         =   "<"
            Height          =   195
            Left            =   3600
            TabIndex        =   39
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton CmdClLL 
            Caption         =   "<<"
            Height          =   195
            Left            =   3600
            TabIndex        =   38
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton CmdClGG 
            Caption         =   ">>"
            Height          =   195
            Left            =   3600
            TabIndex        =   37
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton CmdClG 
            Caption         =   ">"
            Height          =   195
            Left            =   3600
            TabIndex        =   36
            Top             =   360
            Width           =   375
         End
         Begin VB.ListBox LstClientTo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   4200
            MultiSelect     =   2  'Extended
            TabIndex        =   35
            Top             =   240
            Width           =   3255
         End
         Begin VB.ListBox LstClientFrom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   34
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000009&
         Caption         =   "       Agency"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2055
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   11415
         Begin VB.CommandButton CmdMedAg 
            BackColor       =   &H000000FF&
            Caption         =   ">"
            Height          =   195
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   0
            Width           =   255
         End
         Begin VB.ListBox LstAgencyFrom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1620
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   32
            Top             =   240
            Width           =   5175
         End
         Begin VB.ListBox LstAgencyTo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1620
            Left            =   6240
            MultiSelect     =   2  'Extended
            TabIndex        =   31
            Top             =   240
            Width           =   4935
         End
         Begin VB.CommandButton CmdAgG 
            Caption         =   ">"
            Height          =   195
            Left            =   5520
            TabIndex        =   30
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton CmdAgGG 
            Caption         =   ">>"
            Height          =   195
            Left            =   5520
            TabIndex        =   29
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton CmdAGLL 
            Caption         =   "<<"
            Height          =   195
            Left            =   5520
            TabIndex        =   28
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton CmdAGL 
            Caption         =   "<"
            Height          =   195
            Left            =   5520
            TabIndex        =   27
            Top             =   1440
            Width           =   375
         End
      End
      Begin VB.Frame FraRegion 
         BackColor       =   &H80000009&
         Caption         =   "Region"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1335
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   4815
         Begin VB.CommandButton CmdRegL 
            Caption         =   "<"
            Height          =   195
            Left            =   2400
            TabIndex        =   25
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton CmdRegLL 
            Caption         =   "<<"
            Height          =   195
            Left            =   2400
            TabIndex        =   24
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton CmdRegGG 
            Caption         =   ">>"
            Height          =   195
            Left            =   2400
            TabIndex        =   23
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton CmdRegG 
            Caption         =   ">"
            Height          =   195
            Left            =   2400
            TabIndex        =   22
            Top             =   240
            Width           =   375
         End
         Begin VB.ListBox LstRegionTo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   2880
            MultiSelect     =   2  'Extended
            TabIndex        =   21
            Top             =   240
            Width           =   1815
         End
         Begin VB.ListBox LstRegionFrom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   20
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
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
         ForeColor       =   &H8000000D&
         Height          =   3975
         Left            =   7920
         TabIndex        =   13
         Top             =   4440
         Width           =   1935
         Begin VB.OptionButton OptClient 
            BackColor       =   &H80000009&
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
            Left            =   240
            TabIndex        =   60
            Top             =   1680
            Width           =   1335
         End
         Begin VB.OptionButton OptMonth 
            BackColor       =   &H80000009&
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
            Left            =   240
            TabIndex        =   18
            Top             =   2880
            Width           =   1335
         End
         Begin VB.OptionButton OptRegion 
            BackColor       =   &H80000009&
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
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   3480
            Width           =   1215
         End
         Begin VB.OptionButton OptAgency 
            BackColor       =   &H80000009&
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
            Left            =   240
            TabIndex        =   16
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton OptProduct 
            BackColor       =   &H80000009&
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
            Left            =   240
            TabIndex        =   15
            Top             =   2280
            Width           =   1335
         End
         Begin VB.OptionButton OptSubMedia 
            BackColor       =   &H80000009&
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
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.CommandButton Cmdexcel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Excel"
         Height          =   375
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   8040
         Width           =   615
      End
      Begin VB.ComboBox cboCurrency 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00E0E0E0&
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
         Left            =   10200
         Picture         =   "frmAgencyClientProdTO.frx":001D
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
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
         Left            =   10200
         Picture         =   "frmAgencyClientProdTO.frx":045F
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5880
         Width           =   1095
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00E0E0E0&
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
         Left            =   10200
         Picture         =   "frmAgencyClientProdTO.frx":08A1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6840
         Width           =   1095
      End
      Begin VB.ComboBox cbomonthTo 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cbomonthfrom 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label LblMediaName 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   48
         Top             =   3840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Height          =   255
         Left            =   8400
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblSubMediaName 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   3600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "To"
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
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Month From"
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
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "  Year"
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
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAgencyClientPrdTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim i As Integer
Dim d, e, f, g As Integer
Dim C, X, Y, Z As Integer
Dim adddisc As Currency
Dim scharge As Currency
Dim ntra As Currency
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim sqlqry3 As String
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim rs3 As Recordset
Dim crdtamt As Currency
Dim crdtper As Currency
Dim crdtgross As Currency
Dim Addiscamt As Currency
Dim totaddiscamt As Currency
Dim fmname As String
Dim fmid As String
Public n, m

Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then LstRegionFrom.SetFocus
End Sub

        
Private Sub cbomonthfrom_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbomonthTo.SetFocus
End Sub

Private Sub cbomonthTo_GotFocus()
cbomonthTo.Clear
If cbomonthfrom.ListIndex = 0 Then
    cbomonthTo.AddItem "January"
    cbomonthTo.AddItem "February"
    cbomonthTo.AddItem "March"
    cbomonthTo.AddItem "April"
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 1 Then
    cbomonthTo.AddItem "February"
    cbomonthTo.AddItem "March"
    cbomonthTo.AddItem "April"
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 2 Then
    cbomonthTo.AddItem "March"
    cbomonthTo.AddItem "April"
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 3 Then
    cbomonthTo.AddItem "April"
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 4 Then
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 5 Then
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 6 Then
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 7 Then
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 8 Then
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 9 Then
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 10 Then
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 11 Then
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
Else
    cbomonthTo.SetFocus
End If
End Sub
Private Sub cbomonthTo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboCurrency.SetFocus
End Sub
Private Sub cboyear_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbomonthfrom.SetFocus
End Sub

Private Sub CmdAgCl_Click()
Dim f As Integer

If LstRegionTo.ListCount = 0 Then
  MsgBox "You must select atleast one region"
  LstRegionFrom.SetFocus
  Exit Sub
End If

  
If LstMediaTo.ListCount = 0 Then
   MsgBox "Select Media"
   LstMediaFrom.SetFocus
   Exit Sub
End If

If LstAgencyTo.ListCount = 0 Then
   MsgBox "Select Agency"
   LstAgencyFrom.SetFocus
   Exit Sub
End If


    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Delete * from Dumbo_Mascl"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans



 f = LstAgencyTo.ListIndex
 'For f = 0 To lstdtissuesel.ListCount - 1
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 For f = 0 To LstAgencyTo.ListCount - 1
     Sqlqry = "Select * from dumbo_masag where year='" & Val(cboyear.Text) & "' and agency='" & Trim(LstAgencyTo.List(f)) & "' order by client"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                      
                      Sqlqry = " Insert into DumBo_Mascl values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & rs!tra_gamount & "," & rs!tra_namount & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & rs!gross_amount & "," _
                                             & rs!Tot_free & "," _
                                             & rs!Tot_barter & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & rs!add_discount & "," & rs!surcharge & "," _
                                             & rs!NET_Amount & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
              Next
     populateclient
End Sub

Private Sub cmdBack_Click()
 Unload Me
End Sub
Private Sub cmdClear_Click()
 textclear
End Sub

Private Sub CmdClPr_Click()
Dim f As Integer

If LstRegionTo.ListCount = 0 Then
  MsgBox "You must select atleast one region"
  LstRegionFrom.SetFocus
  Exit Sub
End If

If LstMediaTo.ListCount = 0 Then
   MsgBox "Select Media"
   LstMediaFrom.SetFocus
   Exit Sub
End If

If LstAgencyTo.ListCount = 0 Then
   MsgBox "Select Agency"
   LstAgencyFrom.SetFocus
   Exit Sub
End If

If LstClientTo.ListCount = 0 Then
   MsgBox "Select Client"
   LstClientFrom.SetFocus
   Exit Sub
End If

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = " Delete * from Dumbo_Maspr"
  ws.BeginTrans
  db.Execute (Sqlqry)
  ws.CommitTrans

 f = LstClientTo.ListIndex
 'For f = 0 To lstdtissuesel.ListCount - 1
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 For f = 0 To LstClientTo.ListCount - 1
     Sqlqry = "Select * from dumbo_mascl where year='" & Val(cboyear.Text) & "' and client='" & findfirstfixup(Trim(LstClientTo.List(f))) & "' order by product"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                      
                      Sqlqry1 = " Insert into DumBo_Maspr values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & rs!tra_gamount & "," & rs!tra_namount & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & rs!gross_amount & "," _
                                             & rs!Tot_free & "," _
                                             & rs!Tot_barter & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & rs!add_discount & "," & rs!surcharge & "," _
                                             & rs!NET_Amount & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (Sqlqry1)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
              Next
     populateproducts
     
End Sub

Private Sub cmdDisplay_Click()
  Dim l, o, p As String

Dim uname As String
Dim compname As String
Dim objnet
Dim fmname
Dim fmid
Dim temp

fmname = Me.Caption
fmid = Me.Name

temp = False
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
     
    If rs!lock_status = "Y" Then
            uname = rs!u_name
            MsgBox "Table has been locked exclusively by the user." & uname
            Exit Sub
        
            cmdDisplay.SetFocus
    
    
    Else
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
    
    
                'cmdModify.Enabled = Fal
     End If
   End If



  
  Populatereport
  Cinadjustments
  curadjustments
  ReportSort
    
  checkout
  
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
Private Sub ReportSort()

         If OptSubMedia.Value = True Then
            With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\agclprsm.rpt"
                .Formulas(0) = "yyy='" & Val(cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "Cur='" & Trim(cboCurrency.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
         ElseIf OptProduct.Value = True Then
             With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\AgclprPr"
                .Formulas(0) = "yyy='" & Val(cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "Cur='" & Trim(cboCurrency.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
           ElseIf OptMonth.Value = True Then
             With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\AgclprMo"
                .Formulas(0) = "yyy='" & Val(cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "Cur='" & Trim(cboCurrency.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
          ElseIf OptAgency.Value = True Then
             With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Agclprag"
                .Formulas(0) = "yyy='" & Val(cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "Cur='" & Trim(cboCurrency.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
          ElseIf OptClient.Value = True Then
             With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Agclprcl"
                .Formulas(0) = "yyy='" & Val(cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "Cur='" & Trim(cboCurrency.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
         Else
          'Region
            With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Agclprre"
                .Formulas(0) = "yyy='" & Val(cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "Cur='" & Trim(cboCurrency.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
          End If
         
          
End Sub
Private Sub Cinadjustments()

Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = " Delete * from DumBo_MasCin"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
              
Sqlqry = " Delete * from DumBo_MasCin2"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     
        Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry = "Select * from dumbo_masrep where Media='Cinema'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry1 = " Insert into DumBo_MasCin values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & rs!tra_gamount & "," & rs!tra_namount & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & rs!gross_amount & "," _
                                             & rs!Tot_free & "," _
                                             & rs!Tot_barter & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & rs!add_discount & "," & rs!surcharge & "," _
                                             & rs!NET_Amount & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (Sqlqry1)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
           
     Sqlqry = " Delete * from Dumbo_MasRep where media='Cinema' "
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
              
              
        adddisc = 0
        scharge = 0
        ntra = 0
                
                       
              
        Sqlqry = "Select * from DumBo_MasCin"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
              
                    
                Do Until rs.EOF
                 
                adddisc = 0
                scharge = 0
                ntra = 0
                
                 Sqlqry1 = "Select * from bo_tracin where serial_no='" & Trim(rs!serial_no) & "' AND TYPE ='Paid'"
                 Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                    If rs1.RecordCount = 0 Then
                        adddisc = rs!add_discount
                        scharge = rs!surcharge
                    Else
                        adddisc = rs!add_discount / rs1.RecordCount
                        scharge = rs!surcharge / rs1.RecordCount
                    End If
                
                
                
                
                Sqlqry1 = "Select * from bo_tracin where serial_no='" & Trim(rs!serial_no) & "'  "
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                 If rs1.RecordCount <> 0 Then
                  rs1.MoveFirst
                   Do Until rs1.EOF
                     If rs1!Type = "Paid" Then
                       
                            If rs!add_discount = 0 Then
                              ntra = Val(rs1!tra_amount) - (Val(rs1!tra_amount) * rs!disc_rate / 100) - (((rs1!tra_amount) - (rs1!tra_amount * rs!disc_rate / 100)) * rs!disc_percentage / 100)
                            Else
                              ntra = Val(rs1!tra_amount) - (Val(rs1!tra_amount) * rs!disc_rate / 100) - (((rs1!tra_amount) - (rs1!tra_amount * rs!disc_rate / 100)) * rs!disc_percentage / 100) - Val(adddisc)
                            End If
        ' check this carefully suspected error
                       '   If rs1!tcurrency = " USD" Then
                       '    ntra = ntra * Val(rs1!tconvertion)
                       '   End If
        ' check this carefully suspected error
        
                          
                        Sqlqry2 = " Insert into DumBo_Masrep values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & rs!tra_gamount & "," & rs!tra_namount & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs1!Media & "','" & rs1!Media & " " & rs1!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & rs1!tra_amount & "," _
                                             & 0 & "," _
                                             & 0 & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & adddisc & "," & 0 & "," _
                                             & ntra & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (Sqlqry2)
                         ws.CommitTrans
                          
                      ElseIf rs1!Type = "Free" Then
                          
                        Sqlqry2 = " Insert into DumBo_Masrep values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & rs!tra_gamount & "," & rs!tra_namount & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs1!Media & "','" & rs1!Media & " " & rs1!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & 0 & "," _
                                             & rs1!tra_amount & "," _
                                             & 0 & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & 0 & "," & 0 & "," _
                                             & 0 & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (Sqlqry2)
                         ws.CommitTrans
                          
                       Else
                         Sqlqry2 = " Insert into DumBo_Masrep values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & rs!tra_gamount & "," & rs!tra_namount & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs1!Media & "','" & rs!Media & " " & rs1!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & 0 & "," _
                                             & 0 & "," _
                                             & rs1!tra_amount & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & 0 & "," & 0 & "," _
                                             & 0 & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (Sqlqry2)
                         ws.CommitTrans
                        
                        End If
                     
                  rs1.MoveNext
                  Loop
                 End If
               rs.MoveNext
               Loop
           End If
           
     Sqlqry = "Select * from dumbo_masCin"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry1 = " Insert into DumBo_MasCin2 values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & rs!tra_gamount & "," & rs!tra_namount & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & rs!gross_amount & "," _
                                             & rs!Tot_free & "," _
                                             & rs!Tot_barter & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & rs!add_discount & "," & rs!surcharge & "," _
                                             & rs!NET_Amount & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (Sqlqry1)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
      
     
     
     
End Sub

Private Sub Cmdexcel_Click()
    Dim objxl As Object
    Dim result As Variant
    Set objxl = CreateObject("Excel.application")
    objxl.Workbooks.Open FileName:=App.Path & "\text.xls"
    objxl.Visible = True
    objxl.Run "udtest"
End Sub

Private Sub CmdMedAg_Click()
Dim f As Integer

If LstRegionTo.ListCount = 0 Then
  MsgBox "You must select atleast one region"
  LstRegionFrom.SetFocus
  Exit Sub
End If

'If LstMediaFrom.ListCount = 0 Then Exit Sub
   
If LstMediaTo.ListCount = 0 Then
   MsgBox "You must select Media"
   LstMediaFrom.SetFocus
   Exit Sub
End If


Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = " Delete * from Dumbo_Masag"
 ws.BeginTrans
 db.Execute (Sqlqry)
 ws.CommitTrans


 f = LstMediaTo.ListIndex
 'For f = 0 To lstdtissuesel.ListCount - 1
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 For f = 0 To LstMediaTo.ListCount - 1
     Sqlqry = "Select * from dumbo_masreg where year='" & Val(cboyear.Text) & "' and sub_media='" & Trim(LstMediaTo.List(f)) & "' order by agency"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                      
                      Sqlqry1 = " Insert into DumBo_Masag values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & rs!tra_gamount & "," & rs!tra_namount & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & rs!gross_amount & "," _
                                             & rs!Tot_free & "," _
                                             & rs!Tot_barter & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & rs!add_discount & "," & rs!surcharge & "," _
                                             & rs!NET_Amount & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (Sqlqry1)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
              Next
     populateagency
   
End Sub

Private Sub CmdRegMed_Click()
Dim f As Integer

ValidateData

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Delete * from Dumbo_Masreg"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans


If LstRegionTo.ListCount = 0 Then
  MsgBox "You must select atleast one region"
  LstRegionFrom.SetFocus
  Exit Sub
End If


 f = LstRegionTo.ListIndex
 'For f = 0 To lstdtissuesel.ListCount - 1
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 For f = 0 To LstRegionTo.ListCount - 1
     Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' and region='" & Trim(LstRegionTo.List(f)) & "'  and CANCELL='N' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " order by sub_media"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                      
                      Sqlqry1 = " Insert into DumBo_MasReg values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & rs!tra_gamount & "," & rs!tra_namount & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & rs!gross_amount & "," _
                                             & rs!Tot_free & "," _
                                             & rs!Tot_barter & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & rs!add_discount & "," & rs!surcharge & "," _
                                             & rs!NET_Amount & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (Sqlqry1)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
              Next
              
              
     populateMedia
                
End Sub

Private Sub Form_Load()
    cbomonthfrom.AddItem "January"
    cbomonthfrom.AddItem "February"
    cbomonthfrom.AddItem "March"
    cbomonthfrom.AddItem "April"
    cbomonthfrom.AddItem "May"
    cbomonthfrom.AddItem "June"
    cbomonthfrom.AddItem "July"
    cbomonthfrom.AddItem "August"
    cbomonthfrom.AddItem "September"
    cbomonthfrom.AddItem "October"
    cbomonthfrom.AddItem "November"
    cbomonthfrom.AddItem "December"
   
      
   
   cboCurrency.AddItem "DHS"
   cboCurrency.AddItem "USD"
 
Populateregion

i = 2000

For i = 2000 To 2100
 cboyear.AddItem i
Next
X = 0

 cboyear.Text = Year(Now())
 
 X = Month(Now())
  
If X = 1 Then
   cbomonthfrom.ListIndex = 0
   'cbomonthTo.ListIndex = 0
ElseIf X = 2 Then
   cbomonthfrom.ListIndex = 1
   'cbomonthTo.ListIndex = 1
ElseIf X = 3 Then
   cbomonthfrom.ListIndex = 2
   'cbomonthTo.ListIndex = 2
ElseIf X = 4 Then
   cbomonthfrom.ListIndex = 3
   'cbomonthTo.ListIndex = 3
ElseIf X = 5 Then
   cbomonthfrom.ListIndex = 4
   'cbomonthTo.ListIndex = 4
ElseIf X = 6 Then
   cbomonthfrom.ListIndex = 5
   'cbomonthTo.ListIndex = 5
ElseIf X = 7 Then
   cbomonthfrom.ListIndex = 6
   'cbomonthTo.ListIndex = 6
ElseIf X = 8 Then
   cbomonthfrom.ListIndex = 7
   'cbomonthTo.ListIndex = 7
ElseIf X = 9 Then
   cbomonthfrom.ListIndex = 8
   'cbomonthTo.ListIndex = 8
ElseIf X = 10 Then
   cbomonthfrom.ListIndex = 9
   'cbomonthTo.ListIndex = 9
ElseIf X = 11 Then
   cbomonthfrom.ListIndex = 10
   'cbomonthTo.ListIndex = 10
Else
   cbomonthfrom.ListIndex = 11
   'cbomonthTo.ListIndex = 11
End If

End Sub

Private Sub Populateregion()
    LstRegionFrom.Clear
    LstRegionTo.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select distinct(region) from bo_mas Order by region"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        rs.MoveFirst
       Do Until rs.EOF
        If IsEmpty(rs!region) = True Then
         rs.MoveNext
        Else
         LstRegionFrom.AddItem rs!region
         rs.MoveNext
        End If
       Loop
    End If
 End Sub
Private Sub Populatereport()
Dim f As Integer

If LstRegionTo.ListCount = 0 Then
  MsgBox "You must select atleast one region"
  LstRegionFrom.SetFocus
  Exit Sub
End If


If LstMediaTo.ListCount = 0 Then
   MsgBox "Select Media"
   LstMediaFrom.SetFocus
   Exit Sub
End If


If LstAgencyTo.ListCount = 0 Then
   MsgBox "Select Agency"
   LstAgencyFrom.SetFocus
   Exit Sub
End If



If LstClientTo.ListCount = 0 Then
   MsgBox "Select Client"
   LstClientFrom.SetFocus
   Exit Sub
End If



If LstProductTo.ListCount = 0 Then
   MsgBox "Select Product"
   LstProductFrom.SetFocus
   Exit Sub
End If

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Delete * from Dumbo_Masrep"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans




 f = LstProductTo.ListIndex
 'For f = 0 To lstdtissuesel.ListCount - 1
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 For f = 0 To LstProductTo.ListCount - 1
     Sqlqry = "Select * from dumbo_maspr where year='" & Val(cboyear.Text) & "' and Product='" & Trim(LstProductTo.List(f)) & "'"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                      
                      Sqlqry1 = " Insert into DumBo_Masrep values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & rs!tra_gamount & "," & rs!tra_namount & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & rs!gross_amount & "," _
                                             & rs!Tot_free & "," _
                                             & rs!Tot_barter & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & rs!add_discount & "," & rs!surcharge & "," _
                                             & rs!NET_Amount & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (Sqlqry1)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
              Next
     
     
End Sub
Private Sub populateproducts()
    LstProductFrom.Clear
    LstProductTo.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select distinct(Product) from dumBo_maspr Order by Product"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
     
        rs.MoveFirst
            Do Until rs.EOF
              LstProductFrom.AddItem rs!Product
            rs.MoveNext
       Loop
    End If
 End Sub

Private Sub populateMedia()
    LstMediaFrom.Clear
    LstMediaTo.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select distinct(sub_media) from dumbo_masreg Order by sub_media"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
          rs.MoveFirst
            Do Until rs.EOF
              If rs!sub_media <> "Cinema" Then
               LstMediaFrom.AddItem Trim(rs!sub_media)
               rs.MoveNext
              Else
               rs.MoveNext
              End If
            Loop
    End If
 End Sub
Private Sub curadjustments()
Dim crdtgross
Dim crdtper
Dim crdtamt
Dim trano

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Delete * from dumbo_mascur"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     
              
     If cboCurrency.Text = "USD" Then
       
        Sqlqry = "Select * from dumbo_masrep where Tcurrency='DHS'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                    
                    Sqlqry2 = "select * from DumBo_masCin2 where Serial_no='" & rs!serial_no & "'"
                    Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                    trano = rs2.RecordCount
                                            
                    
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "' and tcurrency='DHS'"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                    If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                          If crdtamt > 0 Then
                           crdtamt = crdtamt / trano
                           crdtamt = Round(crdtamt / convertion, 3)
                          Else
                           crdtamt = 0
                          End If
                          
                     sqlqry3 = " Insert into DumBo_Mascur values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & Round(rs!tra_gamount / convertion, 3) & "," & Round(rs!tra_namount / convertion, 3) & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Round(rs!gross_amount / convertion, 3) & "," _
                                             & Round(rs!Tot_free / convertion, 3) & "," _
                                             & Round(rs!Tot_barter / convertion, 3) & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & Round(rs!add_discount / convertion, 3) + crdtamt & "," & Round(rs!surcharge / convertion, 3) & "," _
                                             & Round(rs!NET_Amount / convertion, 3) - crdtamt & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (sqlqry3)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
           
           
        Sqlqry = "Select * from dumbo_masrep where Tcurrency='USD'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                    Sqlqry2 = "select * from DumBo_masCin2 where Serial_no='" & rs!serial_no & "'"
                    Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                    trano = rs2.RecordCount
                      
                                            
                    
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "' and tcurrency='USD'"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                    If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                          If crdtamt > 0 Then
                            If trano <> 0 Then
                              crdtamt = crdtamt / trano
                            End If
                           crdtamt = Round(crdtamt, 3)
                          Else
                           crdtamt = 0
                          End If
                  
                  
                     sqlqry3 = " Insert into DumBo_Mascur values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & rs!tra_gamount & "," & rs!tra_namount & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & rs!gross_amount & "," _
                                             & rs!Tot_free & "," _
                                             & rs!Tot_barter & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & rs!add_discount + crdtamt & "," & rs!surcharge & "," _
                                             & rs!NET_Amount - crdtamt & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (sqlqry3)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
               
           
   Else
     
        Sqlqry = "Select * from dumbo_masrep where Tcurrency='USD'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                    Sqlqry2 = "select * from DumBo_masCin2 where Serial_no='" & rs!serial_no & "'"
                    Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                    trano = rs2.RecordCount
                                            
                    
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "' and tcurrency='USD'"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                    If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                          If crdtamt > 0 Then
                           crdtamt = crdtamt / trano
                           crdtamt = Round(crdtamt * convertion, 3)
                          Else
                           crdtamt = 0
                          End If
                  
                      
                     sqlqry3 = " Insert into DumBo_Mascur values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & Round(rs!tra_gamount * convertion, 3) & "," & Round(rs!tra_namount * convertion, 3) & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Round(rs!gross_amount * convertion, 3) & "," _
                                             & Round(rs!Tot_free * convertion, 3) & "," _
                                             & Round(rs!Tot_barter * convertion, 3) & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & Round(rs!add_discount * convertion, 3) - crdtamt & "," & Round(rs!surcharge * convertion, 3) & "," _
                                             & Round(rs!NET_Amount * convertion, 3) + crdtamt & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (sqlqry3)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
           
           
        Sqlqry = "Select * from dumbo_masrep where Tcurrency='DHS'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                    Sqlqry2 = "select * from DumBo_masCin2 where Serial_no='" & rs!serial_no & "'"
                    Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                    trano = rs2.RecordCount
                                            
                    
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "' and tcurrency='DHS'"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                    If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                          If crdtamt > 0 Then
                           crdtamt = crdtamt / trano
                           crdtamt = Round(crdtamt, 3)
                          Else
                           crdtamt = 0
                          End If
                  
                     
                     Sqlqry1 = " Insert into DumBo_Mascur values('" & rs!serial_no & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & rs!tra_gamount & "," & rs!tra_namount & ",'" & rs!Year & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" _
                                             & findfirstfixup(rs!region) & "','" & findfirstfixup(rs!boremarks) & "','" & findfirstfixup(rs!acremarks) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & rs!gross_amount & "," _
                                             & rs!Tot_free & "," _
                                             & rs!Tot_barter & ",'" _
                                             & rs!disc_percentage & "','" _
                                             & rs!disc_rate & "'," _
                                             & rs!add_discount + crdtamt & "," & rs!surcharge & "," _
                                             & rs!NET_Amount - crdtamt & ",'" & rs!invoice_date & "','" & rs!acct_code & "','" & rs!Status & "','" & rs!cancell & "')"
            
                         ws.BeginTrans
                         db.Execute (Sqlqry1)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
      End If
                 
End Sub

Private Sub populateagency()
    LstAgencyFrom.Clear
    LstAgencyTo.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select distinct(Agency) from dumbo_masag Order by Agency"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
      rs.MoveFirst
        LstAgencyFrom.Clear
         
        Do Until rs.EOF
            LstAgencyFrom.AddItem rs!Agency
            rs.MoveNext
        Loop
    End If
        
End Sub
Private Sub populateclient()
    LstClientFrom.Clear
    LstClientTo.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select distinct(client) from dumbo_mascl Order by client"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
      rs.MoveFirst
        LstClientFrom.Clear
         
        Do Until rs.EOF
            LstClientFrom.AddItem rs!client
            rs.MoveNext
        Loop
    End If
        
End Sub

Private Function ValidateData()

ValidateData = False
If cboyear.Text = "" Then
   MsgBox "Invalid year", vbInformation, "Invalid Entry"
   cboyear.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
 ElseIf cbomonthfrom.Text = "" Then
   MsgBox "Select Month From", vbInformation, "Invalid Entry"
   cbomonthfrom.SetFocus
   SendKeys " {Home} + {end} "
   Exit Function
 ElseIf cbomonthTo.Text = "" Then
   MsgBox "Select Month To", vbInformation, "Invalid Entry"
   cbomonthTo.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
 ElseIf cboCurrency.Text = "" Then
   MsgBox "Select Currency", vbInformation, "Invalid Entry"
   cboCurrency.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
   
 
Else
  ValidateData = True
End If
End Function
Private Sub Cmdregg_Click()
   For i = LstRegionFrom.ListCount - 1 To 0 Step -1
       If LstRegionFrom.Selected(i) Then
          LstRegionTo.AddItem LstRegionFrom.List(i)
          LstRegionFrom.RemoveItem (i)
       End If
   Next
End Sub
Private Sub Cmdreggg_Click()
    For i = LstRegionFrom.ListCount - 1 To 0 Step -1
           LstRegionTo.AddItem LstRegionFrom.List(i)
           LstRegionFrom.RemoveItem (i)
    Next i
End Sub
Private Sub Cmdregl_Click()
    For f = LstRegionTo.ListCount - 1 To 0 Step -1
       If LstRegionTo.Selected(f) Then
          LstRegionFrom.AddItem LstRegionTo.Text
          LstRegionTo.RemoveItem (f)
       End If
    Next
End Sub
Private Sub Cmdregll_Click()
    For i = LstRegionTo.ListCount - 1 To 0 Step -1
            LstRegionFrom.AddItem LstRegionTo.List(i)
            LstRegionTo.RemoveItem (i)
    Next i
End Sub

Private Sub textclear()
 cboyear.ListIndex = -1
 cbomonthfrom.ListIndex = -1
 cbomonthTo.ListIndex = -1
 LstRegionFrom.Clear
 LstRegionTo.Clear
 LstMediaFrom.Clear
 LstMediaTo.Clear
 LstAgencyFrom.Clear
 LstAgencyTo.Clear
 LstClientFrom.Clear
 LstClientTo.Clear
 LstProductFrom.Clear
 LstProductTo.Clear
 
End Sub
Private Sub Cmdmediag_Click()
    For i = LstMediaFrom.ListCount - 1 To 0 Step -1
       If LstMediaFrom.Selected(i) Then
          LstMediaTo.AddItem LstMediaFrom.List(i)
          LstMediaFrom.RemoveItem (i)
       End If
    Next
End Sub
Private Sub Cmdmediagg_Click()
    For i = LstMediaFrom.ListCount - 1 To 0 Step -1
           LstMediaTo.AddItem LstMediaFrom.List(i)
           LstMediaFrom.RemoveItem (i)
    Next i
End Sub
Private Sub Cmdmedial_Click()
    For f = LstMediaTo.ListCount - 1 To 0 Step -1
       If LstMediaTo.Selected(f) Then
          LstMediaFrom.AddItem LstMediaTo.Text
          LstMediaTo.RemoveItem (f)
       End If
    Next
End Sub
Private Sub Cmdmediall_Click()
    For i = LstMediaTo.ListCount - 1 To 0 Step -1
            LstMediaFrom.AddItem LstMediaTo.List(i)
            LstMediaTo.RemoveItem (i)
    Next i
End Sub

Private Sub Cmdagg_Click()
    For i = LstAgencyFrom.ListCount - 1 To 0 Step -1
       If LstAgencyFrom.Selected(i) Then
          LstAgencyTo.AddItem LstAgencyFrom.List(i)
          LstAgencyFrom.RemoveItem (i)
       End If
    Next
End Sub
Private Sub Cmdaggg_Click()
    For i = LstAgencyFrom.ListCount - 1 To 0 Step -1
           LstAgencyTo.AddItem LstAgencyFrom.List(i)
           LstAgencyFrom.RemoveItem (i)
    Next i
End Sub
Private Sub Cmdagl_Click()
    For f = LstAgencyTo.ListCount - 1 To 0 Step -1
       If LstAgencyTo.Selected(f) Then
          LstAgencyFrom.AddItem LstAgencyTo.Text
          LstAgencyTo.RemoveItem (f)
       End If
    Next
End Sub
Private Sub Cmdagll_Click()
    For i = LstAgencyTo.ListCount - 1 To 0 Step -1
            LstAgencyFrom.AddItem LstAgencyTo.List(i)
            LstAgencyTo.RemoveItem (i)
    Next i
End Sub

Private Sub Cmdclg_Click()
    For i = LstClientFrom.ListCount - 1 To 0 Step -1
       If LstClientFrom.Selected(i) Then
          LstClientTo.AddItem LstClientFrom.List(i)
          LstClientFrom.RemoveItem (i)
       End If
    Next
End Sub
Private Sub Cmdclgg_Click()
    For i = LstClientFrom.ListCount - 1 To 0 Step -1
           LstClientTo.AddItem LstClientFrom.List(i)
           LstClientFrom.RemoveItem (i)
    Next i
End Sub
Private Sub Cmdcll_Click()
    For f = LstClientTo.ListCount - 1 To 0 Step -1
       If LstClientTo.Selected(f) Then
          LstClientFrom.AddItem LstClientTo.Text
          LstClientTo.RemoveItem (f)
       End If
    Next
End Sub
Private Sub Cmdclll_Click()
    For i = LstClientTo.ListCount - 1 To 0 Step -1
            LstClientFrom.AddItem LstClientTo.List(i)
            LstClientTo.RemoveItem (i)
    Next i
End Sub

Private Sub Cmdprdg_Click()
    For i = LstProductFrom.ListCount - 1 To 0 Step -1
       If LstProductFrom.Selected(i) Then
          LstProductTo.AddItem LstProductFrom.List(i)
          LstProductFrom.RemoveItem (i)
       End If
    Next
End Sub
Private Sub Cmdprdgg_Click()
    For i = LstProductFrom.ListCount - 1 To 0 Step -1
           LstProductTo.AddItem LstProductFrom.List(i)
           LstProductFrom.RemoveItem (i)
    Next i
End Sub
Private Sub Cmdprdl_Click()
    For f = LstProductTo.ListCount - 1 To 0 Step -1
       If LstProductTo.Selected(f) Then
          LstProductFrom.AddItem LstProductTo.Text
          LstProductTo.RemoveItem (f)
       End If
    Next
End Sub
Private Sub Cmdprdll_Click()
    For i = LstProductTo.ListCount - 1 To 0 Step -1
            LstProductFrom.AddItem LstProductTo.List(i)
            LstProductTo.RemoveItem (i)
    Next i
End Sub
