VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmAgency 
   BackColor       =   &H80000005&
   Caption         =   "Agency Details"
   ClientHeight    =   8370
   ClientLeft      =   -75
   ClientTop       =   225
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   11850
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   10920
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   855
      Left            =   1680
      TabIndex        =   41
      Top             =   7680
      Width           =   6615
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Add"
         DisabledPicture =   "frmAgency.frx":0000
         DownPicture     =   "frmAgency.frx":0532
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
         Left            =   120
         MaskColor       =   &H008080FF&
         Picture         =   "frmAgency.frx":0A64
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdMod 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Modify"
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
         Left            =   1200
         Picture         =   "frmAgency.frx":0EA6
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "C&lear"
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
         Left            =   4440
         Picture         =   "frmAgency.frx":12E8
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00E0E0E0&
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
         Left            =   5400
         Picture         =   "frmAgency.frx":13EA
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Delete"
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
         Left            =   2280
         Picture         =   "frmAgency.frx":191C
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
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
         Height          =   660
         Left            =   3360
         Picture         =   "frmAgency.frx":1D5E
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.ListBox lstAgencies 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   7350
      Left            =   8040
      TabIndex        =   40
      Top             =   120
      Width           =   3735
   End
   Begin VB.Frame frmAgency 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7935
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
         Caption         =   "Opening Balance"
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
         Height          =   735
         Left            =   120
         TabIndex        =   43
         Top             =   6720
         Width           =   7455
         Begin VB.TextBox txtopDhs 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5040
            TabIndex        =   32
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtopUSD 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2160
            TabIndex        =   31
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H80000005&
            Caption         =   "DHS"
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
            Left            =   4440
            TabIndex        =   45
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label13 
            BackColor       =   &H80000005&
            Caption         =   "USD"
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
            Left            =   1560
            TabIndex        =   44
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.TextBox txtDiscount 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   30
         Top             =   6240
         Width           =   375
      End
      Begin VB.TextBox txtFax 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   39
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox txtweb 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   29
         Top             =   6240
         Width           =   1935
      End
      Begin VB.TextBox txtOffTel 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   28
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox txtAreaCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   27
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Frame FraAddress 
         BackColor       =   &H80000005&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1935
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   6735
         Begin VB.TextBox txtPOBox 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2160
            TabIndex        =   18
            Top             =   480
            Width           =   3615
         End
         Begin VB.TextBox txtcity 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2160
            TabIndex        =   17
            Top             =   960
            Width           =   3615
         End
         Begin VB.TextBox txtCountry 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2160
            TabIndex        =   16
            Top             =   1440
            Width           =   3615
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000005&
            Caption         =   "  P.O.Box"
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
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000005&
            Caption         =   "  City"
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
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000005&
            Caption         =   "  Country"
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
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   1440
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000005&
         Caption         =   "                              Name            Mobile        e-mail   "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2055
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   7575
         Begin VB.TextBox txtMobile1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4440
            ScrollBars      =   1  'Horizontal
            TabIndex        =   11
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtMobile2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4440
            ScrollBars      =   1  'Horizontal
            TabIndex        =   10
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtMobile3 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4440
            ScrollBars      =   1  'Horizontal
            TabIndex        =   9
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtmail1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   5760
            ScrollBars      =   1  'Horizontal
            TabIndex        =   8
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtmail2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   5760
            ScrollBars      =   1  'Horizontal
            TabIndex        =   7
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtmail3 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   5760
            ScrollBars      =   1  'Horizontal
            TabIndex        =   6
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txtConName3 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2160
            TabIndex        =   5
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox txtConName2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2160
            ScrollBars      =   1  'Horizontal
            TabIndex        =   4
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtConName1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2160
            ScrollBars      =   1  'Horizontal
            TabIndex        =   3
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label14 
            BackColor       =   &H80000005&
            Caption         =   "  Media Manager"
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
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000005&
            Caption         =   "  Media Director"
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
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000005&
            Caption         =   "  Managing Director"
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
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.TextBox txtAgencyName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   4560
         TabIndex        =   42
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000005&
         Caption         =   "  Web Site"
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
         Height          =   225
         Left            =   360
         TabIndex        =   26
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "  Fax"
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
         Height          =   300
         Left            =   4440
         TabIndex        =   25
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Caption         =   " Agency Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Caption         =   " Telephone (Off)"
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
         Height          =   300
         Left            =   360
         TabIndex        =   23
         Top             =   5760
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000005&
         Caption         =   " Area Code"
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
         Height          =   300
         Left            =   360
         TabIndex        =   22
         Top             =   5160
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAgency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim AgnNm As String
Dim ws As Workspace
Dim db As Database
Dim rs As Recordset



Private Sub Form_Unload(Cancel As Integer)
 Exit Sub
End Sub

Private Sub cmdadd_Click()

  If ValidateData = True Then
  
    Set ws = DBEngine.Workspaces(0)
    ' Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Select * from agndtls where agentname='" & findfirstfixup(Trim(UCase(txtAgencyName))) & "' "
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
         MsgBox " Agency Already existing"
         Exit Sub
        Else
    Sqlqry1 = " Insert into agndtls values('" & findfirstfixup(Trim(UCase(txtAgencyName))) & "','" _
              & Trim(txtPOBox) & "','" _
              & findfirstfixup(Trim(txtcity)) & "','" _
              & findfirstfixup(Trim(txtCountry)) & "','" _
              & Trim(txtOffTel) & "','" _
              & Trim(txtMobile1) & "','" _
              & Trim(txtMobile2) & "','" _
              & Trim(txtMobile3) & "','" _
              & Trim(txtFax) & "','" _
              & Trim(txtmail1) & "','" _
              & Trim(txtmail2) & "','" _
              & Trim(txtmail3) & "','" _
              & Trim(txtweb) & "','" _
              & Trim(txtConName1) & "','" _
              & Trim(txtConName2) & "','" _
              & Trim(txtConName3) & "'," _
              & Val(txtDiscount) & ",'" _
              & Trim(txtAreaCode) & "'," _
              & Val(txtopUSD) & "," & Val(txtopDhs) & ")"
                ws.BeginTrans
                db.Execute (Sqlqry1)
                ws.CommitTrans
                
                 MsgBox "Record is inserted", vbDefaultButton3, "Status"
                 textclear
                 PopulateAgencycodes
                Exit Sub
            End If
        Else
          MsgBox "Information not properly keyned", vbDefaultButton1, "Improper data"
     Exit Sub
  End If
      
End Sub

Private Sub cmdBack_Click()
 Unload Me
End Sub

Private Sub cmdClear_Click()
 textclear
End Sub

Private Function textclear()
txtAgencyName.Text = ""
txtOffTel = ""
txtMobile1 = ""
txtMobile2 = ""
txtMobile3 = ""
txtCountry = ""
txtFax = ""
txtPOBox = ""
txtcity = ""
txtmail1 = ""
txtmail2 = ""
txtweb = ""
txtConName1 = ""
txtConName2 = ""
txtConName3 = ""
txtAreaCode = ""
txtopUSD.Text = ""
txtopDhs.Text = ""
End Function

Private Function ValidateData()

ValidateData = False

If txtAgencyName.Text = "" Or IsNumeric(txtAgencyName) = True Then
   MsgBox "Invalid Agency Name", vbInformation, "Invalid Entry"
   txtAgencyName.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
   
ElseIf txtCountry.Text = "" Or IsNumeric(txtCountry.Text) = True Then
   MsgBox "Invalid Country", vbInformation, "Invalid Entry"
   txtCountry.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf txtcity.Text = "" Or IsNumeric(txtcity.Text) = True Then
   MsgBox "Invalid City", vbInformation, "Invalid Entry"
   txtcity.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
End If
ValidateData = True
End Function

Private Sub cmdDelete_Click()
Dim tempStr
 If lstAgencies.SelCount = 0 Then
        MsgBox "Select the Agency Name for Deletion.", vbInformation, "Selection Error"
        lstAgencies.SetFocus
        Exit Sub
 End If
        If ValidateData = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Delete the Agency Name : " & txtAgencyName, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If DeleteData = False Then Exit Sub
        Else
            MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
            lstAgencies.SetFocus
            Exit Sub
        End If
End Sub

Private Sub cmdMod_Click()
 
 Dim tempStr

    If lstAgencies.SelCount = 0 Then
        MsgBox "Select the Agency Name for Modification.", vbInformation, "Selection Error"
        lstAgencies.SetFocus
        Exit Sub
    End If
        AgnNm = " "
        If ValidateData = False Then Exit Sub
        AgnNm = UCase(lstAgencies.Text)
        tempStr = MsgBox("Do You Want To Modify the Agency Details :" & lstAgencies.Text, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
          If ModifyData = False Then Exit Sub
         
        Else
              MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
              lstAgencies.SetFocus
              Exit Sub
        End If
    End Sub
Private Sub populateagencyname()
 Dim i, j
 i = UCase(Trim(txtAgencyName.Text))
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
    i = Trim(txtAgencyName.Text)
    j = Trim(lstAgencies.Text)
        Sqlqry = "Select agency from bo_mas  where agency ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
           
              Sqlqry1 = "Update bo_mas set Agency = '" & findfirstfixup(i) & "' WHERE AGENCY='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
        Sqlqry = "Select agency from bo_tracin  where agency ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
             Sqlqry1 = "Update bo_tracin set Agency = '" & findfirstfixup(i) & "' WHERE AGENCY='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
            
            
        Sqlqry = "Select agency from bo_tramag  where agency ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
             Sqlqry1 = "Update bo_tramag set Agency = '" & findfirstfixup(i) & "' WHERE AGENCY='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
            
        Sqlqry = "Select agency from bo_tratv  where agency ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
           
              Sqlqry1 = "Update bo_tratv set Agency = '" & findfirstfixup(i) & "' WHERE AGENCY='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
            
        Sqlqry = "Select agency from bo_traol  where agency ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            
              Sqlqry1 = "Update bo_traol set Agency = '" & findfirstfixup(i) & "' WHERE AGENCY='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
            
        Sqlqry = "Select Acct_Name from bpmt_tra  where acct_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            
              Sqlqry1 = "Update bpmt_tra set Acct_name = '" & findfirstfixup(i) & "' WHERE Acct_Name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
        
        
        Sqlqry = "Select Acct_Name from brpt_tra  where acct_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update brpt_tra set Acct_name = '" & findfirstfixup(i) & "' WHERE Acct_Name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
        
        Sqlqry = "Select Acct_Name from cpmt_tra  where acct_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update cpmt_tra set Acct_name = '" & findfirstfixup(i) & "' WHERE Acct_Name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
        
        Sqlqry = "Select Acct_Name from crdt_mas  where acct_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update crdt_mas set Acct_name = '" & findfirstfixup(i) & "' WHERE Acct_Name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
        Sqlqry = "Select Acct_Name from crdt_mas  where supp_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            
              Sqlqry1 = "Update crdt_mas set supp_name = '" & findfirstfixup(i) & "' WHERE Acct_Name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
        
        Sqlqry = "Select Acct_Name from crpr_mas  where supp_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update crpr_mas set supp_name = '" & findfirstfixup(i) & "' WHERE Acct_Name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
        
        
       
         
        Sqlqry = "Select Acct_Name from debt_mas  where acct_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
             Sqlqry1 = "Update debt_mas set Acct_name = '" & findfirstfixup(i) & "' WHERE Acct_Name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
        Sqlqry = "Select Acct_Name from jrnl_tra  where acct_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update jrnl_tra set Acct_name = '" & findfirstfixup(i) & "' WHERE Acct_Name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
        Sqlqry = "Select Acct_Name from ppmt_tra  where acct_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update ppmt_tra set Acct_name = '" & findfirstfixup(i) & "' WHERE Acct_Name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
        Sqlqry = "Select Acct_Name from prpt_mas1  where acct_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update prpt_mas1 set Acct_name = '" & findfirstfixup(i) & "' WHERE Acct_Name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
         
        Sqlqry = "Select agency from deal  where agency ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update deal set Agency = '" & findfirstfixup(i) & "' WHERE AGENCY='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
       
        Sqlqry = "Select cust_name from debt_mas  where cust_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update debt_mas set cust_name = '" & findfirstfixup(i) & "' WHERE cust_name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
       
        Sqlqry = "Select agent_name from productS  where agent_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update productS set Agent_name = '" & findfirstfixup(i) & "' WHERE AGENt_name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
       
End Sub

Private Function ModifyData() As Boolean
    Dim i
    ModifyData = False
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
    i = Trim(txtAgencyName.Text)
           
           Sqlqry = "Update Agndtls Set " _
                  & " Agentname = '" & findfirstfixup(Trim(UCase(txtAgencyName.Text))) & "'," _
                  & " pobox = '" & Trim(txtPOBox.Text) & "'," _
                  & " city = '" & findfirstfixup(Trim(txtcity.Text)) & "'," _
                  & " country = '" & findfirstfixup(Trim(txtCountry.Text)) & "'," _
                  & " tel_off = '" & Trim(txtOffTel.Text) & "'," _
                  & " Mobile1 = '" & Trim(txtMobile1.Text) & "'," _
                  & " Mobile2 = '" & Trim(txtMobile2.Text) & "'," _
                  & " Mobile3 = '" & Trim(txtMobile3.Text) & "'," _
                  & " fax ='" & Trim(txtFax.Text) & "'," _
                  & " e_mail1 = '" & Trim(txtmail1.Text) & "'," _
                  & " e_mail2 = '" & Trim(txtmail2.Text) & "'," _
                  & " e_mail3 = '" & Trim(txtmail3.Text) & "'," _
                  & " web = '" & Trim(txtweb) & "', " _
                  & " name1 = '" & Trim(txtConName1) & "', " _
                  & " name2 = '" & Trim(txtConName2) & "', " _
                  & " name3 = '" & Trim(txtConName3) & "', " _
                  & " Discount = " & Val(txtDiscount) & ", " _
                  & " Area_code = '" & Trim(txtAreaCode) & "', " _
                  & " Op_Usd = " & Val(txtopUSD) & ", " _
                  & " Op_DHS = " & Val(txtopDhs) & " " _
                  & " Where Agentname ='" & findfirstfixup(Trim(lstAgencies.Text)) & "'"
                 ' & " Where Agentname ='" & Trim(txtAgencyName.Text) & "'"
                                                
                                                     
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        
        If Trim(AgnNm) <> UCase(Trim(txtAgencyName)) Then
            populateagencyname
        End If
        MsgBox "Record Modified With " & Chr(10) & Chr(10) & _
               "Agency Name = " & i, vbInformation, "Data Modified"
        textclear
        PopulateAgencycodes
        ModifyData = True
        Exit Function
End Function

Private Function DeleteData() As Boolean
  Dim i
    
    DeleteData = False
    i = Trim(UCase(txtAgencyName.Text))
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "select agency from bo_mas where agency='" & findfirstfixup(i) & "'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              MsgBox "Transactions are recorded, canno delete Agency . . . "
              textclear
              PopulateAgencycodes
              Exit Function
        Else
    
           Sqlqry1 = "Delete * from agndtls Where AgentName = '" & findfirstfixup(i) & "'"
            ws.BeginTrans
            db.Execute (Sqlqry1)
            ws.CommitTrans
            MsgBox "Record Deleted With " & Chr(10) & Chr(10) & _
                    "Agency Name = " & i, vbInformation, "Data Modified"
                    
             textclear
             PopulateAgencycodes
        End If
               
End Function

Private Sub PopulateAgencycodes()
    textclear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from agndtls Order by AgentName"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
      rs.MoveFirst
        lstAgencies.Clear
        Do Until rs.EOF
            lstAgencies.AddItem rs!agentname
            rs.MoveNext
        Loop
    End If
        
End Sub
Private Sub CmdPrint_Click()
    CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
    CrystalReport1.ReportFileName = App.Path & "\AgencyList.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
End Sub
Private Sub Form_Load()
    PopulateAgencycodes
    textclear
End Sub
Private Sub lstagencies_Click()
Dim i
Dim tempBln As String
    If lstAgencies.ListIndex = -1 Then
        tempBln = False
    Else
        tempBln = True
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Trim(lstAgencies.Text)
        Sqlqry = "Select * from agndtls Where agentname= '" & findfirstfixup(i) & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MsgBox "Data Mis Matching", vbInformation, "Deleted Status"
            Exit Sub
         Else
           txtAgencyName = rs!agentname
           
           If IsNull(rs!pobox) = True Then
              txtPOBox = ""
           Else
              txtPOBox = rs!pobox
           End If
                               
           If IsNull(rs!city) = True Then
              txtcity = ""
           Else
              txtcity = rs!city
           End If
           
           If IsNull(rs!country) = True Then
              txtCountry = ""
           Else
              txtCountry = rs!country
           End If
           
          If IsNull(rs!tel_off) = True Then
              txtOffTel = ""
           Else
              txtOffTel = rs!tel_off
          End If
          
          
          If IsNull(rs!mobile1) = True Then
              txtMobile1 = ""
           Else
              txtMobile1 = rs!mobile1
          End If
          
          If IsNull(rs!mobile2) = True Then
              txtMobile2 = ""
           Else
              txtMobile2 = rs!mobile2
          End If
          
          If IsNull(rs!mobile3) = True Then
              txtMobile3 = ""
           Else
              txtMobile3 = rs!mobile3
          End If
          
          If IsNull(rs!fax) = True Then
              txtFax = ""
           Else
              txtFax = rs!fax
          End If
      
          
          If IsNull(rs!E_mail1) = True Then
              txtmail1 = ""
           Else
              txtmail1 = rs!E_mail1
          End If
      
          If IsNull(rs!E_mail2) = True Then
              txtmail2 = ""
           Else
              txtmail2 = rs!E_mail2
          End If
      
          If IsNull(rs!E_mail3) = True Then
              txtmail3 = ""
           Else
              txtmail3 = rs!E_mail3
          End If
      
          If IsNull(rs!web) = True Then
              txtweb = ""
           Else
              txtweb = rs!web
          End If
             
          If IsNull(rs!name1) = True Then
              txtConName1 = ""
           Else
              txtConName1 = rs!name1
          End If
          
          If IsNull(rs!name2) = True Then
              txtConName2 = ""
           Else
              txtConName2 = rs!name2
          End If
          
          If IsNull(rs!name3) = True Then
              txtConName3 = ""
           Else
              txtConName3 = rs!name3
          End If
          
          If IsNull(rs!Discount) = True Then
              txtDiscount = ""
           Else
              txtDiscount = rs!Discount
          End If
               
          If IsNull(rs!Area_Code) = True Then
              txtAreaCode = ""
           Else
              txtAreaCode = rs!Area_Code
          End If
          
                              
          If IsNull(rs!op_USD) = True Then
              txtopUSD = ""
           Else
              txtopUSD = rs!op_USD
          End If
          
          If IsNull(rs!op_DHS) = True Then
              txtopDhs = ""
           Else
              txtopDhs = rs!op_DHS
          End If
          
          
       End If
    
End Sub

Private Sub txtAgencyName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtConName1.SetFocus
End Sub

Private Sub txtAreaCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtOffTel.SetFocus
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCountry.SetFocus
End Sub
Private Sub txtConName1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtMobile1.SetFocus
End Sub
Private Sub txtConName2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtMobile2.SetFocus
End Sub
Private Sub txtConName3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtMobile3.SetFocus
End Sub
Private Sub txtCountry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAreaCode.SetFocus
End Sub
Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtopUSD.SetFocus
End Sub
Private Sub txtFax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtweb.SetFocus
End Sub
Private Sub txtmail1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtConName2.SetFocus
End Sub
Private Sub txtmail2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtConName3.SetFocus
End Sub
Private Sub txtmail3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPOBox.SetFocus
End Sub
Private Sub txtMobile1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmail1.SetFocus
End Sub
Private Sub txtMobile2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmail2.SetFocus
End Sub
Private Sub txtMobile3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmail3.SetFocus
End Sub
Private Sub txtOffTel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtFax.SetFocus
End Sub

Private Sub txtopDhs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub

Private Sub txtopUSD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtopDhs.SetFocus
End Sub

Private Sub txtPOBox_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtcity.SetFocus
End Sub

Private Sub txtweb_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtDiscount.SetFocus
End Sub

Function replacestr(Textin, ByVal searchstr As String, _
                    ByVal Replacement As String, _
                    ByVal CompMode As Integer)

  Dim Worktext As String, Pointer As Integer
   If IsNull(Textin) Then
    replacestr = Null
   Else
    Worktext = Textin
    Pointer = InStr(1, Worktext, searchstr, CompMode)
     Do While Pointer > 0
      Worktext = Left(Worktext, Pointer - 1) & Replacement & _
                 Mid(Worktext, Pointer + Len(searchstr))
                 
      Pointer = InStr(Pointer + Len(Replacement), Worktext, _
                 searchstr, CompMode)
                 
    Loop
    
    replacestr = Worktext
    
  
   End If
End Function

Function sqlfixup(Textin)
 sqlfixup = replacestr(Textin, "'", "''", 0)
End Function
Function jetsqlfixup(Textin)
 Dim Temp
  Temp = replacestr(Textin, "'", "''", 0)
  jetsqlfixup = replacestr(Temp, "|", "' & Chr(124) & '", 0)
End Function
 
Function findfirstfixup(Textin)
 Dim Temp
  Temp = replacestr(Textin, "'", "' & Chr(39) & '", 0)
  findfirstfixup = replacestr(Temp, "|", "' & Chr(124) & '", 0)

End Function


