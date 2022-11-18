VERSION 5.00
Begin VB.Form frmCinemaRates 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Tariff - Cinema"
   ClientHeight    =   8535
   ClientLeft      =   -45
   ClientTop       =   285
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   12045
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   1215
      Left            =   360
      TabIndex        =   24
      Top             =   7080
      Width           =   7575
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFF00&
         Caption         =   "&Add"
         DisabledPicture =   "frmCinemaRates.frx":0000
         DownPicture     =   "frmCinemaRates.frx":0532
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
         Left            =   120
         MaskColor       =   &H008080FF&
         Picture         =   "frmCinemaRates.frx":0A64
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdMod 
         BackColor       =   &H00FFFF00&
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
         Height          =   780
         Left            =   1320
         Picture         =   "frmCinemaRates.frx":0EA6
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFF00&
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
         Height          =   780
         Left            =   4920
         Picture         =   "frmCinemaRates.frx":12E8
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFF00&
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
         Left            =   6120
         Picture         =   "frmCinemaRates.frx":13EA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00FFFF00&
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
         Height          =   780
         Left            =   2520
         Picture         =   "frmCinemaRates.frx":191C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFF00&
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
         Left            =   3720
         Picture         =   "frmCinemaRates.frx":1D5E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ListBox lstsubMedia 
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
      Height          =   7560
      Left            =   8520
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Frame frmClient 
      BackColor       =   &H00FFFFC0&
      Caption         =   "         Cinema  Sub Media   -   Rates and other details            "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6855
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   8295
      Begin VB.ComboBox CboRegion 
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
         Height          =   360
         Left            =   2400
         TabIndex        =   39
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txttype 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2400
         TabIndex        =   4
         Top             =   2640
         Width           =   4455
      End
      Begin VB.TextBox txtseats 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2400
         TabIndex        =   2
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtshows 
         BackColor       =   &H00C0FFFF&
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
         Left            =   5280
         TabIndex        =   3
         Top             =   2040
         Width           =   855
      End
      Begin VB.Frame FraAddress 
         BackColor       =   &H00FFFFC0&
         Caption         =   "                               Rates                                   "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   3255
         Left            =   120
         TabIndex        =   22
         Top             =   3360
         Width           =   8055
         Begin VB.TextBox txtbiw15 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5520
            TabIndex        =   36
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtmon15 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5520
            TabIndex        =   35
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtmonoth 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   6720
            TabIndex        =   14
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtbiwoth 
            BackColor       =   &H00C0FFC0&
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
            Left            =   6720
            TabIndex        =   13
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtmon10 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4440
            TabIndex        =   12
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtbiw10 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4440
            TabIndex        =   11
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtmon90 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3480
            TabIndex        =   10
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtbiw90 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3480
            TabIndex        =   9
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtmon60 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2520
            TabIndex        =   8
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtbiw60 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2520
            TabIndex        =   7
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtmon30 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1560
            TabIndex        =   6
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtbiw30 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1560
            TabIndex        =   5
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Filmlets"
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
            Left            =   5520
            TabIndex        =   37
            Top             =   960
            Width           =   855
         End
         Begin VB.Line Line12 
            X1              =   5280
            X2              =   5280
            Y1              =   600
            Y2              =   2880
         End
         Begin VB.Line Line11 
            X1              =   6600
            X2              =   6600
            Y1              =   600
            Y2              =   2880
         End
         Begin VB.Line Line10 
            X1              =   2400
            X2              =   2400
            Y1              =   600
            Y2              =   2880
         End
         Begin VB.Line Line9 
            X1              =   3360
            X2              =   3360
            Y1              =   600
            Y2              =   2880
         End
         Begin VB.Line Line8 
            X1              =   4320
            X2              =   4320
            Y1              =   600
            Y2              =   2880
         End
         Begin VB.Line Line7 
            X1              =   1440
            X2              =   1440
            Y1              =   600
            Y2              =   2880
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   7920
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line5 
            X1              =   120
            X2              =   7920
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   7920
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line3 
            X1              =   7920
            X2              =   7920
            Y1              =   600
            Y2              =   2880
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   7920
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   120
            Y1              =   600
            Y2              =   2880
         End
         Begin VB.Label Label18 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Slides"
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
            Left            =   4440
            TabIndex        =   31
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label17 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Others"
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
            Left            =   6840
            TabIndex        =   30
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFFFC0&
            Caption         =   "90 Sec"
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
            Left            =   3480
            TabIndex        =   29
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFFC0&
            Caption         =   "60 Sec"
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
            Left            =   2520
            TabIndex        =   28
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFC0&
            Caption         =   "30 Sec"
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
            Left            =   1560
            TabIndex        =   27
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Bi- Weekly"
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
            Left            =   120
            TabIndex        =   26
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "  Monthly"
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
            TabIndex        =   25
            Top             =   2280
            Width           =   1095
         End
      End
      Begin VB.TextBox txtsubMedia 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2400
         TabIndex        =   1
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Region"
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
         Height          =   345
         Left            =   720
         TabIndex        =   38
         Top             =   1440
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Type  of Movies  "
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
         Height          =   315
         Left            =   360
         TabIndex        =   34
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "# of seats "
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
         Height          =   345
         Left            =   720
         TabIndex        =   33
         Top             =   2040
         Width           =   1605
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   " # of shows"
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
         Height          =   315
         Left            =   3600
         TabIndex        =   32
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sub Media"
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
         Height          =   345
         Left            =   720
         TabIndex        =   23
         Top             =   840
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmCinemaRates"
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

Private Sub CboRegion_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtseats.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub
Private Sub cmdadd_Click()

  If ValidateData = True Then
  
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Select * from Cinema_rates where Sub_media='" & findfirstfixup(Trim(txtsubMedia)) & "' "
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
         MsgBox " Rates already existing "
         Exit Sub
        Else
    Sqlqry1 = " Insert into Cinema_rates values('" & findfirstfixup(Trim(txtsubMedia)) & "','" _
              & Trim(CboRegion) & "','" _
              & Trim(txtseats) & "','" _
              & Trim(txtshows) & "','" _
              & Trim(txttype) & "'," _
              & Val(Trim(txtmon30)) & "," _
              & Val(Trim(txtbiw30)) & "," _
              & Val(Trim(txtmon60)) & "," _
              & Val(Trim(txtbiw60)) & "," _
              & Val(Trim(txtmon90)) & "," _
              & Val(Trim(txtbiw90)) & "," _
              & Val(Trim(txtmon10)) & "," _
              & Val(Trim(txtbiw10)) & "," _
              & Val(Trim(txtmon15)) & "," _
              & Val(Trim(txtbiw15)) & "," _
              & Val(Trim(txtmonoth)) & "," _
              & Val(Trim(txtbiwoth)) & ")"
              
      ws.BeginTrans
      db.Execute (Sqlqry1)
      ws.CommitTrans
                
                 MsgBox "Record is inserted", vbDefaultButton3, "Status"
                 textclear
                 Populatesubmedia
                Exit Sub
            End If
        Else
          MsgBox "Improper Data", vbDefaultButton1, "Improper data"
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
txtsubMedia.Text = ""
CboRegion.Text = ""
txtseats = ""
txtshows = ""
txttype = ""
txtmon30 = ""
txtbiw30 = ""
txtmon60 = ""
txtbiw60 = ""
txtmon90 = ""
txtbiw90 = ""
txtmon10 = ""
txtbiw10 = ""
txtmon15 = ""
txtbiw15 = ""
txtmonoth = ""
txtbiwoth = ""
End Function

Private Function ValidateData()

ValidateData = False

If txtsubMedia.Text = "" Or IsNumeric(txtsubMedia) = True Then
   MsgBox "Invalid Sub Media Name", vbInformation, "Invalid Entry"
   txtsubMedia.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
   
ElseIf CboRegion.Text = "" Or IsNumeric(CboRegion) = True Then
   MsgBox "Invalid Region", vbInformation, "Invalid Entry"
   CboRegion.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
   
ElseIf IsNumeric(txtmon30) = False Then
   MsgBox "Invalid Monthly 30Sec Rate", vbInformation, "Invalid Entry"
   txtmon30.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf IsNumeric(txtbiw30) = False Then
   MsgBox "Invalid Bi-Weekly 30Sec Rate", vbInformation, "Invalid Entry"
   txtbiw30.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf IsNumeric(txtmon60) = False Then
   MsgBox "Invalid Monthly 60Sec Rate", vbInformation, "Invalid Entry"
   txtmon60.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf IsNumeric(txtbiw60) = False Then
   MsgBox "Invalid Bi-Weekly 60Sec Rate", vbInformation, "Invalid Entry"
   txtbiw60.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf IsNumeric(txtmon90) = False Then
   MsgBox "Invalid Monthly 90Sec Rate", vbInformation, "Invalid Entry"
   txtmon90.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf IsNumeric(txtbiw90) = False Then
   MsgBox "Invalid Bi-Weekly 90Sec Rate", vbInformation, "Invalid Entry"
   txtbiw90.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf IsNumeric(txtmon10) = False Then
   MsgBox "Invalid Monthly 10Sec Rate", vbInformation, "Invalid Entry"
   txtmon10.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf IsNumeric(txtbiw10) = False Then
   MsgBox "Invalid Bi-Weekly 10Sec Rate", vbInformation, "Invalid Entry"
   txtbiw10.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
   
ElseIf IsNumeric(txtmon15) = False Then
   MsgBox "Invalid Monthly 15Sec Rate", vbInformation, "Invalid Entry"
   txtmon15.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf IsNumeric(txtbiw15) = False Then
   MsgBox "Invalid Bi-Weekly 15Sec Rate", vbInformation, "Invalid Entry"
   txtbiw15.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
   
End If
ValidateData = True
End Function

Private Sub cmdDelete_Click()
Dim tempStr
If lstsubmedia.SelCount = 0 Then
        MsgBox "Select the Sub Media for Deletion.", vbInformation, "Selection Error"
        lstsubmedia.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Delete the Sub Media : " & txtsubMedia, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If DeleteData = False Then Exit Sub
        Else
            MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
            lstsubmedia.SetFocus
            Exit Sub
        End If
End Sub
Private Sub cmdMod_Click()
Dim i
 Dim tempStr
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       i = Trim(lstsubmedia.Text)
    Sqlqry = "Select * from cinema_rates"
    
    If lstsubmedia.SelCount = 0 Then
        MsgBox "Select the Sub Media for Modification.", vbInformation, "Selection Error"
        lstsubmedia.SetFocus
        Exit Sub
    End If
        AgnNm = " "
        If ValidateData = False Then Exit Sub
        AgnNm = lstsubmedia.Text
        tempStr = MsgBox("Do You Want To Modify the Sub Media Details :" & lstsubmedia.Text, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If ModifyData = False Then Exit Sub
        Else
              MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
              lstsubmedia.SetFocus
              Exit Sub
        End If
    End Sub
Private Function ModifyData() As Boolean
    Dim i
    ModifyData = False
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
    i = Trim(lstsubmedia.Text)
    
           
           Sqlqry = "Update Cinema_rates Set " _
                  & " Sub_media = '" & findfirstfixup(Trim(UCase(txtsubMedia.Text))) & "'," _
                  & " Region = '" & Trim(CboRegion.Text) & "'," _
                  & " Seats = '" & Trim(txtseats.Text) & "'," _
                  & " Shows = '" & Trim(txtshows.Text) & "'," _
                  & " Type = '" & Trim(txttype.Text) & "'," _
                  & " Mon30 = " & Val(Trim(txtmon30.Text)) & "," _
                  & " Biw30 = " & Val(Trim(txtbiw30.Text)) & "," _
                  & " Mon60 = " & Val(Trim(txtmon60.Text)) & "," _
                  & " biw60 = " & Val(Trim(txtbiw60.Text)) & "," _
                  & " Mon90 = " & Val(Trim(txtmon90.Text)) & "," _
                  & " biw90 = " & Val(Trim(txtbiw90.Text)) & "," _
                  & " mon10 = " & Val(Trim(txtmon10.Text)) & "," _
                  & " biw10 = " & Val(Trim(txtbiw10.Text)) & "," _
                  & " mon15 = " & Val(Trim(txtmon15.Text)) & "," _
                  & " biw15 = " & Val(Trim(txtbiw15.Text)) & "," _
                  & " Monoth = " & Val(Trim(txtmonoth.Text)) & ", " _
                  & " Biwoth = " & Val(Trim(txtbiwoth.Text)) & " " _
                  & " Where Sub_media ='" & i & "'"
                 ' & " Where Agentname ='" & Trim(txtsubmedia.Text) & "'"
                                                
                                                     
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        MsgBox "Record Modified With " & Chr(10) & Chr(10) & _
               "Sub Media = " & i, vbInformation, "Data Modified"
        textclear
        Populatesubmedia
        ModifyData = True
        Exit Function
End Function
Private Function DeleteData() As Boolean
  Dim i
    
    DeleteData = False
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
     i = findfirstfixup(Trim(lstsubmedia.Text))
        
       Sqlqry = "Delete * from Cinema_rates Where Sub_media = '" & i & "'"
                                              
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        MsgBox "Record Deleted With " & Chr(10) & Chr(10) & _
               "Sub Media = " & i, vbInformation, "Data Modified"
        textclear
        Populatesubmedia
   
               
End Function

Private Sub Populatesubmedia()
    textclear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Media Where Media_type='Cinema' Order by Sub_media "
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
      rs.MoveFirst
        lstsubmedia.Clear
        Do Until rs.EOF
            lstsubmedia.AddItem rs!Sub_media
            rs.MoveNext
        Loop
    End If
        
End Sub
Private Sub Form_Load()
    Populatesubmedia
    Populateregion
    textclear
End Sub
Private Sub Populateregion()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select distinct(region) from cinema_rates Order by region "
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
      rs.MoveFirst
        CboRegion.Clear
        Do Until rs.EOF
            CboRegion.AddItem rs!region
            rs.MoveNext
        Loop
    End If
        
End Sub
Private Sub lstSubMedia_Click()

Dim i
Dim tempBln As String
    If lstsubmedia.ListIndex = -1 Then
        tempBln = False
    Else
        tempBln = True
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = findfirstfixup(Trim(lstsubmedia.Text))
        Sqlqry = "Select * from Cinema_rates Where Sub_media= '" & i & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            textclear
            txtsubMedia.Text = Trim(lstsubmedia.Text)
            txtseats.SetFocus
            Exit Sub
         Else
           txtsubMedia = rs!Sub_media
           CboRegion = rs!region
           
           If IsNull(rs!seats) = True Then
              txtseats = ""
           Else
              txtseats = rs!seats
           End If
                               
           If IsNull(rs!Shows) = True Then
              txtshows = ""
           Else
              txtshows = rs!Shows
           End If
           
           If IsNull(rs!Type) = True Then
              txttype = ""
           Else
              txttype = rs!Type
           End If
                     
           If IsNull(rs!mon30) = True Then
              txtmon30 = ""
           Else
              txtmon30 = rs!mon30
           End If
          
           If IsNull(rs!biw30) = True Then
              txtbiw30 = ""
           Else
              txtbiw30 = rs!biw30
           End If
           
           If IsNull(rs!mon60) = True Then
              txtmon60 = ""
           Else
              txtmon60 = rs!mon60
           End If
          
          If IsNull(rs!biw60) = True Then
              txtbiw60 = ""
           Else
              txtbiw60 = rs!biw60
           End If
           
           If IsNull(rs!mon90) = True Then
              txtmon90 = ""
           Else
              txtmon90 = rs!mon90
           End If
          
          If IsNull(rs!biw90) = True Then
              txtbiw90 = ""
           Else
              txtbiw90 = rs!biw90
           End If
           
          
          If IsNull(rs!mon10) = True Then
              txtmon10 = ""
           Else
              txtmon10 = rs!mon10
           End If
          
          If IsNull(rs!biw10) = True Then
              txtbiw10 = ""
           Else
              txtbiw10 = rs!biw10
           End If
                     
          If IsNull(rs!mon15) = True Then
              txtmon15 = ""
          Else
              txtmon15 = rs!mon15
          End If
          
          If IsNull(rs!biw15) = True Then
              txtbiw15 = ""
          Else
              txtbiw15 = rs!biw15
          End If
           
          If IsNull(rs!monoth) = True Then
              txtmonoth = ""
           Else
              txtmonoth = rs!monoth
           End If
          
          If IsNull(rs!biwoth) = True Then
              txtbiwoth = ""
           Else
              txtbiwoth = rs!biwoth
           End If
           
          
       End If
    
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
Private Sub LstSubmedia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtsubMedia.SetFocus
End Sub
Private Sub txtbiw10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmon10.SetFocus
End Sub
Private Sub txtbiw15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmon15.SetFocus
End Sub
Private Sub txtbiw30_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmon30.SetFocus
End Sub
Private Sub txtbiw60_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmon60.SetFocus
End Sub
Private Sub txtbiw90_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmon90.SetFocus
End Sub
Private Sub txtbiwoth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmonoth.SetFocus
End Sub
Private Sub txtmon10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtbiw15.SetFocus
End Sub
Private Sub txtmon15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtbiwoth.SetFocus
End Sub
Private Sub txtmon30_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtbiw60.SetFocus
End Sub
Private Sub txtmon60_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtbiw90.SetFocus
End Sub
Private Sub txtmon90_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtbiw10.SetFocus
End Sub
Private Sub txtmonoth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub
Private Sub txtseats_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtshows.SetFocus
End Sub
Private Sub txtshows_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txttype.SetFocus
End Sub
Private Sub txtsubmedia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboRegion.SetFocus
End Sub
Private Sub txttype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtbiw30.SetFocus
End Sub
