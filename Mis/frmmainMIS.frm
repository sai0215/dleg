VERSION 5.00
Begin VB.Form frmmainMIS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0FF&
   Caption         =   "frmMainMIs"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11835
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0FF&
      Height          =   6495
      Left            =   0
      ScaleHeight     =   6435
      ScaleWidth      =   11835
      TabIndex        =   5
      Top             =   2040
      Width           =   11895
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   9600
      Picture         =   "frmmainMIS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton cmdFA 
      BackColor       =   &H00C0C0FF&
      Caption         =   "FA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4920
      Picture         =   "frmmainMIS.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton cmdRep 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7320
      Picture         =   "frmmainMIS.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton cmdBo 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Booking Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2520
      Picture         =   "frmmainMIS.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton cmdACMP 
         BackColor       =   &H00C0C0FF&
         Caption         =   "General Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   0
         Picture         =   "frmmainMIS.frx":82C7
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Menu mnugen 
      Caption         =   "General Data"
      Begin VB.Menu mnuAgency 
         Caption         =   "&Agency"
      End
      Begin VB.Menu mnuClient 
         Caption         =   "&Cleint"
      End
      Begin VB.Menu mnuMediaType 
         Caption         =   "&Media Type"
      End
      Begin VB.Menu mnuproduct 
         Caption         =   "&Product"
      End
   End
   Begin VB.Menu mnubo 
      Caption         =   "Bookiing &Order"
      Begin VB.Menu mnubonew 
         Caption         =   "&New Entry"
      End
      Begin VB.Menu mnubomod 
         Caption         =   "&Modify"
      End
      Begin VB.Menu mnuboDel 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuboback 
         Caption         =   "&Back"
      End
   End
   Begin VB.Menu mnufa 
      Caption         =   "F&A"
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
   End
   Begin VB.Menu mnuexit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmmainMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdACMP_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbLeftButton Then PopupMenu mnugen
End Sub

Private Sub cmdBo_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbLeftButton Then PopupMenu mnubo
End Sub

'Private Sub MDIForm_Load()
' WebBrowser1.Navigate App.Path & "\UDAYMA.HTML"
'End Sub
'Private Sub MDIForm_Unload(Cancel As Integer)
'If Cancel = 1 Then End
'End Sub

Private Sub mnuagency_Click()
Picture2.Visible = False
frmAgency.Show
End Sub

