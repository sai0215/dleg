VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmDBNTPrint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Debit Note Printing"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Debit Note - Printing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5055
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   7215
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   6480
         Top             =   4320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin VB.ListBox LstDBNT 
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
         Height          =   2940
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   6735
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
         Height          =   975
         Left            =   4080
         OLEDropMode     =   1  'Manual
         Picture         =   "frmDBNTPrint.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3840
         Width           =   1455
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
         Height          =   975
         Left            =   2640
         Picture         =   "frmDBNTPrint.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3840
         Width           =   1455
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
         Height          =   975
         Left            =   1320
         Picture         =   "frmDBNTPrint.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   7200
         X2              =   0
         Y1              =   3600
         Y2              =   3600
      End
   End
End
Attribute VB_Name = "frmDBNTPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim rs As Recordset
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim rs1 As Recordset
Dim ttlamount As Currency
Dim ctype As String
Private Sub cmdBack_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub cmdClear_Click()
LstDBNT.ListIndex = 0
End Sub
Private Sub CmdPrint_Click()

 If LstDBNT.SelCount = 0 Then
  MsgBox "Select Debit Note voucher to print from the list box "
  LstDBNT.SetFocus
  Exit Sub
 Else
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = " Select * from DEBT_MAS where Vouc_No=" & Val(LstDBNT) & ""
   Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
   If rs.RecordCount <> 0 Then
        ttlamount = 0
        rs.MoveFirst
        ctype = rs!tcurrency
        ttlamount = rs!tra_amount
   End If
  If ctype = "DHS" Then
   CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
   CrystalReport1.ReportFileName = App.Path & "\dbntvou.rpt"
   CrystalReport1.SelectionFormula = "{DEBT_MAS.Vouc_no}=" & LstDBNT & ""
   CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(ttlamount)) & " Only" & "'"
   CrystalReport1.Formulas(1) = "curtype='" & ctype & "'"
   CrystalReport1.WindowState = crptMaximized
   CrystalReport1.Action = 1
  Else
   CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
   CrystalReport1.ReportFileName = App.Path & "\dbntvou.rpt"
   CrystalReport1.SelectionFormula = "{DEBT_MAS.Vouc_no}=" & LstDBNT & ""
   CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(ttlamount)) & " Only" & "'"
   CrystalReport1.Formulas(1) = "curtype='" & ctype & "'"
   CrystalReport1.WindowState = crptMaximized
   CrystalReport1.Action = 1
 End If
 End If
End Sub

Private Sub Form_Load()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = " Select vouc_no from DEBT_MAS order by Vouc_No"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

LstDBNT.Clear

If rs.RecordCount = 0 Then
    MsgBox "No Records found in the Debit Note Register"
Else
   rs.MoveFirst
   Do Until rs.EOF
      LstDBNT.AddItem rs!vouc_no
      rs.MoveNext
   Loop
End If

End Sub

