VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmCrntPrint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Credit Note Printing"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ChkMis 
      BackColor       =   &H80000009&
      Caption         =   "Mis"
      Height          =   195
      Left            =   9000
      TabIndex        =   5
      Top             =   240
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Credit Note - Printing"
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
      Height          =   7215
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   11655
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   8520
         Top             =   6240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin VB.ListBox LstCrnt 
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
         Height          =   4860
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   11175
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
         Left            =   6120
         OLEDropMode     =   1  'Manual
         Picture         =   "frmCrntPrint.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6000
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
         Left            =   4680
         Picture         =   "frmCrntPrint.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6000
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
         Left            =   3360
         Picture         =   "frmCrntPrint.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   11640
         X2              =   0
         Y1              =   5640
         Y2              =   5640
      End
   End
End
Attribute VB_Name = "frmCrntPrint"
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
Dim MTYPE As String

Private Sub cmdBack_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub cmdClear_Click()
 LstCrnt.ListIndex = 0
End Sub
Private Sub CmdPrint_Click()
 If LstCrnt.SelCount = 0 Then
  MsgBox "Select Credit Note voucher to print from the list box "
  LstCrnt.SetFocus
  Exit Sub
 Else
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = " Select * from Crdt_mas where Vouc_No=" & Val(Mid(LstCrnt, 1, 7)) & ""
   Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
   If rs.RecordCount <> 0 Then
        ttlamount = 0
        rs.MoveFirst
        ctype = rs!tcurrency
        ttlamount = rs!tra_amount
   End If
   
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry1 = " Select * from BO_mas where Serial_No='" & Mid(rs!ref_no, 1, 7) & "'"
   Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
   If rs1.RecordCount <> 0 Then
        rs1.MoveFirst
        MTYPE = rs1!SUB_MEDIA
        
   End If
   
   MTYPE = Trim(Mid(MTYPE, 1, 11))
 '  If MTYPE = "ZEINA" Or MTYPE = "ALAM ASSAYA" Then
  If ChkMis.Value = Unchecked Then
          
        If ctype = "DHS" Then
        
             CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
             CrystalReport1.ReportFileName = App.Path & "\crntvouMPS.rpt"
             CrystalReport1.SelectionFormula = "{crdt_mas.Vouc_no}=" & Mid(LstCrnt, 1, 7) & ""
             CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(ttlamount)) & " Only" & "'"
             CrystalReport1.Formulas(1) = "curtype='" & ctype & "'"
             CrystalReport1.WindowState = crptMaximized
             CrystalReport1.Action = 1
        Else
             CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
             CrystalReport1.ReportFileName = App.Path & "\crntvouMPS.rpt"
             CrystalReport1.SelectionFormula = "{crdt_mas.Vouc_no}=" & Mid(LstCrnt, 1, 7) & ""
             CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(ttlamount)) & " Only" & "'"
             CrystalReport1.Formulas(1) = "curtype='" & ctype & "'"
             CrystalReport1.WindowState = crptMaximized
             CrystalReport1.Action = 1
        End If
    Else
         If ctype = "DHS" Then
        
             CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
             CrystalReport1.ReportFileName = App.Path & "\crntvou.rpt"
             CrystalReport1.SelectionFormula = "{crdt_mas.Vouc_no}=" & Mid(LstCrnt, 1, 7) & ""
             CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(ttlamount)) & " Only" & "'"
             CrystalReport1.Formulas(1) = "curtype='" & ctype & "'"
             CrystalReport1.WindowState = crptMaximized
             CrystalReport1.Action = 1
        Else
             CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
             CrystalReport1.ReportFileName = App.Path & "\crntvou.rpt"
             CrystalReport1.SelectionFormula = "{crdt_mas.Vouc_no}=" & Mid(LstCrnt, 1, 7) & ""
             CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(ttlamount)) & " Only" & "'"
             CrystalReport1.Formulas(1) = "curtype='" & ctype & "'"
             CrystalReport1.WindowState = crptMaximized
             CrystalReport1.Action = 1
        End If
    End If
 End If
End Sub

Private Sub Form_Load()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = " Select * from crdt_mas order by ref_no"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

LstCrnt.Clear

If rs.RecordCount = 0 Then
    MsgBox "No Records found in the Credit Note Register"
Else
   rs.MoveFirst
   Do Until rs.EOF
   ' rs!SUPP_NAME
      LstCrnt.AddItem rs!vouc_no & "   :   " & rs!ref_no & "    :   " & rs!supp_name & "   :  " & rs!tcurrency & "  :  " & rs!tra_amount
      rs.MoveNext
   Loop
End If

End Sub


