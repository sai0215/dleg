VERSION 5.00
Begin VB.Form frmPosting 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Freezing The Transactions"
   ClientHeight    =   8775
   ClientLeft      =   -90
   ClientTop       =   345
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Posting Entry Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   2280
      TabIndex        =   4
      Top             =   1080
      Width           =   6735
      Begin VB.CommandButton cmdWork 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Post"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<< &Back"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Month-Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   1335
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Width           =   5415
         Begin VB.ComboBox cboYear 
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
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   600
            Width           =   1455
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
            ForeColor       =   &H00000080&
            Height          =   360
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Year"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   3840
            TabIndex        =   7
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Month"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   1320
            TabIndex        =   6
            Top             =   360
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "frmPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim Sqlqry3 As String
Dim Sqlqry4 As String
Dim Sqlqry5 As String
Dim Sqlqry6 As String
Dim Sqlqry7 As String
Dim Sqlqry8 As String
Dim Sqlqry9 As String
Dim Sqlqry10 As String
Dim SQLQRY11 As String
Dim SQLQRY12 As String
Dim SNo As Long
Dim SelMonth As Integer
Dim FirstDate As Date
Dim LastDate As Date
Private Sub cmdBack_Click()
 Unload Me
End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub
Private Sub cmdWork_Click()
Dim X
If ValidateData = False Then Exit Sub

X = MsgBox("Do You Want to Post the transactions", vbInformation + vbYesNo, "Warning")
If X = vbNo Then Exit Sub

Dim i As Integer
Dim j As Integer
    j = cbomonth.ListIndex + 1
    FirstDate = Now
    LastDate = Now
    i = DaysinMonth(j, cboyear.Text)
    FirstDate = DateValue("1-" & j & "-" & cboyear)
    LastDate = DateValue(i & "-" & j & "-" & cboyear)
    
    Sqlqry = "Update bpmt_mas set Status='Y' Where tDate Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "#"
    Sqlqry1 = "Update brpt_mas set Status='Y' Where tDate Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "#"
    Sqlqry2 = "Update capr_mas set Status='Y' Where tDate Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "#"
  '  Sqlqry3 = "Update casl_mas set Status='Y' Where tDate Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "#"
    Sqlqry4 = "Update cpmt_mas set Status='Y' Where tDate Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "#"
    Sqlqry5 = "Update crdt_mas set Status='Y' Where tDate Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "#"
    Sqlqry6 = "Update crpr_mas set Status='Y' Where tDate Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "#"
    Sqlqry7 = "Update crpt_mas set Status='Y' Where tDate Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "#"
    Sqlqry8 = "Update bo_mas set Status='Y' Where invoice_Date Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "#"
    Sqlqry9 = "Update debt_mas set Status='Y' Where tDate Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "#"
    Sqlqry10 = "Update jrnl_tra set Status='Y' Where tDate Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "#"
          
    ws.BeginTrans
    db.Execute Sqlqry
    db.Execute Sqlqry1
    db.Execute Sqlqry2
    db.Execute Sqlqry4
    db.Execute Sqlqry5
    db.Execute Sqlqry6
    db.Execute Sqlqry7
    db.Execute Sqlqry8
    db.Execute Sqlqry9
    db.Execute Sqlqry10
  
    ws.CommitTrans
    MsgBox "Data has been Posted for the month of " & cbomonth.Text & ", Year " & cboyear.Text, vbInformation, "Data Modified"
    cmdBack.Value = True
End Sub

Private Function ValidateData() As Boolean
  ValidateData = False
    If cboyear.Text = "" Then
        MsgBox "Invalid Year. Select the Year", vbInformation, "Invalid Entry"
        cboyear.SetFocus
        SendKeys "{Home}+{End}"
        Exit Function
    Else
         ValidateData = True
    End If
End Function

Private Sub Form_Load()
Dim i As Integer
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
cbomonth.Clear
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
cbomonth.ListIndex = Month(Now) - 1
cboyear.Clear
For i = 2000 To 2200
    cboyear.AddItem i
Next i

cboyear.Text = Year(Now)
PopulateUnFreezed
End Sub

Private Sub PopulateUnFreezed()
 Dim Last As Date
 Dim i As Integer
 Dim j As Integer
 Sqlqry = "Select TDate from cpmt_mas Where Status='Y' order by tdate"
 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
  If rs.RecordCount <> 0 Then
    rs.MoveLast
    Last = Format(rs!tdate, "mm/dd/yyyy")
    i = Month(Last)
    SelMonth = i
    cbomonth.Clear
    Select Case i
    Case 12
        cbomonth.AddItem "January"
    Case 1
        cbomonth.AddItem "February"
    Case 2
        cbomonth.AddItem "March"
    Case 3
        cbomonth.AddItem "April"
    Case 4
        cbomonth.AddItem "May"
    Case 5
        cbomonth.AddItem "June"
    Case 6
        cbomonth.AddItem "July"
    Case 7
        cbomonth.AddItem "August"
    Case 8
        cbomonth.AddItem "September"
    Case 9
        cbomonth.AddItem "October"
    Case 10
        cbomonth.AddItem "November"
    Case 11
        cbomonth.AddItem "December"
    End Select
    j = Year(Last)
    cboyear.Clear
    If i = 12 Then j = j + 1
    cboyear.AddItem j
    cbomonth.ListIndex = 0
    cboyear.ListIndex = 0
  Else
    cbomonth.ListIndex = 0
  End If
End Sub
