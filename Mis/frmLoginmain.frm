VERSION 5.00
Begin VB.Form frmLoginMAIN 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Login To Financial Information"
   ClientHeight    =   6525
   ClientLeft      =   30
   ClientTop       =   645
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9570
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtUserName 
      Height          =   330
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Administrator"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtOldPwd 
         BackColor       =   &H80000018&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmdCan 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "&Change"
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtConfirmPwd 
         BackColor       =   &H80000018&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtNewPwd 
         BackColor       =   &H80000018&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Confirm Password"
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
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "New Password"
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
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Old Password"
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
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1170
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name "
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   840
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuChange 
         Caption         =   "Change Password"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmLoginMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim logString As String
Dim ws As Workspace
Dim db As Database
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim rs  As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset

Private Sub cmdCan_Click()
Me.Height = 2500
txtPassword.Text = ""
txtPassword.SetFocus
SendKeys "{Home}+{End}"
Frame1.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdChange_Click()
If ValidateData = False Then Exit Sub
ChangePassword
End Sub

Private Function ValidateData() As Boolean
ValidateData = False
    If Trim(txtOldPwd) = "" Then
        MsgBox "Enter the Old Password.", vbInformation, "Invalid Entry"
        txtOldPwd.SetFocus
        Exit Function
    ElseIf Trim(txtNewPwd) = "" Then
        MsgBox "Enter the New Password.", vbInformation, "Invalid Entry"
        txtNewPwd.SetFocus
        Exit Function
    ElseIf Trim(txtConfirmPwd) = "" Then
        MsgBox "Enter the Confirm Password.", vbInformation, "Invalid Entry"
        txtConfirmPwd.SetFocus
        Exit Function
    Else
        ValidateData = True
    End If
End Function

Private Sub ChangePassword()
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from PWDLog"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        MsgBox "Password Not Found. Enter New Password.", vbInformation, "Password Status"
        Exit Sub
    Else
        rs.MoveFirst
        logString = rs!logString
    End If
    If StrComp(Trim(txtOldPwd), logString, vbTextCompare) = 0 Then
        If StrComp(Trim(txtNewPwd), Trim(txtConfirmPwd), vbTextCompare) = 0 Then
            Sqlqry = "Update PWDLog Set LogString='" & Trim(txtConfirmPwd) & "'"
            ws.BeginTrans
            db.Execute Sqlqry
            ws.CommitTrans
            MsgBox "Password Successfully Changed.", vbInformation, "Change Status"
            cmdCan.Value = True
        Else
            MsgBox "Confirm Password Should be same as New Password", vbInformation, "Invalid Entry"
            txtNewPwd.SetFocus
            SendKeys "{Home}+{End}"
            Exit Sub
        End If
    Else
        MsgBox "Invalid Old Password. Enter Correct Password.", vbInformation, "Invalid Entry"
        txtOldPwd.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
End Sub

Private Sub cmdOK_Click()

  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from PWDLog"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Sqlqry = "Insert into PWDLog Values('" & Trim(txtPassword) & "')"
        ws.BeginTrans
        db.Execute Sqlqry
        ws.CommitTrans
    Else
        rs.MoveFirst
        logString = rs!PWD
    End If
         
    If StrComp(logString, Trim(txtPassword), vbTextCompare) = 0 Then
        Unload Me
        MDIMIS.Show
    Else
        MsgBox "Invalid Password. Enter Correct Password.", vbInformation, "Invalid Password"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
txtUserName.SetFocus
SendKeys "{Home}+{End}"
End Sub

Private Sub Form_Load()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Me.Height = 2500
End Sub

Private Sub mnuChange_Click()
Me.Height = 4100
txtNewPwd.Text = ""
txtOldPwd.Text = ""
txtConfirmPwd.Text = ""
Frame1.Visible = True
Frame1.ZOrder 0
End Sub

Private Sub mnuExit_Click()
cmdCancel.Value = True
End Sub

