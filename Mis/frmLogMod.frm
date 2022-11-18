VERSION 5.00
Begin VB.Form frmLogMod 
   BackColor       =   &H80000018&
   Caption         =   "Password for Modification"
   ClientHeight    =   4110
   ClientLeft      =   3075
   ClientTop       =   4380
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   4860
   Begin VB.Frame FRAMEPASSWORD 
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
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Administrator"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "User Name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1035
      End
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
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmdCan 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "&Change"
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtConfirmPwd 
         BackColor       =   &H80000018&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtNewPwd 
         BackColor       =   &H80000018&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtOldPwd 
         BackColor       =   &H80000018&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   480
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
         TabIndex        =   11
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
         TabIndex        =   9
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
         TabIndex        =   4
         Top             =   600
         Width           =   1170
      End
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
Attribute VB_Name = "frmLogMod"
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
FRAMEPASSWORD.Visible = True

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
    Sqlqry = "Select * from PWDLOGMod"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        MsgBox "Password Not Found. Enter New Password.", vbInformation, "Password Status"
        Exit Sub
    Else
        rs.MoveFirst
        logString = rs!pwd
    End If
    If StrComp(Trim(txtOldPwd), logString, vbTextCompare) = 0 Then
        If StrComp(Trim(txtNewPwd), Trim(txtConfirmPwd), vbTextCompare) = 0 Then
            Sqlqry = "Update PWDLOGMod Set PWD='" & Trim(txtConfirmPwd) & "'"
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
    Sqlqry = "Select * from PWDLOGMod"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Sqlqry = "Insert into PWDLOGMod Values('" & Trim(txtPassword) & "')"
        ws.BeginTrans
        db.Execute Sqlqry
        ws.CommitTrans
    Else
        rs.MoveFirst
        logString = rs!pwd
    End If
         
    If StrComp(logString, Trim(txtPassword), vbTextCompare) = 0 Then
        Unload Me
        MDIMIS.mnurep.Enabled = True
        Z = 1
        PopulateModification
        
        'PopupMenu MDIMIS.mnurep
        'If Button = vbLeftButton Then PopupMenu mnurep
    Else
        MsgBox "Invalid Password. Enter Correct Password.", vbInformation, "Invalid Password"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    FRAMEPASSWORD.Visible = True
    Frame1.Visible = False
    Z = 0
End Sub

Private Sub mnuChange_Click()
txtNewPwd.Text = ""
txtOldPwd.Text = ""
txtConfirmPwd.Text = ""
FRAMEPASSWORD.Visible = False
Frame1.Visible = True
Frame1.ZOrder 0
End Sub

Private Sub mnuExit_Click()
cmdCancel.Value = True
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOK.SetFocus
End Sub

Private Sub PopulateModification()



End Sub

