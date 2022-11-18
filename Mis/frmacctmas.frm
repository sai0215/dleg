VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmacctmas 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8775
   ClientLeft      =   -105
   ClientTop       =   285
   ClientWidth     =   12060
   LinkTopic       =   "form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   360
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Chart of Accounts"
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
      Height          =   8175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.Frame fraAccount 
         BackColor       =   &H80000005&
         Height          =   6735
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   11415
         Begin VB.TextBox txtClbal 
            BackColor       =   &H00FFFFFF&
            DataField       =   "ACCT_CODE"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6480
            TabIndex        =   12
            Top             =   5910
            Width           =   2175
         End
         Begin VB.TextBox txtOpbal 
            BackColor       =   &H00FFFFFF&
            DataField       =   "ACCT_CODE"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2280
            TabIndex        =   11
            Top             =   5910
            Width           =   2175
         End
         Begin VB.TextBox txtaccode 
            BackColor       =   &H00FFFFFF&
            DataField       =   "ACCT_CODE"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2280
            TabIndex        =   10
            Top             =   5190
            Width           =   2175
         End
         Begin VB.TextBox txtacdesc 
            BackColor       =   &H00FFFFFF&
            DataField       =   "ACCT_NAME"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6480
            ScrollBars      =   1  'Horizontal
            TabIndex        =   9
            Top             =   5190
            Width           =   4695
         End
         Begin VB.ListBox lstAccodes 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   4620
            Left            =   720
            TabIndex        =   8
            Top             =   240
            Width           =   9975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Account Code"
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
            Height          =   240
            Left            =   480
            TabIndex        =   16
            Top             =   5280
            Width           =   1695
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Description"
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
            Height          =   240
            Left            =   4680
            TabIndex        =   15
            Top             =   5280
            Width           =   1680
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   360
            TabIndex        =   14
            Top             =   6000
            Width           =   1800
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Closing Balance"
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
            Height          =   240
            Left            =   4680
            TabIndex        =   13
            Top             =   6000
            Width           =   1710
         End
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5880
         Picture         =   "frmacctmas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4800
         Picture         =   "frmacctmas.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<<&Back<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         Picture         =   "frmacctmas.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6960
         Picture         =   "frmacctmas.frx":0986
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmdMod 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Modify"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3720
         Picture         =   "frmacctmas.frx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Add"
         DisabledPicture =   "frmacctmas.frx":120A
         DownPicture     =   "frmacctmas.frx":164C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2640
         MaskColor       =   &H008080FF&
         Picture         =   "frmacctmas.frx":1A8E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7320
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmacctmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
    Dim ws As Workspace
    Dim db As Database
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim Sqlqry As String
    Dim Sqlqry1 As String
    
 ValidateData
 If ValidateData = True Then
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    
   Sqlqry1 = " select * from acct_mas where acct_code='" & Trim(txtaccode) & "' "
   Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
   If rs.RecordCount <> 0 Then
      MsgBox " Account Code already exists"
      Exit Sub
   Else
        
   Sqlqry = " Insert into acct_mas values('" & txtaccode & "','" & _
            findfirstfixup(Trim(txtacdesc)) & "','" & _
            Trim(txtOpbal) & "','" & _
            Trim(txtClbal) & "')"
            
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    PopulateAccodes
    MsgBox "Record is inserted", vbDefaultButton3, "Status"
    textclear
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
 txtaccode = ""
 txtacdesc = ""
 txtOpbal = ""
 txtClbal = ""
End Function

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub

Private Function ValidateData()

ValidateData = False
If txtaccode.Text = "" Then
   MsgBox "Invalid Account Code", vbInformation, "Invalid Entry"
   txtaccode.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf txtacdesc.Text = "" Then
   MsgBox "Invalid Account Description", vbInformation, "Invalid Entry"
   txtacdesc.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsNumeric(txtOpbal) = False Then
   MsgBox "Invalid Opening Balance", vbInformation, "Invalid Entry"
   txtOpbal.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
  ValidateData = True
End If

End Function

Private Sub cmdDelete_Click()
 If lstAccodes.SelCount = 0 Then
        MsgBox "Select the Account Code for Deletion.", vbInformation, "Selection Error"
        lstAccodes.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
          
           i = Trim(txtaccode.Text)
           
           
        If txtOpbal.Text <> 0 And txtClbal.Text <> 0 Then
           MsgBox "You can not Delete since the transactions are recorded"
           Exit Sub
        End If
        
        tempStr = MsgBox("Do You Want To Delete the Account Code : " & txtaccode, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If DeleteData = False Then Exit Sub
        Else
            MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
            txtaccode.SetFocus
            Exit Sub
        End If
End Sub

Private Sub cmdMod_Click()
  Z = 1
    If lstAccodes.SelCount = 0 Then
        MsgBox "Select the Account Code for Modification.", vbInformation, "Selection Error"
        lstAccodes.SetFocus
        Exit Sub
    End If
    If txtOpbal.Text = "" Then
      txtOpbal = 0
    End If
    
    If txtClbal.Text = "" Then
      txtClbal = 0
    End If
    
    
    If txtOpbal.Text > 0 Or txtClbal > 0 Then
       MsgBox "Transactions are registered, cannot be modified.", vbInformation, "User Information"
       Exit Sub
    End If
        If ValidateData = False Then Exit Sub
           
          i = Val(Trim(txtaccode.Text))
         
           Set ws = DBEngine.Workspaces(0)
           Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           Sqlqry = "Select * from Acct_mas where Acct_code='" & i & "'"
           Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount = 0 Then
            MsgBox " Account Code not found in the account register"
            Exit Sub
           End If
           
           
        tempStr = MsgBox("Do You Want To Modify the Account Code :" & txtaccode, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If ModifyData = False Then Exit Sub
        Else
            MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
            txtaccode.SetFocus
            Exit Sub
        End If
 End Sub

Private Function ModifyData() As Boolean
    Dim i
    ModifyData = False
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
    i = Trim(txtaccode.Text)
        
        
        Sqlqry = "Update acct_mas Set " _
                   & " Acct_name = '" & findfirstfixup(Trim(txtacdesc.Text)) & "'," _
                   & " Open_bal = '" & Val(txtOpbal.Text) & "'," _
                   & " Close_bal = '" & Val(txtClbal.Text) & "' " _
                   & " Where acct_code = '" & i & "'"
                                           
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        MsgBox "Record Modified With " & Chr(10) & Chr(10) & _
               "Account Code = " & i, vbInformation, "Data Modified"
        textclear
        PopulateAccodes
        tempBln = False
        ModifyData = True
        Exit Function
End Function

Private Function DeleteData() As Boolean
 Dim i
    
    DeleteData = False
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           
    i = Trim(txtaccode.Text)
               
       If txtOpbal > 0 Or txtClbal > 0 Then
         MsgBox " Account Cannot be Deleted since the transactions are recorded"
         DeleteData = False
         Exit Function
       Else
           
        Sqlqry1 = "Select * from bank_mas where bank_code='" & i & "'"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
           
           Sqlqry = "Delete * from acct_mas Where acct_code = '" & i & "'"
                                           
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
            MsgBox "Record Deleted With " & Chr(10) & Chr(10) & _
               "Account Code = " & i, vbInformation, "Data Modified"
            textclear
            PopulateAccodes
            tempBln = False
            If Validate1 = False Then Exit Function
            DeleteData = True
            Exit Function
          Else
            MsgBox "You cannot Delete Bank Code", vbInformation, "Invalid Attempt"
            DeleteData = False
            Exit Function
          End If
        End If
          
End Function

Private Sub PopulateAccodes()

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from acct_mas where val(acct_code)<900000 or val(acct_code)>999999 order by Acct_code"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
     rs.MoveFirst
        lstAccodes.Clear
        Do Until rs.EOF
            lstAccodes.AddItem rs!acct_code & "    :    " & rs!acct_name
            rs.MoveNext
        Loop
    End If
        
End Sub

Private Sub CmdPrint_Click()
 CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
 CrystalReport1.ReportFileName = App.Path & "\AccountList.rpt"
 CrystalReport1.WindowState = crptMaximized
 CrystalReport1.Action = 1
End Sub

Private Sub Form_Load()
    tempBln = False
    PopulateAccodes
    textclear
    txtOpbal.Text = 0
End Sub

Private Sub lstaccodes_Click()
Dim i

    If lstAccodes.ListIndex = -1 Then
        tempBln = False
    Else
        tempBln = True
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Val(Mid(lstAccodes.Text, 1, 6))
        Sqlqry = "Select * from acct_mas Where acct_code= '" & i & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MsgBox "Particular Record was Deleted.", vbInformation, "Deleted Status"
            Exit Sub
        End If
           txtaccode = rs!acct_code
           txtacdesc = rs!acct_name
           txtOpbal = rs!open_bal
           txtClbal = rs!Close_bal
      
    txtaccode.SetFocus
    SendKeys "{home}+{end}"
   
End Sub

Private Sub txtaccode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtacdesc.SetFocus
End Sub

Private Sub txtacdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtOpbal.SetFocus
End Sub

Private Sub txtclbal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub

Private Sub txtopbal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtClbal.SetFocus
End Sub
