VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmagencyBudget 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8565
   ClientLeft      =   15
   ClientTop       =   255
   ClientWidth     =   11850
   LinkTopic       =   "form1"
   ScaleHeight     =   8565
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   8295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.Frame fraAccount 
         BackColor       =   &H00FFFFFF&
         Height          =   6975
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   11415
         Begin VB.TextBox txtBudget 
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
            ForeColor       =   &H80000007&
            Height          =   375
            Left            =   8280
            TabIndex        =   18
            Top             =   5760
            Width           =   1335
         End
         Begin VB.ComboBox CboCurrency 
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
            Height          =   360
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   5760
            Width           =   1455
         End
         Begin VB.ComboBox CboMedia 
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
            Height          =   360
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   5760
            Width           =   2775
         End
         Begin VB.ComboBox CboAgency 
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
            Height          =   360
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   4920
            Width           =   4335
         End
         Begin VB.ComboBox CboYear 
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
            Height          =   360
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   4920
            Width           =   1215
         End
         Begin VB.ListBox lstAgency 
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
            Height          =   4335
            Left            =   480
            TabIndex        =   8
            Top             =   240
            Width           =   9975
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
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
            Height          =   240
            Left            =   4320
            TabIndex        =   13
            Top             =   5880
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Year"
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
            Left            =   720
            TabIndex        =   12
            Top             =   5040
            Width           =   510
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   4440
            TabIndex        =   11
            Top             =   5040
            Width           =   795
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   120
            TabIndex        =   10
            Top             =   5880
            Width           =   1125
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Budget"
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
            Left            =   7320
            TabIndex        =   9
            Top             =   5880
            Width           =   870
         End
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00E0E0E0&
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
         Picture         =   "frmagencybudget.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7440
         Width           =   975
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00E0E0E0&
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
         Picture         =   "frmagencybudget.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7440
         Width           =   975
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00E0E0E0&
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
         Picture         =   "frmagencybudget.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7440
         Width           =   975
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
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
         Picture         =   "frmagencybudget.frx":0986
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7440
         Width           =   975
      End
      Begin VB.CommandButton cmdMod 
         BackColor       =   &H00E0E0E0&
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
         Picture         =   "frmagencybudget.frx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7440
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Add"
         DisabledPicture =   "frmagencybudget.frx":120A
         DownPicture     =   "frmagencybudget.frx":164C
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
         Picture         =   "frmagencybudget.frx":1A8E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7440
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.Label lblserialno2 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   5880
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Serial No. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   7440
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Lblserialno 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   8880
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmagencyBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idref As Integer
Dim X
Private Sub AutoIncrementVoucher()
X = 0
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select tserial from agmediabudget order by tserial"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
If rs.RecordCount <> 0 Then
   rs.MoveLast
   X = Val(rs!tserial)
   lblserialno.Caption = X + 1
Else
   lblserialno = 100001
End If
End Sub
Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtBudget.SetFocus
End Sub

Private Sub cmdadd_Click()
    Dim ws As Workspace
    Dim db As Database
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim Sqlqry As String
    Dim Sqlqry1 As String
    
 
  If ValidateData = True Then
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    ' If Mid(CboMedia.Text, 1, 8) = "Cinema :" Then
    '   Sqlqry1 = " select * from agmediabudget where agency='" & findfirstfixup(Trim(CboAgency)) & "' and tYear='" & Trim(cboyear) & "' and subMedia='" & Trim(Mid(CboMedia, 10, 50)) & "' "
    '     Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    '      If rs.RecordCount <> 0 Then
    '         MsgBox " Budget Already set for the selected Agency and Media"
    '         Exit Sub
    '      Else
               
     '      Sqlqry = " Insert into agmediabudget values('" & Val(lblserialno.Caption) & "','" & findfirstfixup(CboAgency) & "','" & _
                   Trim(cboyear) & "','" & _
                   Trim(Mid(CboMedia, 10, 50)) & "','" & _
                   Trim(cboCurrency) & "'," & _
                   Trim(txtBudget) & ")"
                   
     '      ws.BeginTrans
      '     db.Execute (Sqlqry)
      '     ws.CommitTrans
      '   End If
        
     ' Else
        Sqlqry1 = " select * from agmediabudget where agency='" & findfirstfixup(Trim(CboAgency)) & "' and tYear='" & Trim(Cboyear) & "' and subMedia='" & Trim(CboMedia) & "' "
            Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
            If rs.RecordCount <> 0 Then
               MsgBox " Budget Already set for the selected Agency and Media"
               Exit Sub
            Else
                 
             Sqlqry = " Insert into agmediabudget values('" & Val(lblserialno.Caption) & "','" & findfirstfixup(CboAgency) & "','" & _
                     Trim(Cboyear) & "','" & _
                     Trim(CboMedia) & "','" & _
                     Trim(CboCurrency) & "'," & _
                     Trim(txtBudget) & ")"
                     
             ws.BeginTrans
             db.Execute (Sqlqry)
             ws.CommitTrans
            End If

       
    '  End If
    
    PopulateAccodes
    
        MsgBox "Record is inserted", vbDefaultButton3, "Status"
        textclear
        lblserialno.Caption = lblserialno.Caption + 1
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
 Cboyear.ListIndex = -1
 CboAgency.ListIndex = -1
 CboMedia.ListIndex = -1
 txtBudget.Text = 0
 lblserialno2.Caption = ""
End Function
Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub
Private Function ValidateData()

ValidateData = False
If Cboyear.Text = "" Then
   MsgBox "Invalid year", vbInformation, "Invalid Entry"
   Cboyear.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf CboAgency.Text = "" Then
   MsgBox "Invalid Agency", vbInformation, "Invalid Entry"
   CboAgency.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf CboMedia.Text = "" Then
   MsgBox "Invalid Media", vbInformation, "Invalid Entry"
   CboMedia.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
  ValidateData = True
End If

End Function

Private Sub cmdDelete_Click()
 If lstAgency.SelCount = 0 Then
        MsgBox "Select the Agency in a list for Deletion.", vbInformation, "Selection Error"
        lstAgency.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
          
         i = Trim(Cboyear.Text)
           
        tempStr = MsgBox("Do You Want To Delete", vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If DeleteData = False Then Exit Sub
        Else
            MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
            Cboyear.SetFocus
            Exit Sub
        End If
End Sub

Private Sub cmdMod_Click()
  Z = 1
    If lstAgency.SelCount = 0 Then
        MsgBox "Select the Agency,Media In a list box for Modification.", vbInformation, "Selection Error"
        lstAgency.SetFocus
        Exit Sub
    End If
    
    If lblserialno2.Caption = "" Then
      MsgBox " Cannot modify, entry already existng"
      lstAgency.SetFocus
      Exit Sub
    End If
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   If Mid(CboMedia, 1, 8) = "Cinema :" Then
    Sqlqry1 = " select * from agmediabudget where agency='" & findfirstfixup(Trim(CboAgency)) & "' and tYear='" & Trim(Cboyear) & "' and subMedia='" & Trim(Mid(CboMedia, 10, 50)) & "' and budget=" & Val(txtBudget) & " "
   Else
    Sqlqry1 = " select * from agmediabudget where agency='" & findfirstfixup(Trim(CboAgency)) & "' and tYear='" & Trim(Cboyear) & "' and subMedia='" & Trim(CboMedia) & "' and budget=" & Val(txtBudget) & " "
   End If
   
   Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If rs.RecordCount <> 0 Then
      MsgBox " Budget Already set for the selected Agency and Media"
      Exit Sub
   Else
     
      If ValidateData = False Then Exit Sub
           
      tempStr = MsgBox("Do You Want To Modify the Entry :" & Mid(lstAgency, 10, 45), vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If ModifyData = False Then Exit Sub
        Else
            MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
            Cboyear.SetFocus
            Exit Sub
        End If
  End If
 End Sub

Private Function ModifyData() As Boolean
    Dim i
    ModifyData = False
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    
      i = lblserialno2.Caption
    
         If Mid(CboMedia, 1, 8) = "Cinema :" Then
              Sqlqry = "Update AgMediaBudget set Agency='" & CboAgency.Text & "'," _
                   & " Tyear='" & Cboyear.Text & "'," _
                   & " SubMedia='" & Trim(Mid(CboMedia.Text, 10, 50)) & "'," _
                   & " TCurrency ='" & Trim(CboCurrency.Text) & "'," _
                   & " Budget=" & Val(txtBudget.Text) & " where TSerial='" & i & "'"
                   
              ws.BeginTrans
              db.Execute (Sqlqry)
              ws.CommitTrans
            
         Else
            Sqlqry = "Update AgMediaBudget set Agency='" & CboAgency.Text & "'," _
                    & " Tyear='" & Cboyear.Text & "'," _
                    & " SubMedia='" & CboMedia.Text & "'," _
                    & " TCurrency ='" & Trim(CboCurrency.Text) & "'," _
                    & " Budget=" & Val(txtBudget.Text) & " where TSerial='" & i & "'"
                    
               ws.BeginTrans
               db.Execute (Sqlqry)
               ws.CommitTrans
         End If
            
        MsgBox "Record Modified "
                
        textclear
        PopulateAccodes
        AutoIncrementVoucher
        tempBln = False
        ModifyData = True
        Exit Function
End Function

Private Function DeleteData() As Boolean
 Dim i
    
    DeleteData = False
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           
    i = Trim(Cboyear.Text)
               
           
           
           Sqlqry = "Delete * from agmediabudget Where tserial = '" & Val(lblserailno2.Caption) & "'"
                                           
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
            
            MsgBox "Record Deleted " & Chr(10) & Chr(10) & _
                vbInformation, "Data Modified"
            textclear
            PopulateAccodes
            tempBln = False
            
        
          
End Function

Private Sub PopulateAccodes()

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from agmediabudget order by agency"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
     rs.MoveFirst
        lstAgency.Clear
        Do Until rs.EOF
            lstAgency.AddItem rs!tserial & "   -  " & rs!Agency & "    -    " & rs!tYear & "   -  " & rs!submedia
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
Dim i

   tempBln = False
    PopulateAccodes
    textclear
    
    CboCurrency.AddItem "DHS"
    CboCurrency.AddItem "USD"
 
i = 2000

For i = 2000 To 2100
 Cboyear.AddItem i
Next
X = 0

Cboyear.Text = Year(Now())
 
populateMedia
PopulateAgencycodes
AutoIncrementVoucher

lblserialno2.Caption = ""

End Sub

Private Sub PopulateAgencycodes()
    CboAgency.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from agndtls Order by AgentName"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
      rs.MoveFirst
        CboAgency.Clear
        
        Do Until rs.EOF
            CboAgency.AddItem rs!agentname
            rs.MoveNext
        Loop
    End If
        
End Sub

Private Sub populateMedia()
    CboMedia.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Media Order by Sub_Media"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        CboMedia.AddItem "Cinema"
       ' CboMedia.AddItem "Online"
       
       rs.MoveFirst
            Do Until rs.EOF
               If rs!Media_Type <> "Cinema" Then
                
                CboMedia.AddItem Trim(rs!Media_Type) & " : " & Trim(rs!sub_media)
               Else
                CboMedia.AddItem Trim(rs!sub_media)
               End If
              
              rs.MoveNext
       Loop
    End If
 End Sub

Private Sub lstagency_Click()
Dim i

 
    lblserialno2.Caption = ""
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Val(Mid(lstAgency.Text, 1, 6))
        Sqlqry = "Select * from agmediabudget Where tserial= '" & i & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MsgBox "Particular Record was Deleted.", vbInformation, "Deleted Status"
            Exit Sub
        End If
            If rs!submedia = "Cinema" Then
               CboMedia = "Cinema"
            ElseIf rs!submedia = "Magazine" Then
               CboMedia = "Magazine"
            ElseIf rs!submedia = "Television" Then
               CboMedia = "Television"
            ElseIf rs!submedia = "Online" Then
               CboMedia = "Online"
            Else
             Sqlqry1 = "Select * from Media where sub_media='" & rs!submedia & "' "
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
             If rs1.RecordCount = 0 Then
                MsgBox "sub media not found.", vbInformation, "Deletion"
                Exit Sub
             Else
                If rs1!Media_Type = "Cinema" Then
                  CboMedia = Trim(rs1!Media_Type) & " : " & Trim(rs1!sub_media)
                Else
                  CboMedia = rs1!sub_media
                End If
                  
             End If
                
             CboMedia = rs!submedia
           End If
         
       Cboyear.Text = rs!tYear
       CboAgency = rs!Agency
       CboCurrency = rs!tcurrency
       txtBudget = rs!budget
       lblserialno2.Caption = Val(rs!tserial)
    
End Sub

Private Sub cboyear_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboAgency.SetFocus
End Sub

Private Sub CboAgency_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboMedia.SetFocus
End Sub

Private Sub txtclbal_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub

Private Sub cbomedia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboCurrency.SetFocus
End Sub
