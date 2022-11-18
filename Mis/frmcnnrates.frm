VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmcnnrates 
   BackColor       =   &H00FFFFFF&
   Caption         =   " CNN Rates"
   ClientHeight    =   8550
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11895
   WindowState     =   2  'Maximized
   Begin VB.Frame Fracnnrates 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CNN - Rates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8415
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   11535
      Begin VB.Frame frmAgency 
         BackColor       =   &H00FFFFFF&
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
         Height          =   6375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   11295
         Begin VB.TextBox txtCode 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1680
            MaxLength       =   15
            ScrollBars      =   1  'Horizontal
            TabIndex        =   1
            Top             =   600
            Width           =   1455
         End
         Begin VB.ComboBox cbospprog 
            BackColor       =   &H00E0E0E0&
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   3000
            Width           =   1935
         End
         Begin VB.ComboBox txtquarter 
            BackColor       =   &H00E0E0E0&
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
            Left            =   1680
            TabIndex        =   5
            Top             =   4680
            Width           =   2295
         End
         Begin VB.TextBox txtRate 
            BackColor       =   &H00E0E0E0&
            DataField       =   "ACCT_NAME"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1680
            MaxLength       =   4
            ScrollBars      =   1  'Horizontal
            TabIndex        =   6
            Top             =   5520
            Width           =   1575
         End
         Begin VB.ListBox lstCNNRates 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   5715
            Left            =   4200
            TabIndex        =   0
            Top             =   480
            Width           =   6975
         End
         Begin VB.ComboBox cbowtype 
            BackColor       =   &H00E0E0E0&
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox txttime 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1680
            ScrollBars      =   1  'Horizontal
            TabIndex        =   4
            Top             =   3840
            Width           =   2295
         End
         Begin VB.ComboBox cboRegion 
            BackColor       =   &H00E0E0E0&
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   1050
            TabIndex        =   23
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Specific Prog."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   300
            Left            =   120
            TabIndex        =   22
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Quarter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   840
            TabIndex        =   19
            Top             =   4800
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Rate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   1080
            TabIndex        =   18
            Top             =   5640
            Width           =   510
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Day Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   300
            Left            =   240
            TabIndex        =   17
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Region"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   300
            Left            =   240
            TabIndex        =   16
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   1080
            TabIndex        =   15
            Top             =   3960
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   960
         TabIndex        =   13
         Top             =   6960
         Width           =   9015
         Begin VB.CommandButton cmdprintall 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pre&view"
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
            Left            =   4440
            Picture         =   "frmcnnrates.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Add"
            DisabledPicture =   "frmcnnrates.frx":0102
            DownPicture     =   "frmcnnrates.frx":0634
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
            Left            =   840
            MaskColor       =   &H008080FF&
            Picture         =   "frmcnnrates.frx":0B66
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton cmdMod 
            BackColor       =   &H00C0C0C0&
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
            Left            =   2040
            Picture         =   "frmcnnrates.frx":0FA8
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H00C0C0C0&
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
            Height          =   780
            Left            =   5640
            Picture         =   "frmcnnrates.frx":13EA
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdBack 
            BackColor       =   &H00C0C0C0&
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
            Left            =   6840
            Picture         =   "frmcnnrates.frx":14EC
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmddelete 
            BackColor       =   &H00C0C0C0&
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
            Left            =   3240
            Picture         =   "frmcnnrates.frx":15EE
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   480
         Top             =   5160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
   End
End
Attribute VB_Name = "frmcnnrates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim ws As Workspace
Dim db As Database
Dim rs As Recordset
Private Sub cbospprog_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txttime.SetFocus
End Sub
Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboRegion.SetFocus
End Sub

Private Sub txtquarter_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtRate.SetFocus
End Sub
Private Sub cmdprintall_Click()
    CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
    CrystalReport1.ReportFileName = App.Path & "\CNNratesList.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
Private Sub cmdadd_Click()
  If ValidateData = True Then
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
      
    Sqlqry = " Select * from CNNrates where Code='" & Trim(txtCode) & "' and region='" & Trim(CboRegion) & "' and wtype='" & Trim(cbowtype) & "' and ttime = '" _
                                    & Trim(txttime) & "' and sp_prog='" & cbospprog & "' and quarter='" & txtquarter & "'"
   
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            MsgBox "Record is Already existing "
            Exit Sub
        Else
             Sqlqry1 = " Insert into CNNrates values('" & UCase(Trim(txtCode)) & "','" & Trim(CboRegion) & "','" _
                     & Trim(cbowtype) & "','" _
                     & Trim(cbospprog) & "','" _
                     & Trim(txttime) & "','" _
                     & Trim(txtquarter) & "'," _
                     & Val(Trim(txtRate.Text)) & ")"
                            
              
                ws.BeginTrans
                db.Execute (Sqlqry1)
                ws.CommitTrans
                
                
                 MsgBox "Record is inserted", vbDefaultButton3, "Status"
                 
                 
     
                    textclear
                   
                 populateCNNrates
        
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
CboRegion.ListIndex = -1
txtCode.Text = ""
cbowtype.ListIndex = -1
cbospprog.ListIndex = -1
txttime = ""
txtRate = ""
txtquarter.ListIndex = -1
End Function
Private Function ValidateData()
 ValidateData = False

If CboRegion.Text = "" Then
   MsgBox "Invalid Region", vbInformation, "Invalid Entry"
   CboRegion.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsNumeric(txtRate.Text) = False Then
   MsgBox " Invalid rate", vbInformation, "Invalid Entry"
   txtRate.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf cbowtype.Text = "" Then
   MsgBox "Invalid Day Type", vbInformation, "Invalid Entry"
   cbowtype.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf cbospprog.Text = "" Then
   MsgBox "Invalid Specific Program type", vbInformation, "Invalid Entry"
   cbospprog.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf txtCode.Text = "" Or Len(txtCode) <> 6 Then
   MsgBox "Invalid code", vbInformation, "Invalid Entry"
   txtCode.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf txttime.Text = "" Then
   MsgBox "Invalid  Time", vbInformation, "Invalid Entry"
   txttime.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf txtquarter.Text = "" Then
   MsgBox "Invalid Quarter", vbInformation, "Invalid Entry"
   txtquarter.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

End If
ValidateData = True
End Function

Private Sub cmdDelete_Click()
Dim tempStr
If lstCNNRates.SelCount = 0 Then
        MsgBox "Select the code for deletion in the list box ", vbInformation, "Selection Error"
        lstCNNRates.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Delete the Code : " & txtCode, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
      
                 Sqlqry = " Select * from CNNrates where Code='" & Trim(txtCode) & "' and region='" & Trim(CboRegion) & "' and wtype='" & Trim(cbowtype) & "' and ttime = '" _
                                    & Trim(txttime) & "' and sp_prog='" & cbospprog & "' and quarter='" & txtquarter & "'"
   
                 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                 If rs.RecordCount = 0 Then
                  MsgBox "Record is not existing "
                  Exit Sub
                 End If
            If DeleteData = False Then Exit Sub
            Else
              MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
              lstCNNRates.SetFocus
              Exit Sub
            End If
        
End Sub

Private Sub cmdMod_Click()
Dim tempStr
    If lstCNNRates.SelCount = 0 Then
        MsgBox "Select the CNNrates Code for Modification.", vbInformation, "Selection Error"
        lstCNNRates.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Modify the CNNrates Code :" & CboRegion, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
      
                 Sqlqry = " Select * from CNNrates where Code='" & UCase(Trim(txtCode)) & "' and region='" & Trim(CboRegion) & "' and wtype='" & Trim(cbowtype) & "' and ttime = '" _
                                    & Trim(txttime) & "' and sp_prog='" & cbospprog & "' and quarter='" & txtquarter & "'"
   
                 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                 If rs.RecordCount = 0 Then
                  MsgBox "Record is not existing "
                  Exit Sub
                 End If

            If ModifyData = False Then Exit Sub
             Else
              MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
              lstCNNRates.SetFocus
              Exit Sub
        End If
       
    End Sub
Private Function ModifyData() As Boolean
    Dim i
    ModifyData = False
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
    i = Trim(txtCode.Text)
    
                                                 
        Sqlqry = "Update CNNrates Set Region = '" & Trim(CboRegion.Text) & "'," & _
                  " Wtype = '" & Trim(cbowtype) & "'," & _
                  " Sp_Prog = '" & Trim(cbospprog) & "'," & _
                  " TTime = '" & Trim(txttime.Text) & "'," & _
                  " quarter = '" & Trim(txtquarter) & "'," & _
                  " Rate = " & Val(Trim(txtRate)) & " Where Code ='" & UCase(Trim(txtCode)) & "'"
                                                          
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        
        MsgBox "Record Modified With " & Chr(10) & Chr(10) & _
               "Code = " & i, vbInformation, "Data Modified"
        textclear
        populateCNNrates
        ModifyData = True
    Exit Function
End Function

Private Function DeleteData() As Boolean
  Dim i
    
    DeleteData = False
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
     i = Trim(txtCode)
        
       Sqlqry = "Delete * from CNNrates Where Code = '" & i & "'"
                                              
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        MsgBox "Record Deleted With " & Chr(10) & Chr(10) & _
               "Code = " & i, vbInformation, "Data Modified"
        textclear
        populateCNNrates

End Function

Private Sub populateCNNrates()
    lstCNNRates.Clear
    CboRegion.ListIndex = -1
    cbowtype.ListIndex = -1
    cbospprog.ListIndex = -1
    txttime.Text = ""
    txtCode = ""
    txtquarter.ListIndex = -1
    txtRate.Text = ""
    
    Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from CNNrates order by  region,wtype,ttime"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        rs.MoveFirst
            Do Until rs.EOF
              lstCNNRates.AddItem UCase(rs!code) & " : " & rs!region & " : " & rs!wtype & " : " & rs!ttime & " : " & rs!quarter & " : " & rs!sp_Prog
           rs.MoveNext
       Loop
   End If
 End Sub
 
Private Sub Form_Load()
    txtquarter.AddItem "All"
    txtquarter.AddItem "Q1"
    txtquarter.AddItem "Q2"
    txtquarter.AddItem "Q3"
    txtquarter.AddItem "Q4"
    
    txtquarter.AddItem "Q1 - Q2"
    txtquarter.AddItem "Q1 - Q3"
    txtquarter.AddItem "Q1 - Q4"
    
    txtquarter.AddItem "Q2 - Q1"
    txtquarter.AddItem "Q2 - Q3"
    txtquarter.AddItem "Q2 - Q4"
    
    txtquarter.AddItem "Q3 - Q1"
    txtquarter.AddItem "Q3 - Q2"
    txtquarter.AddItem "Q3 - Q4"
    
    txtquarter.AddItem "Q4 - Q1"
    txtquarter.AddItem "Q4 - Q2"
    txtquarter.AddItem "Q4 - Q3"
    
    cbospprog.AddItem "Specific"
    cbospprog.AddItem "Any"
    
    CboRegion.AddItem "EMEA"
    CboRegion.AddItem "SOUTH ASIA"
    CboRegion.AddItem "ASIA"
    CboRegion.AddItem "LATIN AMERICA"
    CboRegion.AddItem "OTHERS"
    
    cbowtype.AddItem "WEEK DAYS"
    cbowtype.AddItem "WEEK ENDING"
    cbowtype.AddItem "OTHERS"
    
    populateCNNrates
       
    textclear
    
End Sub

Private Sub lstCNNrates_Click()
Dim i
Dim tempBln As String
    
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = UCase(Trim(Mid(lstCNNRates, 1, 6)))
        Sqlqry = "Select * from CNNrates Where Code= '" & i & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        
        If rs.RecordCount <> 0 Then
           CboRegion = rs!region
           cbowtype = rs!wtype
           cbospprog = rs!sp_Prog
           txttime = Trim(rs!ttime)
           txtquarter = rs!quarter
           txtRate = rs!Rate
           txtCode = rs!code
        
         End If
          CboRegion.SetFocus
          
         
End Sub
Private Sub cbowtype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cbospprog.SetFocus
End Sub
Private Sub CboRegion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cbowtype.SetFocus
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtquarter.SetFocus
End Sub
