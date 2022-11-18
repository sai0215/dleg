VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmMediaType 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Media Type"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FraMedia 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Media Type and  Sub Media"
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
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   11535
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Height          =   1215
         Left            =   1680
         TabIndex        =   12
         Top             =   6960
         Width           =   7575
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00FFFF80&
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
            Picture         =   "frmMediaType.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmddelete 
            BackColor       =   &H00FFFF80&
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
            Picture         =   "frmMediaType.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdBack 
            BackColor       =   &H00FFFF80&
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
            Picture         =   "frmMediaType.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H00FFFF80&
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
            Left            =   4920
            Picture         =   "frmMediaType.frx":0646
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdMod 
            BackColor       =   &H00FFFF80&
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
            Picture         =   "frmMediaType.frx":0748
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00FFFF80&
            Caption         =   "&Add"
            DisabledPicture =   "frmMediaType.frx":0B8A
            DownPicture     =   "frmMediaType.frx":10BC
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
            Picture         =   "frmMediaType.frx":15EE
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame frmAgency 
         BackColor       =   &H00FFFFC0&
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
         Height          =   6135
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   11055
         Begin VB.ComboBox cbomedia 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2040
            Width           =   4575
         End
         Begin VB.TextBox txtSubMedia 
            BackColor       =   &H00C0FFC0&
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
            Left            =   2280
            ScrollBars      =   1  'Horizontal
            TabIndex        =   1
            Top             =   3120
            Width           =   4575
         End
         Begin VB.ListBox lstMedia 
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
            ForeColor       =   &H00800000&
            Height          =   5160
            Left            =   7200
            TabIndex        =   0
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Height          =   300
            Left            =   360
            TabIndex        =   11
            Top             =   3120
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Media Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   10
            Top             =   2160
            Width           =   1380
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
Attribute VB_Name = "frmMediaType"
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

Private Sub cboAgent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
 End Sub

Private Sub cmdadd_Click()

  If ValidateData = True Then
  
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
      
    Sqlqry = " Select * from Media where Sub_Media='" & Trim(txtSubMedia) & "' "
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
         MsgBox " Sub-Media Already existing in the records"
         Exit Sub
        Else
    Sqlqry1 = " Insert into Media values('" & cbomedia & "','" _
              & findfirstfixup(Trim(UCase(txtSubMedia))) & "')"
                            
              
                ws.BeginTrans
                db.Execute (Sqlqry1)
                ws.CommitTrans
                
                 MsgBox "Record is inserted", vbDefaultButton3, "Status"
                 textclear
                 populateMedia
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

cbomedia.ListIndex = -1
txtSubMedia = ""

End Function

Private Function ValidateData()

 ValidateData = False

If cbomedia.Text = "" Then
   MsgBox "Invalid Media Type", vbInformation, "Invalid Entry"
   cbomedia.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf txtSubMedia.Text = "" Then
   MsgBox "Invalid Sub_Media", vbInformation, "Invalid Entry"
   txtSubMedia.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
End If
ValidateData = True
End Function

Private Sub cmdDelete_Click()
Dim tempStr
If lstMedia.SelCount = 0 Then
        MsgBox "Select the Sub_Media for Deletion.", vbInformation, "Selection Error"
        lstMedia.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Delete the Product Name : " & cbomedia, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If DeleteData = False Then Exit Sub
        Else
            MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
            lstMedia.SetFocus
            Exit Sub
        End If
End Sub

Private Sub cmdMod_Click()

Dim tempStr

    If lstMedia.SelCount = 0 Then
        MsgBox "Select the Sub_media for Modification.", vbInformation, "Selection Error"
        lstMedia.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Modify the Sub_Media :" & cbomedia, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If ModifyData = False Then Exit Sub
        Else
              MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
              lstMedia.SetFocus
              Exit Sub
        End If
    End Sub

Private Function ModifyData() As Boolean
    Dim i
    ModifyData = False
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
    i = Trim(lstMedia.Text)
                
        Sqlqry = "Update Media Set " _
                  & " Media_Type = '" & Trim(cbomedia) & "'," _
                  & " sub_Media = '" & findfirstfixup(Trim(UCase(txtSubMedia))) & "'" _
                  & " Where Sub_Media ='" & findfirstfixup(i) & "'"
                                                
                                                     
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        populatesm
        MsgBox "Record Modified With " & Chr(10) & Chr(10) & _
               " Sub_Media  = " & i, vbInformation, "Data Modified"
        textclear
        populateMedia
      '  TEMPBLN = False
        ModifyData = True
        Exit Function
End Function

Private Sub populatesm()
 Dim i, j
 i = UCase(Trim(txtSubMedia.Text))
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       j = Trim(lstMedia.Text)
        Sqlqry = "Select sub_media from bo_mas  where sub_media ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            
              Sqlqry1 = "Update bo_mas set sub_media = '" & findfirstfixup(i) & "' WHERE sub_media='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
        Sqlqry = "Select sub_media from bo_tracin  where sub_media ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            
              Sqlqry1 = "Update bo_tracin set sub_media = '" & findfirstfixup(i) & "' WHERE sub_media='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
            
            
        Sqlqry = "Select sub_media from bo_tramag  where sub_media ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            
              Sqlqry1 = "Update bo_tramag set sub_media = '" & findfirstfixup(i) & "' WHERE sub_media='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
            
        Sqlqry = "Select sub_media from bo_tratv  where sub_media ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            
              Sqlqry1 = "Update bo_tratv set sub_media = '" & findfirstfixup(i) & "' WHERE sub_media='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
            
        Sqlqry = "Select sub_media from bo_traol  where sub_media ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            
              Sqlqry1 = "Update bo_traol set sub_media = '" & findfirstfixup(i) & "' WHERE sub_media='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
 
End Sub
Private Function DeleteData() As Boolean
  Dim i
    
    DeleteData = False
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
     i = Trim(lstMedia.Text)
     
       Sqlqry1 = "select * from bo_mas where sub_media='" & findfirstfixup(i) & "'"
       Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            MsgBox " Cannot delete transactions already existing"
            Exit Function
        End If
        
       Sqlqry = "Delete * from Media Where Sub_Media = '" & findfirstfixup(i) & "'"
                                              
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        MsgBox "Record Deleted With " & Chr(10) & Chr(10) & _
               "Sub_Media = " & i, vbInformation, "Data Modified"
        textclear
        populateMedia
              
End Function

Private Sub populateMedia()
    lstMedia.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Media Order by Sub_Media"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        rs.MoveFirst
            Do Until rs.EOF
              lstMedia.AddItem rs!sub_media
            rs.MoveNext
       Loop
    End If
 End Sub

Private Sub CmdPrint_Click()
    CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
    CrystalReport1.ReportFileName = App.Path & "\MediaList.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
End Sub

Private Sub Form_Load()
    populateMedia
    cbomedia.AddItem "Cinema"
    cbomedia.AddItem "Magazine"
    cbomedia.AddItem "Online"
    cbomedia.AddItem "Television"
    textclear
    
End Sub

Private Sub lstmedia_Click()
Dim i

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Trim(lstMedia.Text)
        Sqlqry = "Select * from Media Where Sub_Media= '" & findfirstfixup(i) & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MsgBox "Data Mis Matching", vbInformation, "Deleted Status"
            Exit Sub
        Else
          rs.MoveFirst
            cbomedia.Text = rs!Media_Type
        
          ' If Mid(rs!media_type, 1, 3) = "Cin" Then
          '     cbomedia.Text = "Cinema"
          ' ElseIf Mid(rs!media_type, 1, 3) = "Mag" Then
          '     cbomedia.ListIndex = 1
          '     cbomedia.Text = "Magazine"
           'ElseIf Mid(rs!media_type, 1, 3) = "Onl" Then
           '   'cbomedia.ListIndex = 2
           '   cbomedia.Text = "Online"
           'ElseIf Mid(rs!media_type, 1, 3) = "Tel" Then
           '  ' cbomedia.ListIndex = 3
           '   cbomedia.Text = "Television"
           'End If
              
           
           If IsNull(rs!sub_media) = True Then
              txtSubMedia = ""
           Else
              txtSubMedia = rs!sub_media
           End If
           
        End If
         ' cbomedia.SetFocus
         ' SendKeys "{home}+{end}"
         
End Sub
Private Sub txtsubmedia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub
Private Sub cbomedia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtSubMedia.SetFocus
End Sub




