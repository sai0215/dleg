VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmMagRates 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12060
   LinkTopic       =   "form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   360
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Magazine Rates"
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
      Height          =   7095
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   11655
      Begin VB.Frame fraAccount 
         BackColor       =   &H80000005&
         Height          =   4695
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   10815
         Begin VB.CommandButton cmdEdit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtmodspace 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6000
            TabIndex        =   2
            Top             =   1440
            Width           =   3615
         End
         Begin VB.ComboBox CboSpace 
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
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtremarks 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6000
            TabIndex        =   4
            Top             =   2910
            Width           =   3615
         End
         Begin VB.TextBox txtamount 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6000
            TabIndex        =   3
            Top             =   2190
            Width           =   1575
         End
         Begin VB.ListBox lstsubmedia 
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
            Height          =   3765
            Left            =   240
            TabIndex        =   0
            Top             =   360
            Width           =   3975
         End
         Begin VB.TextBox txtpage 
            BackColor       =   &H00FFFFFF&
            DataField       =   "ACCT_NAME"
            DataSource      =   "Data1"
            Enabled         =   0   'False
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
            Left            =   3120
            ScrollBars      =   1  'Horizontal
            TabIndex        =   11
            Top             =   3000
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Space"
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
            Left            =   5160
            TabIndex        =   15
            Top             =   1560
            Width           =   690
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Page"
            Enabled         =   0   'False
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
            Left            =   1920
            TabIndex        =   14
            Top             =   3120
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Amount"
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
            Left            =   5100
            TabIndex        =   13
            Top             =   2280
            Width           =   780
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Remarks"
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
            Left            =   4965
            TabIndex        =   12
            Top             =   3000
            Width           =   945
         End
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
         Left            =   6480
         Picture         =   "frmMagRates.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5760
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
         Left            =   5400
         Picture         =   "frmMagRates.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5760
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
         Left            =   4320
         Picture         =   "frmMagRates.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Add"
         DisabledPicture =   "frmMagRates.frx":0CC6
         DownPicture     =   "frmMagRates.frx":1108
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
         Left            =   3240
         MaskColor       =   &H008080FF&
         Picture         =   "frmMagRates.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMagRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim j
Dim I

Private Sub CboSpace_Click()

Dim I

   ' If lstsubmedia.ListIndex = -1 Then
   '     tempBln = False
   ' Else
   '     tempBln = True
   ' End If
   
    txtmodspace.Text = ""
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    I = CboSpace.Text
        Sqlqry = "Select * from Mag_rates Where Space= '" & I & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            
            rs.MoveFirst
             
            txtmodspace = CboSpace.Text
            
            If IsNull(rs!Amount) = True Then
               txtamount = 0
            Else
               txtamount = rs!Amount
            End If
            
            If IsNull(rs!remarks) = True Then
               txtremarks = ""
            Else
               txtremarks = rs!remarks
            End If
            
            If IsNull(rs!Page) = True Then
               txtpage = ""
            Else
               txtpage = rs!Page
            End If
            
            
       End If
    
         
    txtmodspace.SetFocus
    SendKeys "{home}+{end}"
End Sub

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
    
   Sqlqry1 = " select * from Mag_rates where Space='" & Trim(txtmodspace) & "'  "
   Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
   If rs.RecordCount <> 0 Then
      MsgBox " Record already exists"
      Exit Sub
   Else
        
   Sqlqry = " Insert into Mag_rates values('" & UCase(lstsubmedia.Text) & "','" & UCase(txtmodspace) & "','" & _
            findfirstfixup(Trim(txtpage)) & "','" & _
            Trim(txtamount) & "','" & _
            Trim(txtremarks) & "')"
            
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    Populatesubmedia
    MsgBox "Record is inserted", vbDefaultButton3, "Status"
    Populatespace
    CboSpace.ListIndex = -1
    CboSpace.Visible = False
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
 CboSpace.ListIndex = -1
 txtmodspace = ""
 'txtPage = ""
 txtamount = "0"
 txtremarks = ""
End Function

Private Sub cmdEdit_Click()
CboSpace.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub

Private Function ValidateData()

ValidateData = False

If txtmodspace.Text = "" Then
   MsgBox "Invalid Space", vbInformation, "Invalid Entry"
   txtmodspace.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsNumeric(txtamount) = False Then
   MsgBox "Invalid Amount", vbInformation, "Invalid Entry"
   txtamount.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
  ValidateData = True
End If

End Function

'Private Sub cmdDelete_Click()
' If lstsubmedia.SelCount = 0 Then
'        MsgBox "Select the e for Deletion.", vbInformation, "Selection Error"
'        lstsubmedia.SetFocus
'        Exit Sub
'    End If
'        If ValidateData = False Then Exit Sub
          
'           i = Trim(cbospace.Text)
           
           
'        If txtAmount.Text <> 0 And txtremarks.Text <> 0 Then
'           MsgBox "You can not Delete since the transactions are recorded"
'           Exit Sub
'        End If
        
'        tempStr = MsgBox("Do You Want To Delete the Account Code : " & cbospace, vbQuestion + vbYesNo, "Confirmation")
'        If tempStr = vbYes Then
'            If DeleteData = False Then Exit Sub
'        Else
'            MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
'            cbospace.SetFocus
'            Exit Sub
'        End If
'End Sub

Private Sub cmdMod_Click()
Dim I
  Z = 1
    If lstsubmedia.SelCount = 0 Then
        MsgBox "Select the Sub Media for Modification.", vbInformation, "Selection Error"
        lstsubmedia.SetFocus
        Exit Sub
    End If
    
    
    
        If ValidateData = False Then Exit Sub
        
        If UCase(CboSpace.Text) = UCase(txtmodspace.Text) Then
           MsgBox " Record already exists"
           Exit Sub
        End If
           
          I = Trim(CboSpace.Text)
         
           Set ws = DBEngine.Workspaces(0)
           Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           Sqlqry = "Select * from Mag_rates where Space='" & I & "'"
           Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
            
           
                tempStr = MsgBox("Do You Want To Modify the Space & Page :" & CboSpace, vbQuestion + vbYesNo, "Confirmation")
                If tempStr = vbYes Then
                    If ModifyData = False Then Exit Sub
                Else
                    MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
                    CboSpace.SetFocus
                    Exit Sub
                End If
          End If
 End Sub

Private Function ModifyData() As Boolean
    Dim I
    ModifyData = False
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
    I = Trim(CboSpace.Text)
        
        
        Sqlqry = "Update Mag_rates Set " _
                   & " Space = '" & UCase(txtmodspace) & "'," _
                   & " Page = '" & txtpage & "'," _
                   & " Amount = " & Val(txtamount.Text) & "," _
                   & " Remarks = '" & Val(txtremarks.Text) & "' " _
                   & " Where Space = '" & I & "'"
                                           
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        MsgBox "Record Modified With " & Chr(10) & Chr(10) & _
               "Space = " & I, vbInformation, "Data Modified"
        Populatespace
        textclear
        Populatesubmedia
        CboSpace.Visible = False
        tempBln = False
        ModifyData = True
        Exit Function
End Function

'Private Function DeleteData() As Boolean
' Dim i
    
'    DeleteData = False
    
'    Set ws = DBEngine.Workspaces(0)
'    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           
'    i = Trim(cbospace.Text)
               
'       If txtAmount > 0 Or txtremarks > 0 Then
'         MsgBox " Account Cannot be Deleted since the transactions are recorded"
'         DeleteData = False
'         Exit Function
'       Else
           
'        Sqlqry1 = "Select * from bank_mas where bank_code='" & i & "'"
'        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
'        If rs.RecordCount = 0 Then
           
'           Sqlqry = "Delete * from Mag_rates Where Space = '" & i & "'"
                                           
'            ws.BeginTrans
'            db.Execute (Sqlqry)
'            ws.CommitTrans
'            MsgBox "Record Deleted With " & Chr(10) & Chr(10) & _
'               "Account Code = " & i, vbInformation, "Data Modified"
'            textclear
'            PopulateAccodes
'            tempBln = False
'            If Validate1 = False Then Exit Function
'            DeleteData = True
'            Exit Function
'          Else
'            MsgBox "You cannot Delete Bank Code", vbInformation, "Invalid Attempt"
'            DeleteData = False
'            Exit Function
'          End If
'        End If
          
'End Function
Private Sub Populatespace()
Dim I

    If lstsubmedia.ListIndex = -1 Then
        tempBln = False
    Else
        tempBln = True
    End If
     CboSpace.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    I = lstsubmedia.Text
        Sqlqry = "Select * from Mag_rates Where Sub_media= '" & I & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            
            rs.MoveFirst
            CboSpace.Clear
          Do Until rs.EOF
            CboSpace.AddItem rs!Space
            'lstsubmedia.AddItem rs!space & "    :    " & rs!Page
            rs.MoveNext
         Loop
       End If
    
         
    txtmodspace.SetFocus
    SendKeys "{home}+{end}"
   
End Sub
Private Sub Populatesubmedia()
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select distinct(sub_media) from Media where Media_type='Magazine' order by sub_Media"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
     rs.MoveFirst
        lstsubmedia.Clear
        Do Until rs.EOF
            lstsubmedia.AddItem rs!SUB_MEDIA
            'lstsubmedia.AddItem rs!space & "    :    " & rs!Page
            rs.MoveNext
        Loop
    End If
        
End Sub

'Private Sub CmdPrint_Click()
' CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
' CrystalReport1.ReportFileName = App.Path & "\AccountList.rpt"
' CrystalReport1.WindowState = crptMaximized
' CrystalReport1.Action = 1
'End Sub

Private Sub Form_Load()
    tempBln = False
    Populatesubmedia
    CboSpace.Visible = False
    textclear
    'txtAmount.Text = 0
End Sub

Private Sub lstSubMedia_Click()
Dim I

    If lstsubmedia.ListIndex = -1 Then
        tempBln = False
    Else
        tempBln = True
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    I = lstsubmedia.Text
    CboSpace.Clear
        Sqlqry = "Select * from Mag_rates Where Sub_media= '" & I & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            
            rs.MoveFirst
            CboSpace.Clear
          Do Until rs.EOF
            CboSpace.AddItem rs!Space
            'lstsubmedia.AddItem rs!space & "    :    " & rs!Page
            rs.MoveNext
         Loop
       End If
    
         
    txtmodspace.SetFocus
    SendKeys "{home}+{end}"
   
End Sub

 
 
Private Sub cbospace_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmodspace.SetFocus
End Sub

Private Sub txtmodspace_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtamount.SetFocus
End Sub

Private Sub txtremarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub

Private Sub txtamount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtremarks.SetFocus
End Sub

