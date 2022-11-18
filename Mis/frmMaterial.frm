VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmmaterial 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Products"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame FraMaterial 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Material Details"
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
         TabIndex        =   14
         Top             =   600
         Width           =   11055
         Begin VB.TextBox txtTime 
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
            TabIndex        =   3
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox txtCode 
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
            MaxLength       =   4
            ScrollBars      =   1  'Horizontal
            TabIndex        =   1
            Top             =   2160
            Width           =   1335
         End
         Begin VB.ListBox lstMaterial 
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
            ForeColor       =   &H00800000&
            Height          =   3660
            Left            =   7200
            TabIndex        =   11
            Top             =   1800
            Width           =   3615
         End
         Begin VB.ComboBox cboMedia 
            BackColor       =   &H00C0FFC0&
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
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   4680
            Width           =   2895
         End
         Begin VB.TextBox txtName 
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
            TabIndex        =   2
            Top             =   3000
            Width           =   4575
         End
         Begin VB.ComboBox cboProduct 
            BackColor       =   &H00C0FFC0&
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
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   960
            Width           =   6615
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Actual Time"
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
            Left            =   495
            TabIndex        =   19
            Top             =   3960
            Width           =   1425
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Material Code"
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
            Left            =   120
            TabIndex        =   18
            Top             =   2280
            Width           =   1800
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
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
            Left            =   120
            TabIndex        =   17
            Top             =   4800
            Width           =   1815
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Products"
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
            Left            =   240
            TabIndex        =   16
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Material Name"
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
            Left            =   240
            TabIndex        =   15
            Top             =   3120
            Width           =   1740
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Height          =   1215
         Left            =   960
         TabIndex        =   13
         Top             =   6960
         Width           =   9255
         Begin VB.CommandButton cmdprintall 
            BackColor       =   &H00FFFF80&
            Caption         =   "Pre&view All"
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
            Picture         =   "frmMaterial.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00FFFF80&
            Caption         =   "&Add"
            DisabledPicture =   "frmMaterial.frx":0102
            DownPicture     =   "frmMaterial.frx":0634
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
            Picture         =   "frmMaterial.frx":0B66
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            UseMaskColor    =   -1  'True
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
            Picture         =   "frmMaterial.frx":0FA8
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1215
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
            Left            =   6120
            Picture         =   "frmMaterial.frx":13EA
            Style           =   1  'Graphical
            TabIndex        =   9
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
            Left            =   7320
            Picture         =   "frmMaterial.frx":14EC
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1335
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
            Picture         =   "frmMaterial.frx":15EE
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
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
            Picture         =   "frmMaterial.frx":1A30
            Style           =   1  'Graphical
            TabIndex        =   8
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
Attribute VB_Name = "frmmaterial"
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


Private Sub cbomedia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub
Private Sub cboProduct_Click()
 populateMaterial
End Sub
Private Sub cboProduct_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtCode.SetFocus
End Sub
Private Sub cmdprintall_Click()
    CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
    CrystalReport1.ReportFileName = App.Path & "\MaterialList.rpt"
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
      
    Sqlqry = " Select * from Material where Code='" & UCase(Trim(txtCode)) & "' "
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
         MsgBox "Material Code Already existed "
         Exit Sub
        Else
    Sqlqry1 = " Insert into Material values('" & UCase(Trim(txtCode)) & "','" _
              & findfirstfixup(Trim(UCase(txtName))) & "','" _
              & Trim(txttime) & "','" _
              & Trim(CboMedia.Text) & "','" _
              & findfirstfixup(UCase(Trim(CboProduct.Text))) & "')"
                            
              
                ws.BeginTrans
                db.Execute (Sqlqry1)
                ws.CommitTrans
                
                 MsgBox "Record is inserted", vbDefaultButton3, "Status"
                 textclear
                 populateMaterial
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
txtCode = ""
txtName = ""
txttime = ""
CboProduct.ListIndex = -1
CboMedia.ListIndex = -1
End Function
Private Function ValidateData()
 ValidateData = False

If txtCode.Text = "" Then
   MsgBox "Invalid Material Code", vbInformation, "Invalid Entry"
   txtCode.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf Len(txtCode.Text) <> 4 Then
   MsgBox "Re-enter Material Code Minimum  should be 4 characters", vbInformation, "Invalid Entry"
   txtCode.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf txtName.Text = "" Then
   MsgBox "Invalid Material name", vbInformation, "Invalid Entry"
   txtName.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
   
ElseIf txttime.Text = "" Then
   MsgBox "Invalid Duration / Time", vbInformation, "Invalid Entry"
   txttime.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf CboMedia.Text = "" Then
   MsgBox "Media is not selected", vbInformation, "Invalid Entry"
   CboMedia.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf CboProduct.Text = "" Then
   MsgBox "Product Name is not selected", vbInformation, "Invalid Entry"
   CboProduct.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
End If
ValidateData = True
End Function
Private Function ValidateData1()
 ValidateData1 = False

If CboProduct.Text = "" Then
   MsgBox "Select Product from product list", vbInformation, "Invalid Entry"
   CboProduct.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
End If
ValidateData1 = True
End Function
Private Sub cmdDelete_Click()
Dim tempStr
If lstMaterial.SelCount = 0 Then
        MsgBox "Select the Material Name for Deletion.", vbInformation, "Selection Error"
        lstMaterial.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Delete the Material Code : " & txtCode, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If DeleteData = False Then Exit Sub
        Else
            MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
            lstMaterial.SetFocus
            Exit Sub
        End If
End Sub

Private Sub cmdMod_Click()
Dim tempStr
    If lstMaterial.SelCount = 0 Then
        MsgBox "Select the Material Code for Modification.", vbInformation, "Selection Error"
        lstMaterial.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Modify the Material Code :" & txtCode, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If ModifyData = False Then Exit Sub
        Else
              MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
              lstMaterial.SetFocus
              Exit Sub
        End If
    End Sub
Private Function ModifyData() As Boolean
    Dim i
    ModifyData = False
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
    i = UCase(Trim(txtCode.Text))
    
                                                 
        Sqlqry = "Update Material Set Code = '" & UCase(Trim(txtCode.Text)) & "'," & _
                  " Name = '" & findfirstfixup(UCase(Trim(txtName))) & "'," & _
                  " TTime = '" & Trim(txttime.Text) & "'," & _
                  " Media = '" & Trim(CboMedia) & "'," & _
                  " Product = '" & findfirstfixup(UCase(Trim(CboProduct))) & "' Where code ='" & UCase(Trim(txtCode.Text)) & "'"
                                                          
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        
        MsgBox "Record Modified With " & Chr(10) & Chr(10) & _
               "Material Code = " & i, vbInformation, "Data Modified"
        textclear
        populateMaterial
        ModifyData = True
    Exit Function
End Function

Private Function DeleteData() As Boolean
  Dim i
    
    DeleteData = False
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
     i = UCase(Trim(txtCode.Text))
        
       Sqlqry = "Delete * from Material Where Code = '" & i & "'"
                                              
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        MsgBox "Record Deleted With " & Chr(10) & Chr(10) & _
               "Material Code = " & i, vbInformation, "Data Modified"
        textclear
        populateMaterial

End Function

Private Sub populateMaterial()
    lstMaterial.Clear
    txtCode.Text = ""
    txtName.Text = ""
    txttime.Text = ""
    CboMedia.ListIndex = -1
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Material  where product='" & Trim(CboProduct.Text) & "'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        rs.MoveFirst
            Do Until rs.EOF
              lstMaterial.AddItem rs!code & " " & rs!Name
            rs.MoveNext
       Loop
    End If
 End Sub

Private Sub populateproduct()
    CboProduct.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Products order by product_name"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        rs.MoveFirst
            Do Until rs.EOF
              CboProduct.AddItem rs!product_name
            rs.MoveNext
       Loop
    End If
 End Sub

Private Sub CmdPrint_Click()
If ValidateData1 = True Then

    CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
    CrystalReport1.ReportFileName = App.Path & "\MaterialList.rpt"
    CrystalReport1.SelectionFormula = "{Material.product}='" & CboProduct.Text & "'"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
End If
 
End Sub


Private Sub Form_Load()

    populateproduct
    CboMedia.AddItem "All"
    CboMedia.AddItem "Cinema"
    CboMedia.AddItem "Magazine"
    CboMedia.AddItem "Online"
    CboMedia.AddItem "Television"
    textclear
End Sub

Private Sub lstMaterial_Click()
Dim i
Dim tempBln As String
    If lstMaterial.ListIndex = -1 Then
        tempBln = False
    Else
        tempBln = True
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Trim(Mid(lstMaterial.Text, 1, 4))
        Sqlqry = "Select * from Material Where code= '" & i & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        
        If rs.RecordCount <> 0 Then
           txtCode = rs!code
           txtName = rs!Name
           txttime.Text = rs!ttime
           CboMedia.Text = rs!media
         End If
          txtCode.SetFocus
          SendKeys "{home}+{end}"
         
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txttime.SetFocus
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtName.SetFocus
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboMedia.SetFocus
End Sub
