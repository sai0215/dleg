VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmProducts 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Products"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame FraProduct1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Product Details"
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
         Height          =   6135
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   11055
         Begin VB.TextBox txtcommission 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   5
            Top             =   5040
            Width           =   495
         End
         Begin VB.TextBox txtproduct 
            BackColor       =   &H00FFFFFF&
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
            Top             =   1440
            Width           =   4575
         End
         Begin VB.ListBox lstProducts 
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
            ForeColor       =   &H00800000&
            Height          =   5160
            Left            =   7200
            TabIndex        =   0
            Top             =   600
            Width           =   3615
         End
         Begin VB.ComboBox cboAgent 
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
            Height          =   420
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   4200
            Width           =   4575
         End
         Begin VB.TextBox txtCategory 
            BackColor       =   &H00FFFFFF&
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
            Top             =   2400
            Width           =   4575
         End
         Begin VB.ComboBox cboClient 
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
            Height          =   420
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   3360
            Width           =   4575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "  Commission"
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
            TabIndex        =   19
            Top             =   5040
            Width           =   1605
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Product"
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
            TabIndex        =   18
            Top             =   1560
            Width           =   945
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFFFFF&
            Caption         =   "  Agency Name"
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
            Left            =   240
            TabIndex        =   17
            Top             =   4200
            Width           =   1815
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "  Client Name"
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
            Left            =   240
            TabIndex        =   16
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "  Category"
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
            Top             =   2400
            Width           =   1245
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   840
         TabIndex        =   13
         Top             =   6960
         Width           =   9015
         Begin VB.CommandButton cmdprintall 
            BackColor       =   &H00FFFFFF&
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
            Picture         =   "frmProducts.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            DisabledPicture =   "frmProducts.frx":0102
            DownPicture     =   "frmProducts.frx":0634
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
            Picture         =   "frmProducts.frx":0B66
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton cmdMod 
            BackColor       =   &H00FFFFFF&
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
            Picture         =   "frmProducts.frx":0FA8
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H00FFFFFF&
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
            Picture         =   "frmProducts.frx":13EA
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdBack 
            BackColor       =   &H00FFFFFF&
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
            Picture         =   "frmProducts.frx":14EC
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmddelete 
            BackColor       =   &H00FFFFFF&
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
            Picture         =   "frmProducts.frx":15EE
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00FFFFFF&
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
            Picture         =   "frmProducts.frx":1A30
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
Attribute VB_Name = "frmProducts"
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
Dim X, Y, Z As Integer
Private Sub cboAgent_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then txtcommission.SetFocus
End Sub
Private Sub cboAgent_LostFocus()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Select * from Agndtls where Agentname='" & findfirstfixup(Trim(cboAgent.Text)) & "' "
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
         txtcommission.Text = Val(rs!Discount)
         txtcommission.SetFocus
         Exit Sub
        End If
 
    
End Sub
Private Sub CboClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboAgent.SetFocus
End Sub
Private Sub cboClient_LostFocus()
    Z = 0
    If cboAgent.Text = "" Then Z = 1
End Sub

Private Sub cmdprintall_Click()
    CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
    CrystalReport1.ReportFileName = App.Path & "\ProductList.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Exit Sub
End Sub
Private Sub cmdadd_Click()

    If ValidateData = True Then
  
      Set ws = DBEngine.Workspaces(0)
      Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        
      Sqlqry = " Select * from Products where product_name='" & UCase(Trim(txtproduct)) & "' "
      Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
          If rs.RecordCount <> 0 Then
           MsgBox " Product already existing"
           Exit Sub
          Else
      Sqlqry1 = " Insert into Products values('" & findfirstfixup(UCase(txtproduct)) & "','" _
                & UCase(Trim(txtCategory)) & "','" _
                & findfirstfixup(Trim(CboClient.Text)) & "','" _
                & findfirstfixup(Trim(cboAgent.Text)) & "','" _
                & Val(Trim(txtcommission.Text)) & "')"
                              
                
                  ws.BeginTrans
                  db.Execute (Sqlqry1)
                  ws.CommitTrans
                  
                   MsgBox "Record is inserted", vbDefaultButton3, "Status"
                   textclear
                   populateproducts
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
txtproduct = ""
txtCategory = ""
CboClient.ListIndex = -1
cboAgent.ListIndex = -1
txtcommission.Text = ""
End Function

Private Function ValidateData()
 Dim i
 ValidateData = False
 
        

If txtproduct.Text = "" Then
   MsgBox "Invalid Product Name", vbInformation, "Invalid Entry"
   txtproduct.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf cboAgent.Text = "" Then
   MsgBox "Agency Name is not selected", vbInformation, "Invalid Entry"
   cboAgent.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
   
ElseIf CboClient.Text = "" Then
   MsgBox "Client Name is not selected", vbInformation, "Invalid Entry"
   CboClient.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
End If

ValidateData = True


End Function

Private Function VALIDATEDEL()
Dim i
    VALIDATEDEL = False
   Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Trim(lstProducts.Text)
        Sqlqry = "Select * from bo_mas Where Product= '" & i & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
          VALIDATEDEL = True
        Else
          VALIDATEDEL = False
            MsgBox "Product cannot delete since transactions are recorded", vbInformation, "Deleted Status"
            Exit Function
        End If
    
End Function


Private Function ValidateData1()

 ValidateData1 = False

If lstProducts.Text = "" Then
   MsgBox "Select Product from the product list", vbInformation, "Invalid Entry"
   lstProducts.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

End If
ValidateData1 = True
End Function


Private Sub cmdDelete_Click()
Dim tempStr
If lstProducts.SelCount = 0 Then
        MsgBox "Select the Product Name for Deletion.", vbInformation, "Selection Error"
        lstProducts.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
        If VALIDATEDEL = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Delete the Product Name : " & txtproduct, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If DeleteData = False Then Exit Sub
        Else
            MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
            lstProducts.SetFocus
            Exit Sub
        End If
End Sub

Private Sub cmdMod_Click()

Dim tempStr
    If lstProducts.SelCount = 0 Then
        MsgBox "Select the Product Name for Modification.", vbInformation, "Selection Error"
        lstProducts.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Modify the product Details :" & txtproduct, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If ModifyData = False Then Exit Sub
        Else
              MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
              lstProducts.SetFocus
              Exit Sub
        End If
    End Sub

Private Function ModifyData() As Boolean
    Dim i
    ModifyData = False
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
      i = Trim(UCase(txtproduct))
           
        Sqlqry = "Update Products Set " _
                  & " Product_Name = '" & findfirstfixup(UCase(Trim(txtproduct))) & "'," _
                  & " Category = '" & UCase(Trim(txtCategory)) & "'," _
                  & " Client_name = '" & findfirstfixup(Trim(CboClient.Text)) & "'," _
                  & " Agent_Name = '" & findfirstfixup(Trim(cboAgent.Text)) & "'," _
                  & " Discount = '" & Val(txtcommission.Text) & "'" _
                  & " Where product_name ='" & findfirstfixup(UCase(Trim(lstProducts))) & "'"
                                                
                                                     
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        populateprd
        MsgBox "Record Modified With " & Chr(10) & Chr(10) & _
               "Product Name = " & i, vbInformation, "Data Modified"
        textclear
        If UCase(Trim(txtproduct)) <> UCase(Trim(lstProducts)) Then
          populateprd
        End If
      '  TEMPBLN = False
        ModifyData = True
        Exit Function
End Function
Private Sub populateprd()
 Dim i, j
 i = UCase(Trim(txtproduct.Text))
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       j = UCase(Trim(lstProducts.Text))
        Sqlqry = "Select Product from bo_mas  where Product ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update bo_mas set Product = '" & findfirstfixup(i) & "' WHERE Product='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
        Sqlqry = "Select Product from bo_tracin  where Product ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update bo_tracin set Product = '" & findfirstfixup(i) & "' WHERE Product='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
            
            
        Sqlqry = "Select Product from bo_tramag  where Product ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update bo_tramag set Product = '" & findfirstfixup(i) & "' WHERE Product='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
            
        Sqlqry = "Select Product from bo_tratv  where Product ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
             Sqlqry1 = "Update bo_tratv set Product = '" & findfirstfixup(i) & "' WHERE Product='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
            
        Sqlqry = "Select Product from bo_traol  where Product ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            Sqlqry1 = "Update bo_traol set Product = '" & findfirstfixup(i) & "' WHERE Product='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
        Sqlqry = "Select product from material  where product ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            Sqlqry1 = "Update material set  product= '" & findfirstfixup(i) & "' WHERE product='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
      
End Sub

Private Function DeleteData() As Boolean
 Dim i
    
    DeleteData = False
    i = Trim(UCase(txtproduct.Text))
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    
    Sqlqry = "Select product from bo_mas where product='" & findfirstfixup(i) & "'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              MsgBox "Transactions are recorded, cannot delete Product . . . "
              textclear
              populateproducts
              Exit Function
        Else
             Sqlqry1 = "Delete * from Products Where Product_name = '" & i & "'"
             ws.BeginTrans
             db.Execute (Sqlqry1)
             ws.CommitTrans
             MsgBox "Record Deleted With " & Chr(10) & Chr(10) & _
                    "Product Name = " & i, vbInformation, "Data Modified"
             textclear
             populateproducts
        End If
        
End Function
Private Sub populateagencylist()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Agndtls Order by AgentName"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
        rs.MoveFirst
        cboAgent.Clear
        Do Until rs.EOF
            cboAgent.AddItem rs!agentname
            rs.MoveNext
        Loop
    End If
    
End Sub

Private Sub PopulateClientlist()
    textclear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Clientdtls Order by ClientName"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
        rs.MoveFirst
        CboClient.Clear
        Do Until rs.EOF
            CboClient.AddItem rs!clientname
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub populateproducts()
    lstProducts.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from products Order by Product_Name"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        rs.MoveFirst
            Do Until rs.EOF
              lstProducts.AddItem rs!product_name
            rs.MoveNext
       Loop
    End If
End Sub

Private Sub CmdPrint_Click()
If ValidateData1 = True Then
    CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
    CrystalReport1.ReportFileName = App.Path & "\ProductList.rpt"
    CrystalReport1.SelectionFormula = "{products.product_name}='" & lstProducts.Text & "'"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
End If
End Sub

Private Sub Form_Load()
    populateproducts
    textclear
    PopulateClientlist
    populateagencylist
    
End Sub

Private Sub lstProducts_Click()
Dim i
Dim tempBln As String
    If lstProducts.ListIndex = -1 Then
        tempBln = False
    Else
        tempBln = True
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Trim(lstProducts.Text)
        Sqlqry = "Select * from Products Where Product_name= '" & i & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MsgBox "Data Mis Matching", vbInformation, "Deleted Status"
            Exit Sub
        End If
           txtproduct = rs!product_name
           
           If IsNull(rs!Category) = True Then
              txtCategory = ""
           Else
              txtCategory = rs!Category
           End If
           
           CboClient.Text = rs!CLIENT_NAME
           cboAgent.Text = rs!AGENT_NAME
              
           If IsNull(rs!Discount) = True Then
              txtcommission = ""
           Else
              txtcommission = rs!Discount
           End If
           
             
          txtproduct.SetFocus
          SendKeys "{home}+{end}"
         
End Sub
Private Sub txtCategory_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboClient.SetFocus
End Sub
Private Sub txtcommission_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub
Private Sub txtproduct_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCategory.SetFocus
End Sub


