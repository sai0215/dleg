VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmClientToorg 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   8595
   ClientLeft      =   30
   ClientTop       =   345
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmclienttoorg.frx":0000
      Left            =   840
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "                                       Turnover / Client                                   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   8055
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   11295
      Begin VB.ComboBox cboCurrency 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   4965
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3120
         Width           =   1095
      End
      Begin VB.ComboBox cboProduct 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   390
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   5040
         Width           =   5415
      End
      Begin VB.ComboBox cboMediaType 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   390
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   4080
         Width           =   5415
      End
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00C0FFC0&
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
         Height          =   975
         Left            =   3480
         Picture         =   "frmclienttoorg.frx":001D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6960
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00C0FFC0&
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
         Height          =   975
         Left            =   4920
         Picture         =   "frmclienttoorg.frx":045F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6960
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00C0FFC0&
         Caption         =   "<<&Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6240
         Picture         =   "frmclienttoorg.frx":08A1
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6960
         Width           =   1455
      End
      Begin VB.ComboBox CboClient 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   390
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   6000
         Width           =   5415
      End
      Begin VB.ComboBox cbomonthTo 
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
         ForeColor       =   &H000040C0&
         Height          =   420
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2400
         Width           =   2295
      End
      Begin VB.ComboBox cbomonthfrom 
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
         ForeColor       =   &H000040C0&
         Height          =   420
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox cboyear 
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
         ForeColor       =   &H000040C0&
         Height          =   420
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Currency"
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
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblSubMediaName 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   6480
         TabIndex        =   18
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblMedianame 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   11280
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Client"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Month To"
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
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Month From"
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
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "  Year"
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
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmClientToorg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Dim ws As Workspace
Dim db As Database
Dim i As Integer
Dim X, y, Z As Integer
Dim adddisc As Currency
Dim scharge As Currency
Dim ntra As Currency
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset

Private Sub cboCurrency_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboMediatype.SetFocus
End Sub

Private Sub cboMediaType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboProduct.SetFocus
End Sub

Private Sub cboMediaType_LostFocus()
If Mid(cboMediatype.Text, 1, 3) = "Cin" Then
   lblMedianame.Caption = "Cinema"
   lblSubMediaName.Caption = Trim(Mid(cboMediatype, 8, 30))
ElseIf Mid(cboMediatype.Text, 1, 3) = "Mag" Then
   lblMedianame.Caption = "Magazine"
   lblSubMediaName.Caption = Trim(Mid(cboMediatype, 10, 30))
ElseIf Mid(cboMediatype.Text, 1, 3) = "Onl" Then
   lblMedianame.Caption = "Online"
   lblSubMediaName.Caption = Trim(Mid(cboMediatype, 8, 30))
ElseIf Mid(cboMediatype.Text, 1, 3) = "Tel" Then
   lblMedianame.Caption = "Television"
   lblSubMediaName.Caption = Trim(Mid(cboMediatype, 12, 30))
End If
cboProduct.SetFocus
   
End Sub

Private Sub cbomonthfrom_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbomonthTo.SetFocus
End Sub


Private Sub cbomonthTo_GotFocus()
cbomonthTo.Clear
If cbomonthfrom.ListIndex = 0 Then
    cbomonthTo.AddItem "January"
    cbomonthTo.AddItem "February"
    cbomonthTo.AddItem "March"
    cbomonthTo.AddItem "April"
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 1 Then
    cbomonthTo.AddItem "February"
    cbomonthTo.AddItem "March"
    cbomonthTo.AddItem "April"
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 2 Then
    cbomonthTo.AddItem "March"
    cbomonthTo.AddItem "April"
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 3 Then
    cbomonthTo.AddItem "April"
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 4 Then
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 5 Then
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 6 Then
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 7 Then
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 8 Then
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 9 Then
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 10 Then
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 11 Then
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
Else
    cbomonthTo.SetFocus
End If
End Sub
Private Sub cbomonthTo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboCurrency.SetFocus
End Sub
Private Sub cboProduct_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboClient.SetFocus
End Sub
Private Sub cboyear_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbomonthfrom.SetFocus
End Sub
Private Sub cmdBack_Click()
 Unload Me
End Sub
Private Sub cmdClear_Click()
 textclear
End Sub

 Private Sub cmdDisplay_Click()
  Dim m, n, o, p As String
  
  If ValidateData = True Then
              
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Delete * from TO_Agency"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
         n = ""
         m = ""
         o = ""
         p = ""
         
       n = Trim(lblMedianame.Caption)
       m = Trim(lblSubMediaName.Caption)
       
       If CboClient.Text <> "All" Then o = Trim(CboClient.Text)
       If cboProduct.Text <> "All" Then p = Trim(cboProduct.Text)
             
       If CboClient.Text = "All" And cboProduct.Text = "All" And cboMediatype.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & ""
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
                
        ElseIf CboClient.Text = "All" And cboProduct.Text = "All" And cboMediatype.Text = "Cinema" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediatype.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
         ' Test this
         ElseIf CboClient.Text = "All" And cboProduct.Text = "All" And cboMediatype.Text = "Magazine" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' and media = '" & Trim(cboMediatype.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
            ElseIf CboClient.Text = "All" And cboProduct.Text = "All" And cboMediatype.Text = "Online" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediatype.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
            ElseIf CboClient.Text = "All" And cboProduct.Text = "All" And cboMediatype.Text = "Television" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediatype.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
               ElseIf CboClient.Text = "All" And cboProduct.Text = "All" And cboMediatype.Text = n & " " & m Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(lblMedianame.Caption) & "' and sub_media='" & Trim(lblSubMediaName.Caption) & " ' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                ElseIf CboClient.Text = "All" And cboProduct.Text = p And cboMediatype.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND Product='" & Trim(cboProduct.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & ""
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                ElseIf CboClient.Text = "All" And cboProduct.Text = p And cboMediatype.Text = "Cinema" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and product='" & Trim(cboProduct.Text) & "' and Media='" & Trim(cboMediatype.Text) & "' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = "All" And cboProduct.Text = p And cboMediatype.Text = "Magazine" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and product='" & Trim(cboProduct.Text) & "' and Media='" & Trim(cboMediatype.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
            '***** stoop
                ElseIf CboClient.Text = "All" And cboProduct.Text = p And cboMediatype.Text = "Television" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and product='" & Trim(cboProduct.Text) & "' and Media='" & Trim(cboMediatype.Text) & "' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = "All" And cboProduct.Text = p And cboMediatype.Text = "Online" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and product='" & Trim(cboProduct.Text) & "' and Media='" & Trim(cboMediatype.Text) & "' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                
                ElseIf CboClient.Text = "All" And cboProduct.Text = p And cboMediatype.Text = n & " " & m Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media=' n ' and sub_media=' m ' and Product ='" & Trim(cboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And cboProduct.Text = p And cboMediatype.Text = n & " " & m Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='n' and sub_media='m' and Agency='" & Trim(CboClient.Text) & "' and Product ='" & Trim(cboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And cboProduct.Text = "All" And cboMediatype.Text = n & " " & m Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(lblMedianame.Caption) & "' and sub_media='" & Trim(lblSubMediaName.Caption) & "' and Agency='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And cboProduct.Text = "All" And cboMediatype.Text = "Cinema" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediatype.Text) & "' and Agency='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                ElseIf CboClient.Text = o And cboProduct.Text = "All" And cboMediatype.Text = "Magazine" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediatype.Text) & "' and Agency='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And cboProduct.Text = "All" And cboMediatype.Text = "Online" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediatype.Text) & "' and Agency='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And cboProduct.Text = "All" And cboMediatype.Text = "Television" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediatype.Text) & "' and Agency='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And cboProduct.Text = "All" And cboMediatype.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Agency='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And cboProduct.Text = p And cboMediatype.Text = "Cinema" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediatype.Text) & "' and Agency='" & Trim(CboClient.Text) & "' and Product ='" & Trim(cboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And cboProduct.Text = p And cboMediatype.Text = "Magazine" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediatype.Text) & "' and Agency='" & Trim(CboClient.Text) & "' and Product ='" & Trim(cboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And cboProduct.Text = p And cboMediatype.Text = "Online" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediatype.Text) & "' and Agency='" & Trim(CboClient.Text) & "' and Product ='" & Trim(cboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And cboProduct.Text = p And cboMediatype.Text = "Television" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediatype.Text) & "' and Agency='" & Trim(CboClient.Text) & "' and Product='" & Trim(cboProduct.Text) & "' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And cboProduct.Text = p And cboMediatype.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Agency='" & Trim(CboClient.Text) & "' and Product ='" & Trim(cboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
              End If
              
 Cinadjustments
 curadjustments
     
                  
     
    If Mid(cboMediatype.Text, 1, 3) = "Mag" Then
              
            With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Toclientmag.rpt"
                .Formulas(0) = "yyy='" & Val(cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "mty='" & Trim(cboMediatype.Text) & "'"
                .Formulas(3) = "prd='" & Trim(cboProduct.Text) & "'"
                .Formulas(4) = "agn='" & Trim(CboClient.Text) & "'"
                .Formulas(5) = "Cur='" & Trim(cboCurrency.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
       Else
              With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Toclient.rpt"
                .Formulas(0) = "yyy='" & Val(cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "mty='" & Trim(cboMediatype.Text) & "'"
                .Formulas(3) = "prd='" & Trim(cboProduct.Text) & "'"
                .Formulas(4) = "agn='" & Trim(CboClient.Text) & "'"
                .Formulas(5) = "Cur='" & Trim(cboCurrency.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
            
      End If
            
   Else
        MsgBox "Improper Dates", vbDefaultButton1, "Invalid entry"
        Exit Sub
  End If
 End Sub
Private Sub Cinadjustments()
Sqlqry = " Delete * from TO_Agency1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
              
        Sqlqry = "Select * from To_agency where Media='Cinema'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency1 values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
           
     Sqlqry = " Delete * from TO_Agency where media='Cinema'"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
              
        adddisc = 0
        scharge = 0
        ntra = 0
                
                       
              
      Sqlqry = "Select * from To_agency1"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
              
                adddisc = rs!add_discount
                scharge = rs!surcharge
                Do Until rs.EOF
                Sqlqry1 = "Select * from bo_tracin where serial_no='" & Trim(rs!serial_no) & "' "
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                 If rs1.RecordCount <> 0 Then
                  rs1.MoveFirst
                   Do Until rs1.EOF
                     If rs1!Type = "Paid" Then
                        If rs!add_discount = 0 Then
                          ntra = Val(rs1!tra_amount) - (Val(rs1!tra_amount) * rs!disc_rate / 100) - (((rs1!tra_amount) - (rs1!tra_amount * rs!disc_rate / 100)) * rs!disc_percentage / 100)
                        Else
                          ntra = Val(rs1!tra_amount) - (Val(rs1!tra_amount) * rs!disc_rate / 100) - (((rs1!tra_amount) - (rs1!tra_amount * rs!disc_rate / 100)) * rs!disc_percentage / 100) - Val(rs!add_discout)
                        End If
                         
                        If rs1!tcurrency = "USD" Then
                            Sqlqry2 = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                 & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs1!tcurrency & "'," & rs1!tra_amount & "," & ntra & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                                 & rs1!sub_media & "','" _
                                                 & findfirstfixup(rs!bo_ref) & "'," _
                                                 & Val(rs1!amount) & "," & 0 & "," & 0 & "," & rs!disc_percentage & "," & scharge & ", '" _
                                                 & Trim(rs!disc_rate) & "'," _
                                                 & adddisc & "," _
                                                 & ntra & ")"
                         Else
                            Sqlqry2 = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                 & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs1!tcurrency & "'," & rs1!tra_amount & "," & ntra & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                                 & rs1!sub_media & "','" _
                                                 & findfirstfixup(rs!bo_ref) & "'," _
                                                 & Val(rs1!amount) & "," & 0 & "," & 0 & "," & rs!disc_percentage & "," & scharge & ", '" _
                                                 & Trim(rs!disc_rate) & "'," _
                                                 & adddisc & "," _
                                                 & ntra * convertion & ")"
                           End If

                          adddisc = 0
                          scharge = 0
                                             
                      ElseIf rs1!Type = "Free" Then
                          Sqlqry2 = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs1!tcurrency & "'," & rs1!tra_amount & "," & 0 & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs1!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs1!amount) & "," & Val(rs1!amount) & "," & 0 & "," & 0 & "," & 0 & ", '" _
                                             & 0 & "'," _
                                             & 0 & "," & 0 & ")"
                       Else
                             Sqlqry2 = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs1!tcurrency & "'," & rs1!tra_amount & "," & 0 & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs1!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs1!amount) & "," & 0 & "," & Val(rs1!amount) & "," & 0 & "," & 0 & ", '" _
                                             & 0 & "'," _
                                             & 0 & "," & 0 & ")"
                       End If
                     
                                ws.BeginTrans
                                db.Execute (Sqlqry2)
                                ws.CommitTrans
                  rs1.MoveNext
                  Loop
                 End If
               rs.MoveNext
               Loop
           End If
     


End Sub
Private Sub Form_Load()
cbomonthfrom.AddItem "January"
cbomonthfrom.AddItem "February"
cbomonthfrom.AddItem "March"
cbomonthfrom.AddItem "April"
cbomonthfrom.AddItem "May"
cbomonthfrom.AddItem "June"
cbomonthfrom.AddItem "July"
cbomonthfrom.AddItem "August"
cbomonthfrom.AddItem "September"
cbomonthfrom.AddItem "October"
cbomonthfrom.AddItem "November"
cbomonthfrom.AddItem "December"

lblMedianame.Caption = ""
lblSubMediaName.Caption = ""

cboCurrency.AddItem "DHS"
cboCurrency.AddItem "USD"
 

i = 2000

For i = 2000 To 2100
 cboyear.AddItem i
Next
X = 0

 cboyear.Text = Year(Now())
 
 X = Month(Now())
  
If X = 1 Then
   cbomonthfrom.ListIndex = 0
ElseIf X = 2 Then
   cbomonthfrom.ListIndex = 1
ElseIf X = 3 Then
   cbomonthfrom.ListIndex = 2
ElseIf X = 4 Then
   cbomonthfrom.ListIndex = 3
ElseIf X = 5 Then
   cbomonthfrom.ListIndex = 4
ElseIf X = 6 Then
   cbomonthfrom.ListIndex = 5
ElseIf X = 7 Then
   cbomonthfrom.ListIndex = 6
ElseIf X = 8 Then
   cbomonthfrom.ListIndex = 7
ElseIf X = 9 Then
   cbomonthfrom.ListIndex = 8
ElseIf X = 10 Then
   cbomonthfrom.ListIndex = 9
ElseIf X = 11 Then
   cbomonthfrom.ListIndex = 10
Else
   cbomonthfrom.ListIndex = 11
End If

PopulateAgencycodes
populateMedia
populateproducts

End Sub
Private Sub populateproducts()
    cboProduct.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from products Order by Product_Name"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
           cboProduct.AddItem "All"
        rs.MoveFirst
            Do Until rs.EOF
              cboProduct.AddItem rs!product_name
            rs.MoveNext
       Loop
    End If
 End Sub

Private Sub populateMedia()
    cboMediatype.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Media Order by Media_Type"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        cboMediatype.AddItem "All"
        cboMediatype.AddItem "Cinema"
        cboMediatype.AddItem "Magazine"
        cboMediatype.AddItem "Online"
        cboMediatype.AddItem "Television"
        rs.MoveFirst
            Do Until rs.EOF
              cboMediatype.AddItem rs!media_type & " " & rs!sub_media
            rs.MoveNext
       Loop
    End If
 End Sub
 
Private Sub curadjustments()

     Sqlqry = " Delete * from TO_Agency1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
              
     If cboCurrency.Text = "USD" Then
       
        Sqlqry = "Select * from To_agency where Tcurrency='DHS'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency1 values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & Round(rs!tra_gamount / convertion, 2) & "," & Round(rs!tra_namount / convertion, 2) & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Round(Val(rs!gross_amount) / convertion, 2) & "," _
                                             & Round(Val(rs!Tot_free) / convertion, 2) & "," _
                                             & Round(Val(rs!Tot_barter) / convertion, 2) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Round(Val(rs!surcharge) / convertion, 2) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Round(Val(rs!add_discount) / convertion, 2) & "," _
                                             & Round(Val(rs!net_amount) / convertion, 2) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
           
            Sqlqry = " Delete * from TO_Agency where Tcurrency='DHS'"
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
                     
            Sqlqry = "Select * from To_agency1 where Tcurrency='DHS'"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & Val(rs!tra_gamount) & "," & Val(rs!tra_namount) & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
           
     Else
           Sqlqry = "Select * from To_agency where Tcurrency='USD'"
           Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     If Mid(rs!media, 1, 3) = "Cin" Then
                      Sqlqry = " Insert into TO_Agency1 values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!gross_amount & "," & rs!tra_namount * convertion & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Round(Val(rs!Tot_free) * convertion, 2) & "," _
                                             & Round(Val(rs!Tot_barter) * convertion, 2) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Round(Val(rs!surcharge) * convertion, 2) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Round(Val(rs!add_discount) * convertion, 2) & "," _
                                             & Val(rs!net_amount) & ")"
                        Else
                          Sqlqry = " Insert into TO_Agency1 values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & rs!gross_amount & "," & rs!net_amount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Round(Val(rs!Tot_free) * convertion, 2) & "," _
                                             & Round(Val(rs!Tot_barter) * convertion, 2) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Round(Val(rs!surcharge) * convertion, 2) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Round(Val(rs!add_discount) * convertion, 2) & "," _
                                             & Val(rs!net_amount) & ")"
                      End If
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
           
            Sqlqry = " Delete * from TO_Agency where Tcurrency='USD'"
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
                     
            Sqlqry = "Select * from To_agency1 where Tcurrency='USD'"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into TO_Agency values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & rs!tcurrency & "'," & Val(rs!tra_gamount) & "," & Val(rs!tra_namount) & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!agency) & "','" & rs!media & "','" _
                                             & rs!sub_media & "','" _
                                             & findfirstfixup(rs!bo_ref) & "'," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!net_amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
       End If
   
End Sub

Private Sub PopulateAgencycodes()
    CboClient.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from agndtls Order by AgentName"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
      rs.MoveFirst
        CboClient.Clear
         CboClient.AddItem "All"
        Do Until rs.EOF
            CboClient.AddItem rs!agentname
            rs.MoveNext
        Loop
    End If
        
End Sub

Private Function ValidateData()

ValidateData = False
If cboyear.Text = "" Then
   MsgBox "Invalid year", vbInformation, "Invalid Entry"
   cboyear.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
 ElseIf cbomonthfrom.Text = "" Then
   MsgBox "Select Month From", vbInformation, "Invalid Entry"
   cbomonthfrom.SetFocus
   SendKeys " {Home} + {end} "
   Exit Function
 ElseIf cbomonthTo.Text = "" Then
   MsgBox "Select Month To", vbInformation, "Invalid Entry"
   cbomonthTo.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
 ElseIf CboClient.Text = "" Then
   MsgBox "Select Agency", vbInformation, "Invalid Entry"
   CboClient.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
 ElseIf cboMediatype.Text = "" Then
   MsgBox "Select Media Type", vbInformation, "Invalid Entry"
   cboMediatype.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
 ElseIf cboProduct.Text = "" Then
   MsgBox "Select Product", vbInformation, "Invalid Entry"
   cboProduct.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
  ValidateData = True
End If
End Function

Private Sub textclear()
 CboClient.ListIndex = -1
 cboProduct.ListIndex = -1
 cboMediatype.ListIndex = -1
 cboyear.ListIndex = -1
 cbomonthfrom.ListIndex = -1
 cbomonthTo.ListIndex = -1
End Sub



