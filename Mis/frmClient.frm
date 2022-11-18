VERSION 5.00
Begin VB.Form frmClient 
   BackColor       =   &H00FFFFC0&
   Caption         =   "ClientDetails"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   32
      Top             =   7200
      Width           =   7575
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Add"
         DisabledPicture =   "frmClient.frx":0000
         DownPicture     =   "frmClient.frx":0532
         Height          =   780
         Left            =   120
         MaskColor       =   &H008080FF&
         Picture         =   "frmClient.frx":0A64
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdMod 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Modify"
         Height          =   780
         Left            =   1320
         Picture         =   "frmClient.frx":0EA6
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFF80&
         Caption         =   "C&lear"
         Height          =   780
         Left            =   4920
         Picture         =   "frmClient.frx":12E8
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFF80&
         Caption         =   "<<&Back<<"
         Height          =   780
         Left            =   6120
         Picture         =   "frmClient.frx":13EA
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Delete"
         Height          =   780
         Left            =   2520
         Picture         =   "frmClient.frx":191C
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Preview"
         Height          =   780
         Left            =   3720
         Picture         =   "frmClient.frx":1D5E
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ListBox lstClient 
      BackColor       =   &H80000018&
      Height          =   7740
      Left            =   8400
      TabIndex        =   31
      Top             =   120
      Width           =   3375
   End
   Begin VB.Frame frmClient 
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
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.TextBox txtFax 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6360
         TabIndex        =   30
         Top             =   6120
         Width           =   1575
      End
      Begin VB.TextBox txtweb 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   29
         Top             =   6720
         Width           =   5415
      End
      Begin VB.TextBox txtOffTel 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   28
         Top             =   6120
         Width           =   2535
      End
      Begin VB.TextBox txtAreaCode 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   27
         Top             =   5520
         Width           =   2535
      End
      Begin VB.Frame FraAddress 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2175
         Left            =   240
         TabIndex        =   15
         Top             =   3240
         Width           =   6735
         Begin VB.TextBox txtPOBox 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2280
            TabIndex        =   18
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox txtcity 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2280
            TabIndex        =   17
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox txtCountry 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2280
            TabIndex        =   16
            Top             =   1680
            Width           =   3375
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "  P.O.Box"
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "  City"
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "  Country"
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   1680
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "                              Name             Mobile     e-mail Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2175
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   8055
         Begin VB.TextBox txtMobile1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4560
            ScrollBars      =   1  'Horizontal
            TabIndex        =   11
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtMobile2 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4560
            ScrollBars      =   1  'Horizontal
            TabIndex        =   10
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtMobile3 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4560
            ScrollBars      =   1  'Horizontal
            TabIndex        =   9
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtmail1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   5880
            ScrollBars      =   1  'Horizontal
            TabIndex        =   8
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtmail2 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   5880
            ScrollBars      =   1  'Horizontal
            TabIndex        =   7
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtmail3 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   5880
            ScrollBars      =   1  'Horizontal
            TabIndex        =   6
            Top             =   1560
            Width           =   2055
         End
         Begin VB.TextBox txtConName3 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2400
            TabIndex        =   5
            Top             =   1560
            Width           =   2055
         End
         Begin VB.TextBox txtConName2 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2400
            ScrollBars      =   1  'Horizontal
            TabIndex        =   4
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtConName1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2400
            ScrollBars      =   1  'Horizontal
            TabIndex        =   3
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFFFC0&
            Caption         =   "  Media Manager"
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "  Media Director"
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFC0&
            Caption         =   "  Managing Director"
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.TextBox txtClientName 
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
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "  Web Site"
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   480
         TabIndex        =   26
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "  Fax"
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   5160
         TabIndex        =   25
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Client Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   480
         TabIndex        =   24
         Top             =   360
         Width           =   2205
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Telephone (Off)"
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   480
         TabIndex        =   23
         Top             =   6120
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Area Code"
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   480
         TabIndex        =   22
         Top             =   5520
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim AgnNm As String
Dim ws As Workspace
Dim db As Database
Dim rs As Recordset

Private Sub Form_Unload(Cancel As Integer)
 Exit Sub
End Sub

Private Sub cmdadd_Click()

  If ValidateData = True Then
  
    Set ws = DBEngine.Workspaces(0)
    ' Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Select * from clientdtls where clientname='" & findfirstfixup(Trim(UCase(txtClientName))) & "' "
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
         MsgBox " Client already existing"
         Exit Sub
        Else
    Sqlqry1 = " Insert into clientdtls values('" & findfirstfixup(UCase(Trim(txtClientName))) & "','" _
              & Trim(txtPOBox) & "','" _
              & Trim(txtcity) & "','" _
              & Trim(txtCountry) & "','" _
              & Trim(txtOffTel) & "','" _
              & Trim(txtMobile1) & "','" _
              & Trim(txtMobile2) & "','" _
              & Trim(txtMobile3) & "','" _
              & Trim(txtFax) & "','" _
              & Trim(txtmail1) & "','" _
              & Trim(txtmail2) & "','" _
              & Trim(txtmail3) & "','" _
              & Trim(txtweb) & "','" _
              & Trim(txtConName1) & "','" _
              & Trim(txtConName2) & "','" _
              & Trim(txtConName3) & "','" _
              & Trim(txtAreaCode) & "')"
                ws.BeginTrans
                db.Execute (Sqlqry1)
                ws.CommitTrans
                
                 MsgBox "Record is inserted", vbDefaultButton3, "Status"
                 textclear
                 PopulateAgencycodes
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
txtClientName.Text = ""
txtOffTel = ""
txtMobile1 = ""
txtMobile2 = ""
txtMobile3 = ""
txtCountry = ""
txtFax = ""
txtPOBox = ""
txtcity = ""
txtmail1 = ""
txtmail2 = ""
txtweb = ""
txtConName1 = ""
txtConName2 = ""
txtConName3 = ""
txtAreaCode = ""
End Function

Private Function ValidateData()

ValidateData = False

If txtClientName.Text = "" Or IsNumeric(txtClientName) = True Then
   MsgBox "Invalid Agency Name", vbInformation, "Invalid Entry"
   txtClientName.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
   
ElseIf txtCountry.Text = "" Or IsNumeric(txtCountry.Text) = True Then
   MsgBox "Invalid Country", vbInformation, "Invalid Entry"
   txtCountry.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ElseIf txtcity.Text = "" Or IsNumeric(txtcity.Text) = True Then
   MsgBox "Invalid City", vbInformation, "Invalid Entry"
   txtcity.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
End If
ValidateData = True
End Function

Private Sub cmdDelete_Click()
Dim tempStr
If lstClient.SelCount = 0 Then
        MsgBox "Select the Client Name for Deletion.", vbInformation, "Selection Error"
        lstClient.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Delete the Client Name : " & txtClientName, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If DeleteData = False Then Exit Sub
        Else
            MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
            lstClient.SetFocus
            Exit Sub
        End If
End Sub

Private Sub cmdMod_Click()
Dim tempStr

    If lstClient.SelCount = 0 Then
        MsgBox "Select the Clinet Name for Modification.", vbInformation, "Selection Error"
        lstClient.SetFocus
        Exit Sub
    End If
        AgnNm = " "
        If ValidateData = False Then Exit Sub
        AgnNm = lstClient.Text
        tempStr = MsgBox("Do You Want To Modify the Client Details :" & lstClient.Text, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If ModifyData = False Then Exit Sub
        Else
              MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
              lstClient.SetFocus
              Exit Sub
        End If
    End Sub

Private Function ModifyData() As Boolean
    Dim i
    Dim AgnNm
    ModifyData = False
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    AgnNm = lstClient.Text
    i = Trim(lstClient.Text)
    
           
           Sqlqry = "Update clientdtls Set " _
                  & " clientname = '" & findfirstfixup(UCase(Trim(txtClientName.Text))) & "'," _
                  & " pobox = '" & Trim(txtPOBox.Text) & "'," _
                  & " city = '" & Trim(txtcity.Text) & "'," _
                  & " country = '" & Trim(txtCountry.Text) & "'," _
                  & " tel_off = '" & Trim(txtOffTel.Text) & "'," _
                  & " Mobile1 = '" & Trim(txtMobile1.Text) & "'," _
                  & " Mobile2 = '" & Trim(txtMobile2.Text) & "'," _
                  & " Mobile3 = '" & Trim(txtMobile3.Text) & "'," _
                  & " fax ='" & Trim(txtFax.Text) & "'," _
                  & " e_mail1 = '" & Trim(txtmail1.Text) & "'," _
                  & " e_mail2 = '" & Trim(txtmail2.Text) & "'," _
                  & " e_mail3 = '" & Trim(txtmail3.Text) & "'," _
                  & " web = '" & Trim(txtweb) & "', " _
                  & " name1 = '" & Trim(txtConName1) & "', " _
                  & " name2 = '" & Trim(txtConName2) & "', " _
                  & " name3 = '" & Trim(txtConName3) & "', " _
                  & " Area_code = '" & Trim(txtAreaCode) & "' " _
                  & " Where clientname ='" & i & "'"
               
                                                
                                                     
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        populateclientname
        
        MsgBox "Record Modified With " & Chr(10) & Chr(10) & _
               "Client Name = " & i, vbInformation, "Data Modified"
          
        If Trim(AgnNm) <> UCase(Trim(txtClientName)) Then
            populateclientname
        End If
        
        textclear
        PopulateAgencycodes
        ModifyData = True
        Exit Function
End Function
Private Sub populateclientname()
 Dim i, j
 i = UCase(Trim(txtClientName.Text))
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       j = Trim(lstClient.Text)
        Sqlqry = "Select Client from bo_mas  where Client ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update bo_mas set Client = '" & findfirstfixup(i) & "' WHERE Client='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
        Sqlqry = "Select Client from bo_tracin  where Client ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update bo_tracin set Client = '" & findfirstfixup(i) & "' WHERE Client='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
            
            
        Sqlqry = "Select Client from bo_tramag  where Client ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update bo_tramag set Client = '" & findfirstfixup(i) & "' WHERE Client='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
            
        Sqlqry = "Select Client from bo_tratv  where Client ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            Sqlqry1 = "Update bo_tratv set Client = '" & findfirstfixup(i) & "' WHERE Client='" & findfirstfixup(j) & "' "
            ws.BeginTrans
            db.Execute (Sqlqry1)
            ws.CommitTrans
        End If
            
        Sqlqry = "Select Client from bo_traol  where Client ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
            Sqlqry1 = "Update bo_traol set Client = '" & findfirstfixup(i) & "' WHERE Client='" & findfirstfixup(j) & "' "
            ws.BeginTrans
            db.Execute (Sqlqry1)
            ws.CommitTrans
         End If
         
        Sqlqry = "Select Client_name from products  where Client_name ='" & findfirstfixup(j) & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              Sqlqry1 = "Update products set Client_name = '" & findfirstfixup(i) & "' WHERE Client_name='" & findfirstfixup(j) & "' "
              ws.BeginTrans
              db.Execute (Sqlqry1)
              ws.CommitTrans
         End If
         
         
 
End Sub
Private Function DeleteData() As Boolean
  Dim i
    
    DeleteData = False
    
    i = Trim(UCase(txtClientName.Text))
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "select client from bo_mas where client='" & findfirstfixup(i) & "'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
              MsgBox "Transactions are recorded, cannot delete Client . . . "
              textclear
              PopulateAgencycodes
              Exit Function
        Else
    
    
            Sqlqry1 = "Delete * from clientdtls Where clientName = '" & i & "'"
                                                   
             ws.BeginTrans
             db.Execute (Sqlqry1)
             ws.CommitTrans
             MsgBox "Record Deleted With " & Chr(10) & Chr(10) & _
                    "Client Name = " & i, vbInformation, "Data Modified"
             textclear
             PopulateAgencycodes
        End If
               
End Function

Private Sub PopulateAgencycodes()
    textclear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from clientdtls Order by clientName"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
      rs.MoveFirst
        lstClient.Clear
        Do Until rs.EOF
            lstClient.AddItem rs!clientname
            rs.MoveNext
        Loop
    End If
        
End Sub

Private Sub CmdPrint_Click()
'CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
'CrystalReport1.ReportFileName = App.Path & "\AgencyList.rpt"
'CrystalReport1.WindowState = crptMaximized
'CrystalReport1.Action = 1
End Sub

Private Sub Form_Load()
    PopulateAgencycodes
    textclear
End Sub

Private Sub lstClient_Click()
Dim i
Dim tempBln As String
    If lstClient.ListIndex = -1 Then
        tempBln = False
    Else
        tempBln = True
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = findfirstfixup(Trim(lstClient.Text))
        Sqlqry = "Select * from clientdtls Where clientname= '" & i & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MsgBox "Data Mis Matching", vbInformation, "Deleted Status"
            Exit Sub
         Else
           txtClientName = rs!clientname
           
           If IsNull(rs!pobox) = True Then
              txtPOBox = ""
           Else
              txtPOBox = rs!pobox
           End If
                               
           If IsNull(rs!city) = True Then
              txtcity = ""
           Else
              txtcity = rs!city
           End If
           
           If IsNull(rs!country) = True Then
              txtCountry = ""
           Else
              txtCountry = rs!country
           End If
           
          If IsNull(rs!tel_off) = True Then
              txtOffTel = ""
           Else
              txtOffTel = rs!tel_off
          End If
          
          
          If IsNull(rs!mobile1) = True Then
              txtMobile1 = ""
           Else
              txtMobile1 = rs!mobile1
          End If
          
          If IsNull(rs!mobile2) = True Then
              txtMobile2 = ""
           Else
              txtMobile2 = rs!mobile2
          End If
          
          If IsNull(rs!mobile3) = True Then
              txtMobile3 = ""
           Else
              txtMobile3 = rs!mobile3
          End If
          
          If IsNull(rs!fax) = True Then
              txtFax = ""
           Else
              txtFax = rs!fax
          End If
      
          
          If IsNull(rs!E_mail1) = True Then
              txtmail1 = ""
           Else
              txtmail1 = rs!E_mail1
          End If
      
          If IsNull(rs!E_mail2) = True Then
              txtmail2 = ""
           Else
              txtmail2 = rs!E_mail2
          End If
      
          If IsNull(rs!E_mail3) = True Then
              txtmail3 = ""
           Else
              txtmail3 = rs!E_mail3
          End If
      
          If IsNull(rs!web) = True Then
              txtweb = ""
           Else
              txtweb = rs!web
          End If
             
          If IsNull(rs!name1) = True Then
              txtConName1 = ""
           Else
              txtConName1 = rs!name1
          End If
          
          If IsNull(rs!name2) = True Then
              txtConName2 = ""
           Else
              txtConName2 = rs!name2
          End If
          
          If IsNull(rs!name3) = True Then
              txtConName3 = ""
           Else
              txtConName3 = rs!name3
          End If
          
          If IsNull(rs!Area_Code) = True Then
              txtAreaCode = ""
           Else
              txtAreaCode = rs!Area_Code
          End If
          
       End If
    
End Sub
Private Sub txtClientName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtConName1.SetFocus
End Sub

Private Sub txtAreaCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtOffTel.SetFocus
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCountry.SetFocus
End Sub
Private Sub txtConName1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtMobile1.SetFocus
End Sub
Private Sub txtConName2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtMobile2.SetFocus
End Sub
Private Sub txtConName3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtMobile3.SetFocus
End Sub
Private Sub txtCountry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAreaCode.SetFocus
End Sub
Private Sub txtFax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtweb.SetFocus
End Sub
Private Sub txtmail1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtConName2.SetFocus
End Sub
Private Sub txtmail2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtConName3.SetFocus
End Sub
Private Sub txtmail3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPOBox.SetFocus
End Sub
Private Sub txtMobile1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmail1.SetFocus
End Sub
Private Sub txtMobile2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmail2.SetFocus
End Sub
Private Sub txtMobile3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtmail3.SetFocus
End Sub
Private Sub txtOffTel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtFax.SetFocus
End Sub
Private Sub txtPOBox_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtcity.SetFocus
End Sub

Private Sub txtweb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub

Function replacestr(Textin, ByVal searchstr As String, _
                    ByVal Replacement As String, _
                    ByVal CompMode As Integer)

  Dim Worktext As String, Pointer As Integer
   If IsNull(Textin) Then
    replacestr = Null
   Else
    Worktext = Textin
    Pointer = InStr(1, Worktext, searchstr, CompMode)
     Do While Pointer > 0
      Worktext = Left(Worktext, Pointer - 1) & Replacement & _
                 Mid(Worktext, Pointer + Len(searchstr))
                 
      Pointer = InStr(Pointer + Len(Replacement), Worktext, _
                 searchstr, CompMode)
                 
    Loop
    
    replacestr = Worktext
    
  
   End If
End Function

Function sqlfixup(Textin)
 sqlfixup = replacestr(Textin, "'", "''", 0)
End Function
Function jetsqlfixup(Textin)
 Dim Temp
  Temp = replacestr(Textin, "'", "''", 0)
  jetsqlfixup = replacestr(Temp, "|", "' & Chr(124) & '", 0)
End Function
 
Function findfirstfixup(Textin)
  Dim Temp
  Temp = replacestr(Textin, "'", "' & Chr(39) & '", 0)
  findfirstfixup = replacestr(Temp, "|", "' & Chr(124) & '", 0)

End Function


