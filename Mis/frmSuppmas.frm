VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmSuppmas 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Supplier Entry, Modification & Deletion"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   11835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11835
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
      Height          =   8295
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   11415
      Begin VB.Frame fraSupplier 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Supplier Register"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   6375
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   11055
         Begin VB.TextBox Txtmaterial 
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
            Height          =   375
            Left            =   1560
            ScrollBars      =   1  'Horizontal
            TabIndex        =   11
            Top             =   5760
            Width           =   4215
         End
         Begin VB.TextBox txtCity 
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
            Height          =   405
            Left            =   1560
            MaxLength       =   30
            ScrollBars      =   1  'Horizontal
            TabIndex        =   4
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox txtTelephone 
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
            Height          =   375
            Left            =   1560
            MaxLength       =   20
            ScrollBars      =   1  'Horizontal
            TabIndex        =   6
            Top             =   3600
            Width           =   1575
         End
         Begin VB.TextBox txtCountry 
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
            Height          =   405
            Left            =   4200
            MaxLength       =   30
            ScrollBars      =   1  'Horizontal
            TabIndex        =   5
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox txtFax 
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
            Height          =   375
            Left            =   4200
            MaxLength       =   20
            ScrollBars      =   1  'Horizontal
            TabIndex        =   7
            Top             =   3600
            Width           =   1575
         End
         Begin VB.TextBox txtCrLimit 
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
            Height          =   375
            Left            =   4200
            TabIndex        =   9
            Top             =   4320
            Width           =   1575
         End
         Begin VB.TextBox txtCrDays 
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
            Height          =   405
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   8
            Top             =   4320
            Width           =   615
         End
         Begin VB.TextBox txtConPerson 
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
            Height          =   375
            Left            =   1560
            MaxLength       =   35
            ScrollBars      =   1  'Horizontal
            TabIndex        =   10
            Top             =   5040
            Width           =   4215
         End
         Begin VB.TextBox txtaddress 
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
            Height          =   375
            Left            =   1560
            MaxLength       =   200
            ScrollBars      =   1  'Horizontal
            TabIndex        =   3
            Top             =   2160
            Width           =   4215
         End
         Begin VB.ListBox lstSuppcodes 
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
            ForeColor       =   &H00404040&
            Height          =   4260
            Left            =   6120
            TabIndex        =   0
            Top             =   360
            Width           =   4695
         End
         Begin VB.TextBox txtName 
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
            Height          =   375
            Left            =   1560
            MaxLength       =   35
            ScrollBars      =   1  'Horizontal
            TabIndex        =   2
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox txtcode 
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
            Height          =   375
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   1
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtOpBal 
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
            Height          =   375
            Left            =   8520
            MaxLength       =   20
            ScrollBars      =   1  'Horizontal
            TabIndex        =   12
            Top             =   5040
            Width           =   1575
         End
         Begin VB.TextBox txtClBal 
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
            Height          =   375
            Left            =   8520
            MaxLength       =   20
            ScrollBars      =   1  'Horizontal
            TabIndex        =   13
            Top             =   5640
            Width           =   1575
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Remarks"
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
            TabIndex        =   34
            Top             =   5880
            Width           =   1215
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Credit Limit"
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
            Left            =   3120
            TabIndex        =   33
            Top             =   4440
            Width           =   960
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Credit Days"
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
            Left            =   480
            TabIndex        =   32
            Top             =   4440
            Width           =   990
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Fax"
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
            Left            =   3720
            TabIndex        =   31
            Top             =   3720
            Width           =   315
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Telephone"
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
            Left            =   600
            TabIndex        =   30
            Top             =   3720
            Width           =   915
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Country"
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
            Left            =   3360
            TabIndex        =   29
            Top             =   3000
            Width           =   660
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "City"
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
            Left            =   1080
            TabIndex        =   28
            Top             =   3000
            Width           =   330
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Contact Person"
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
            Left            =   120
            TabIndex        =   27
            Top             =   5160
            Width           =   1320
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Address"
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
            Left            =   720
            TabIndex        =   26
            Top             =   2280
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Closing Balance"
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
            Left            =   6960
            TabIndex        =   25
            Top             =   5760
            Width           =   1380
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Opening Balance"
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
            Left            =   6960
            TabIndex        =   24
            Top             =   5160
            Width           =   1470
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Name"
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
            Left            =   960
            TabIndex        =   23
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Code"
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
            Left            =   960
            TabIndex        =   22
            Top             =   840
            Width           =   450
         End
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5325
         Picture         =   "frmSuppmas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   7200
         Width           =   1095
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00FFFF80&
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
         Height          =   855
         Left            =   4200
         Picture         =   "frmSuppmas.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7200
         Width           =   1095
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFF80&
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
         Height          =   855
         Left            =   7515
         Picture         =   "frmSuppmas.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   7200
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFF80&
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
         Height          =   855
         Left            =   6420
         Picture         =   "frmSuppmas.frx":0986
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   7200
         Width           =   1095
      End
      Begin VB.CommandButton cmdMod 
         BackColor       =   &H00FFFF80&
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
         Height          =   855
         Left            =   3135
         Picture         =   "frmSuppmas.frx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2040
         Picture         =   "frmSuppmas.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7200
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   7800
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
   End
End
Attribute VB_Name = "frmSuppmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim ws As Workspace
Dim db As Database
Dim rs As Recordset
Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub
Private Sub cmdadd_Click()
  
   Dim ws1 As Workspace
   Dim db1 As Database
   Dim rs1 As Recordset
   Dim Sql As String
   
   
   
   If ValidateData = True Then
  
    Set ws1 = DBEngine.Workspaces(0)
    Set db1 = ws1.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Supp_fin where Supp_No='" & Trim(txtcode) & "' "
    
    Set rs1 = db1.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs1.RecordCount <> 0 Then
        MsgBox " Record already exists"
        Exit Sub
        Else
         Sql = "Insert into SUPP_FIN values('" & Trim(txtcode) & "','" & _
                findfirstfixup(Trim(txtName)) & "','" & _
                Trim(txtOpbal) & "','" & _
                Trim(txtClbal) & "','" & _
                findfirstfixup(Trim(txtaddress)) & "','" & _
                findfirstfixup(Trim(txtConPerson)) & "','" & _
                findfirstfixup(Trim(Txtmaterial)) & "','" & _
                findfirstfixup(Trim(txtcity)) & "','" & _
                findfirstfixup(Trim(txtCountry)) & "','" & _
                Trim(txtTelephone) & "','" & _
                Trim(txtFax) & "','" & _
                Trim(txtCrDays) & "','" & _
                Trim(txtCrLimit) & "')"
               
               ws1.BeginTrans
               db1.Execute (Sql)
               ws1.CommitTrans
                
                MsgBox "Record is inserted", vbDefaultButton3, "Status"
                textclear
                PopulateSuppcodes
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
txtcode = ""
txtName = ""
txtaddress = ""
txtcity = ""
txtCountry = ""
txtTelephone = ""
txtFax = ""
txtCrDays = ""
txtCrLimit = ""
txtConPerson = ""
Txtmaterial = ""
txtOpbal = ""
txtClbal = ""
End Function

Private Function ValidateData()

ValidateData = False
If txtcode.Text = "" Then
   MsgBox "Invalid Supplier Code", vbInformation, "Invalid Entry"
   txtcode.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf txtName.Text = "" Or IsNumeric(txtName) = True Then
   MsgBox "Invalid Supplier Name", vbInformation, "Invalid Entry"
   txtName.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsNumeric(txtCountry) = True Then
   MsgBox "Invalid Country", vbInformation, "Invalid Entry"
   txtCountry.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsNumeric(txtConPerson) = True Then
   MsgBox "Invalid Contact Person Name", vbInformation, "Invalid Entry"
   txtConPerson.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsNumeric(txtcity.Text) = True Then
   MsgBox "Invalid City Name", vbInformation, "Invalid Entry"
   txtcity.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsNumeric(txtOpbal) = False Then
   MsgBox "Invalid Opening Balance", vbInformation, "Invalid Entry"
   txtOpbal.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsNumeric(txtClbal) = False Then
   MsgBox "Invalid Closing Balance", vbInformation, "Invalid Entry"
   txtClbal.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsNumeric(Txtmaterial) = True Then
   MsgBox "Invalid Available Material From Supplier", vbInformation, "Invalid Entry"
   Txtmaterial.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
  
ValidateData = True
End If
ValidateData = True
End Function

Private Sub cmdDelete_Click()
    If lstSuppcodes.SelCount = 0 Then
        MsgBox "Select the Supplier Code for Deletion.", vbInformation, "Selection Error"
        lstSuppcodes.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Delete the Supplier Code : " & txtcode, vbQuestion + vbYesNo, "Confirmation")
        If tempStr = vbYes Then
            If DeleteData = False Then Exit Sub
        Else
            MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
            txtcode.SetFocus
            Exit Sub
        End If
End Sub
Private Sub cmdMod_Click()
  Z = 1
    If lstSuppcodes.SelCount = 0 Then
       MsgBox "Select the Supplier Code for Modification.", vbInformation, "Selection Error"
       lstSuppcodes.SetFocus
       Exit Sub
    End If
       If ValidateData = False Then Exit Sub
       tempStr = MsgBox("Do You Want To Modify the Supplier Code :" & txtcode, vbQuestion + vbYesNo, "Confirmation")
       If tempStr = vbYes Then
          If ModifyData = False Then Exit Sub
       Else
           MsgBox "No Entries Recorded.", vbInformation, "Modify Status"
           txtcode.SetFocus
           Exit Sub
       End If
End Sub

Private Function ModifyData() As Boolean
    Dim i
    ModifyData = False
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     i = Trim(txtcode.Text)
                
         Sqlqry = "Update Supp_fin Set " _
                  & " Supp_name = '" & findfirstfixup(Trim(txtName.Text)) & "'," _
                  & " Telephone = '" & Trim(txtTelephone.Text) & "'," _
                  & " Fax = '" & Trim(txtFax.Text) & "'," _
                  & " Address = '" & findfirstfixup(Trim(txtaddress.Text)) & "'," _
                  & " Con_Person = '" & findfirstfixup(Trim(txtConPerson.Text)) & "'," _
                  & " Avail_material = '" & findfirstfixup(Trim(Txtmaterial.Text)) & "'," _
                  & " City = '" & findfirstfixup(Trim(txtcity.Text)) & "'," _
                  & " Country = '" & findfirstfixup(Trim(txtCountry.Text)) & "'," _
                  & " Crdt_days=" & Val(txtCrDays.Text) & "," _
                  & " Crdt_lmt=" & Val(txtCrLimit.Text) & "," _
                  & " OPEN_BAL =" & Val(txtOpbal.Text) & "" _
                  & " Where Supp_No ='" & Trim(txtcode.Text) & "'"
                                                                                                  
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        MsgBox "Record Modified With " & Chr(10) & Chr(10) & _
               "Supplier Code = " & i, vbInformation, "Data Modified"
        textclear
        PopulateSuppcodes
        tempBln = False
        ModifyData = True
        Exit Function
End Function

Private Function DeleteData() As Boolean
Dim i
    
    DeleteData = False
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       
    i = Trim(txtcode.Text)
        
       If txtOpbal > 0 Or txtClbal > 0 Then
         MsgBox " Supplier Cannot be Deleted since the transactions are recorded"
         DeleteData = False
       Else
       Sqlqry = "Delete * from Supp_fin Where Supp_No = '" & i & "'"
                                         
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        MsgBox "Record Deleted With " & Chr(10) & Chr(10) & _
               "Supplier Code = " & i, vbInformation, "Data Modified"
        textclear
        PopulateSuppcodes
        tempBln = False
        If Validate1 = False Then Exit Function
        DeleteData = True
        Exit Function
        End If
       
End Function

Private Sub PopulateSuppcodes()
       
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from SUPP_FIN Order by Supp_No"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
    rs.MoveFirst
        lstSuppcodes.Clear
        Do Until rs.EOF
            lstSuppcodes.AddItem rs!Supp_no & "    :    " & rs!Supp_name
            rs.MoveNext
        Loop
    End If
       
End Sub

Private Sub CmdPrint_Click()
CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
CrystalReport1.ReportFileName = App.Path & "\Supp_fin.rpt"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
End Sub

Private Sub Form_Load()
    tempBln = False
    PopulateSuppcodes
    textclear
End Sub

Private Sub lstSuppCodes_Click()
Dim i

    If lstSuppcodes.ListIndex = -1 Then
        tempBln = False
    Else
        tempBln = True
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Trim(Mid(lstSuppcodes.Text, 1, 5))
        Sqlqry = "Select * from SUPP_FIN Where Supp_No= '" & i & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MsgBox "Particular Record was Deleted.", vbInformation, "Deleted Status"
            Exit Sub
        End If
           txtcode = rs!Supp_no
           txtName = rs!Supp_name
           txtOpbal = rs!open_bal
           txtClbal = rs!Close_bal
           
           If IsNull(rs!Address) = True Then
            txtaddress = ""
           Else
            txtaddress = rs!Address
           End If
           
           If IsNull(rs!Con_Person) = True Then
              txtConPerson = ""
           Else
              txtConPerson = rs!Con_Person
           End If
           
           If IsNull(rs!Avail_Material) = True Then
              Txtmaterial = ""
           Else
              Txtmaterial = rs!Avail_Material
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
           
          If IsNull(rs!telephone) = True Then
              txtTelephone = ""
           Else
              txtTelephone = rs!telephone
          End If
          
          If IsNull(rs!fax) = True Then
              txtFax = ""
           Else
              txtFax = rs!fax
          End If
      
          If IsNull(rs!crdt_days) = True Then
              txtCrDays = ""
           Else
              txtCrDays = rs!crdt_days
          End If
          
          If IsNull(rs!crdt_lmt) = True Then
              txtCrLimit = ""
           Else
              txtCrLimit = rs!crdt_lmt
          End If
                 
          txtcode.SetFocus
          SendKeys "{home}+{end}"
         
End Sub

Private Sub txtaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtcity.SetFocus
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCountry.SetFocus
End Sub

Private Sub txtclbal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub

Private Sub txtcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtName.SetFocus
End Sub

Private Sub txtConPerson_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtmaterial.SetFocus
End Sub

Private Sub txtCountry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtTelephone.SetFocus
End Sub

Private Sub txtCrDays_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCrLimit.SetFocus
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtName.SetFocus
End Sub

Private Sub txtCrLimit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtConPerson.SetFocus
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCrDays.SetFocus
End Sub


Private Sub Txtmaterial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtOpbal.SetFocus
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtaddress.SetFocus
End Sub

Private Sub txtopbal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtClbal.SetFocus
End Sub

Private Sub txtTelephone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtFax.SetFocus
End Sub
