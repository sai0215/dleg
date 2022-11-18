VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmBankMas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Creation & Modification of Bank"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   12060
   LinkTopic       =   "form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      Begin VB.Frame fraAccount 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Addition and Modifcation of Bank Register"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   6135
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   11175
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Closing Balance"
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
            Height          =   1095
            Left            =   240
            TabIndex        =   18
            Top             =   4440
            Width           =   5895
            Begin VB.TextBox txtclbalDHS 
               BackColor       =   &H00FFFFFF&
               DataField       =   "ACCT_CODE"
               DataSource      =   "Data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   1440
               TabIndex        =   20
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtclbalUSD 
               BackColor       =   &H00FFFFFF&
               DataField       =   "ACCT_CODE"
               DataSource      =   "Data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   3840
               TabIndex        =   19
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "DHS."
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
               Left            =   840
               TabIndex        =   22
               Top             =   600
               Width           =   555
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "USD"
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
               Left            =   3240
               TabIndex        =   21
               Top             =   600
               Width           =   495
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Opening Balance"
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
            Height          =   1095
            Left            =   240
            TabIndex        =   13
            Top             =   2880
            Width           =   5895
            Begin VB.TextBox txtopbalUSD 
               BackColor       =   &H00FFFFFF&
               DataField       =   "ACCT_CODE"
               DataSource      =   "Data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   3840
               TabIndex        =   17
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtOpbaldhs 
               BackColor       =   &H00FFFFFF&
               DataField       =   "ACCT_CODE"
               DataSource      =   "Data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   1440
               TabIndex        =   15
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "USD"
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
               Left            =   3240
               TabIndex        =   16
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "DHS."
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
               Left            =   840
               TabIndex        =   14
               Top             =   600
               Width           =   555
            End
         End
         Begin VB.TextBox txtaccode 
            BackColor       =   &H00FFFFFF&
            DataField       =   "ACCT_CODE"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1680
            TabIndex        =   10
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox txtacdesc 
            BackColor       =   &H00FFFFFF&
            DataField       =   "ACCT_NAME"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1680
            ScrollBars      =   1  'Horizontal
            TabIndex        =   9
            Top             =   2040
            Width           =   4095
         End
         Begin VB.ListBox lstAccodes 
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
            Height          =   4860
            Left            =   6720
            TabIndex        =   8
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Bank Name"
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
            TabIndex        =   12
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Bank Code"
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
            TabIndex        =   11
            Top             =   1320
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   2640
         TabIndex        =   1
         Top             =   6960
         Width           =   5175
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00FFFF80&
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            Picture         =   "frmBankMas.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   975
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
            Height          =   735
            Left            =   1080
            Picture         =   "frmBankMas.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   975
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
            Height          =   735
            Left            =   3000
            Picture         =   "frmBankMas.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   975
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
            Height          =   735
            Left            =   3960
            Picture         =   "frmBankMas.frx":0CC6
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   1095
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
            Height          =   735
            Left            =   2040
            Picture         =   "frmBankMas.frx":1108
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4200
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
End
Attribute VB_Name = "frmBankMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim tempBln
Dim Z
Dim totdhsop As Currency
Dim totdhscl As Currency
Dim totusdcl As Currency

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub cmdadd_Click()
    
  If ValidateData = True Then
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    
        Sqlqry1 = " select * from bank_mas where bank_code='" & Trim(txtaccode) & "' "
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
          MsgBox " Record already exists"
          Exit Sub
        Else
           totdhsop = Round(Val(txtOpbaldhs) + Val(txtopbalUSD) * convertion, 2)
           totdhscl = Round(Val(txtclbalDHS) + Val(txtclbalUSD) * convertion, 2)
             
            Sqlqry = " Insert into acct_mas values('" & txtaccode & "','" & _
                    Trim(txtacdesc) & "'," & _
                    totdhsop & "," & _
                    totdhscl & ")"
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
    
            Sqlqry = " Insert into bank_mas values('" & txtaccode & "','" & _
                    Trim(txtacdesc) & "','" & _
                    Trim(txtOpbaldhs) & "','" & _
                    Trim(txtopbalUSD) & "','" & _
                    Trim(txtclbalDHS) & "','" & _
                    Trim(txtclbalUSD) & "')"
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
txtOpbaldhs = ""
txtclbalDHS = ""
txtopbalUSD = ""
txtclbalUSD = ""

End Function

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
ElseIf IsNumeric(txtOpbaldhs) = False Then
   MsgBox "Invalid Opening Balance - DHS", vbInformation, "Invalid Entry"
   txtOpbaldhs.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsNumeric(txtclbalDHS) = False Then
   MsgBox "Invalid Closing Balance - DHS", vbInformation, "Invalid Entry"
   txtclbalDHS.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsNumeric(txtopbalUSD) = False Then
   MsgBox "Invalid Opening Balance - USD", vbInformation, "Invalid Entry"
   txtopbalUSD.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsNumeric(txtclbalUSD) = False Then
   MsgBox "Invalid Closing Balance - USD", vbInformation, "Invalid Entry"
   txtclbalDHS.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function

ValidateData = True
End If
ValidateData = True
End Function

Private Sub cmdMod_Click()
 Z = 1
 Dim tempStr
    If lstAccodes.SelCount = 0 Then
        MsgBox "Select the Bank Code for Modification.", vbInformation, "Selection Error"
        lstAccodes.SetFocus
        Exit Sub
    End If
        If ValidateData = False Then Exit Sub
        tempStr = MsgBox("Do You Want To Modify the Bank Code :" & txtaccode, vbQuestion + vbYesNo, "Confirmation")
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
       
    i = Val(txtaccode.Text)
    
           totdhsop = Round(Val(txtOpbaldhs) + Val(txtopbalUSD) * convertion, 2)
           totdhscl = Round(Val(txtclbalDHS) + Val(txtclbalUSD) * convertion, 2)
                   
        
       Sqlqry = "Update Acct_mas Set " _
                    & " acct_Name = '" & Trim(txtacdesc.Text) & "'," _
                    & " open_bal = " & totdhsop & "," _
                    & " close_bal = " & totdhscl & " " _
                    & " Where Acct_Code = '" & i & "' "
                                           
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        
        Sqlqry = "Update Bank_mas Set " _
                   & " Bank_name = '" & Trim(txtacdesc.Text) & "'," _
                   & " Open_baldhs = " & Val(txtOpbaldhs.Text) & "," _
                   & " Open_balusd = " & Val(txtopbalUSD.Text) & "," _
                   & " Close_baldhs =" & Val(txtclbalDHS.Text) & ", " _
                   & " Close_balusd =" & Val(txtclbalUSD.Text) & " " _
                   & " Where Bank_code = '" & i & "'"
                                           
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        MsgBox "Record Modified With " & Chr(10) & Chr(10) & _
               "Bank Code = " & i, vbInformation, "Data Modified"
        textclear
        PopulateAccodes
        tempBln = False
        ModifyData = True
        Exit Function
End Function



Private Sub PopulateAccodes()
        
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Bank_mas Order by Bank_code"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
    rs.MoveFirst
        lstAccodes.Clear
        Do Until rs.EOF
            lstAccodes.AddItem rs!bank_code & "    :    " & rs!BANK_NAME
            rs.MoveNext
        Loop
    End If
        
End Sub

Private Sub CmdPrint_Click()
CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
CrystalReport1.ReportFileName = App.Path & "\BankList.rpt"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
End Sub

Private Sub Form_Load()
    tempBln = False
    PopulateAccodes
    textclear
  
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
        Sqlqry = "Select * from bank_mas Where bank_code= '" & i & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MsgBox "Selected Record was not found", vbInformation, "Deleted Status"
            Exit Sub
        End If
           txtaccode = rs!bank_code
           txtacdesc = rs!BANK_NAME
           If IsNull(rs!Open_baldhs) = True Then
             txtOpbaldhs = 0
           Else
             txtOpbaldhs = rs!Open_baldhs
           End If
           
           If IsNull(rs!Close_baldhs) = True Then
             txtclbalDHS = 0
           Else
             txtclbalDHS = rs!Close_baldhs
           End If
           
           
           If IsNull(rs!open_balUSD) = True Then
             txtopbalUSD = 0
           Else
             txtopbalUSD = rs!open_balUSD
           End If
           
           If IsNull(rs!Close_balusd) = True Then
             txtclbalUSD = 0
           Else
             txtclbalUSD = rs!Close_balusd
           End If
           
           
    txtaccode.SetFocus
    SendKeys "{home}+{end}"
        
End Sub
Private Sub txtaccode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtacdesc.SetFocus
End Sub
Private Sub txtacdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtOpbaldhs.SetFocus
End Sub


Private Sub txtclbalDHS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtclbalUSD.SetFocus
End Sub

Private Sub txtclbalUSD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd.SetFocus
End Sub

Private Sub txtopbaldhs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtopbalUSD.SetFocus
End Sub


Private Sub txtopbalUSD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtclbalDHS.SetFocus
End Sub
