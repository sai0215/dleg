VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmPdcRptRepCust 
   BackColor       =   &H00FFFFC0&
   Caption         =   "PDC Details  - Customer Wise"
   ClientHeight    =   8175
   ClientLeft      =   75
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
      Caption         =   "PDC  Detatils - Agency Wise "
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
      Height          =   6495
      Left            =   2040
      TabIndex        =   7
      Top             =   480
      Width           =   7695
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Back <<"
         Height          =   855
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdAllPdc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&All"
         Height          =   855
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Clear"
         Height          =   855
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
         Width           =   1215
      End
      Begin VB.ListBox lstCustomers 
         BackColor       =   &H80000018&
         ForeColor       =   &H00404040&
         Height          =   2700
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   6975
      End
      Begin VB.CommandButton cmdPendingPDC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Pending"
         Height          =   855
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5400
         Width           =   1215
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   6840
         Top             =   5280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   3840
         Width           =   1455
         _Version        =   65541
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   244
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         Text            =   ""
         Mask            =   "##/##/####"
         PromptCharacter =   ""
         BackColor       =   16777215
         ForeColor       =   8388608
         Alignment       =   1
      End
      Begin PVMaskEditLib.PVMaskEdit txtdateto 
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   4440
         Width           =   1455
         _Version        =   65541
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   244
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         Text            =   ""
         Mask            =   "##/##/####"
         PromptCharacter =   ""
         BackColor       =   16777215
         ForeColor       =   8388608
         Alignment       =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   0
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Enter Date To"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1320
         TabIndex        =   9
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Enter Date From"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1320
         TabIndex        =   8
         Top             =   3960
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmPdcRptRepCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim Sqlqry3 As String
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim rs3 As Recordset

Private Sub cmdAllPdc_Click()
 If ValidateData = True Then
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = " Delete * from Pdcreport"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        Sqlqry1 = "select * from prpt_mas1 where Cheque_dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Cheque_dt<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_name='" & lstCustomers & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        
         If rs1.RecordCount <> 0 Then
          rs1.MoveFirst
          Do Until rs1.EOF
           If IsNull(rs1!Description) = True Then
               rs1!Description = ""
           End If
             
             Sqlqry3 = "Insert into pdcreport Values(" & Trim(rs1!VOUC_NO) & ",'" & Trim(rs1!tDate) & "','" _
                        & Trim(rs1!acct_code) & "','" & findfirstfixup(Trim(rs1!acct_name)) & "','" _
                        & findfirstfixup(Trim(rs1!Description)) & "','" & Trim(rs1!tcurrency) & "'," & Trim(rs1!tra_amount) & ",'" & Trim(rs1!BANK_NAME) & "','" _
                        & Trim(rs1!CHEQUE_NO) & "', '" & Trim(rs1!Cheque_Dt) & "' , '" & Trim(rs1!posting_Dt) & "')"
             ws.BeginTrans
             db.Execute (Sqlqry3)
             ws.CommitTrans
            rs1.MoveNext
          Loop
         End If
         
       End If
     With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\Pdcrepcus2.rpt"
     .Formulas(0) = "zzz='" & " From " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
     .Formulas(1) = "yyy='" & lstCustomers.Text & " '"
     .WindowMaxButton = True
     .WindowState = crptMaximized
     .Action = 1
    End With
       
End Sub

Private Sub cmdBack_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub cmdClear_Click()
textclear
End Sub

Private Sub cmdPendingPdc_Click()
If ValidateData = True Then
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry = " Delete * from Pdcreport"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        Sqlqry1 = "select * from prpt_mas1 where Cheque_Dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Cheque_dt<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and isnull(Posting_dt) and acct_name='" & Trim(lstCustomers.Text) & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        
        If rs1.RecordCount <> 0 Then
          rs1.MoveFirst
            Do Until rs1.EOF
            If IsNull(rs1!Description) = True Then
               rs1!Description = ""
            End If
             
              Sqlqry3 = "Insert into pdcreport Values(" & Trim(rs1!VOUC_NO) & ",'" & Trim(rs1!tDate) & "','" _
                        & Trim(rs1!acct_code) & "','" & findfirstfixup(Trim(rs1!acct_name)) & "','" _
                        & findfirstfixup(Trim(rs1!Description)) & "','" & Trim(rs1!tcurrency) & "'," & Trim(rs1!tra_amount) & ",'" & Trim(rs1!BANK_NAME) & "','" _
                        & Trim(rs1!CHEQUE_NO) & "', '" & Trim(rs1!Cheque_Dt) & "' , '')"
              ws.BeginTrans
              db.Execute (Sqlqry3)
              ws.CommitTrans
             rs1.MoveNext
            Loop
        End If
        
       End If
     With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\PdcRepCus1.rpt"
     .Formulas(0) = "zzz='" & " From " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
     .Formulas(1) = "yyy='" & lstCustomers.Text & " '"
     .WindowMaxButton = True
     .WindowState = crptMaximized
     .Action = 1
    End With
       
End Sub

Private Sub Form_Load()
populateCustomers
txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
txtdateto.TextWithMask = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub populateCustomers()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select * from Agndtls order by Agentname"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

lstCustomers.Clear

 If rs.RecordCount = 0 Then
      MsgBox "No Records found in the Agency Register"
 Else
      rs.MoveFirst
   Do Until rs.EOF
      lstCustomers.AddItem rs!agentname
      rs.MoveNext
   Loop
 End If

End Sub
Private Sub lstcustomers_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdatefrom.SetFocus
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdateto.SetFocus
End Sub

Private Sub txtdatefrom_LostFocus()
If IsDate(txtdatefrom.TextWithMask) = False Then
      MsgBox "Invalid Date from ", vbInformation, "Invalid Entry"
      txtdatefrom.SetFocus
      SendKeys "{Home} + {End}"
    End If
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdPendingPdc.SetFocus
End Sub
Private Function ValidateData()
ValidateData = False

If IsDate(DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy"))) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf lstCustomers.SelCount = 0 Then
  MsgBox "Select Customer", vbInformation, "Invalid Entry"
  lstCustomers.SetFocus
  SendKeys " {Home} + {end} "
  Exit Function
ElseIf IsDate(DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy"))) = False Then
   MsgBox "Invalid To Date", vbInformation, "Invalid Entry"
   txtdateto.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
ValidateData = True
End If
End Function
Private Sub textclear()
 lstCustomers.ListIndex = -1
 txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
 txtdateto.TextWithMask = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub txtdateto_LostFocus()
If IsDate(txtdateto.TextWithMask) = False Then
      MsgBox "Invalid Date to ", vbInformation, "Invalid Entry"
      txtdateto.SetFocus
      SendKeys "{Home} + {End}"
End If
End Sub
