VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmPdcPmtRepDt 
   BackColor       =   &H00FFFFC0&
   Caption         =   "PDC Payment - Date Wise"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   FillColor       =   &H00400000&
   ForeColor       =   &H00C0C0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "PDC Issued - Date Range"
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
      Height          =   5055
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Width           =   6855
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
         Height          =   975
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFF80&
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdPending 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Pending PDC"
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdAllPdc 
         BackColor       =   &H00FFFF80&
         Caption         =   "&All PDC"
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
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3840
         Width           =   1215
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   6000
         Top             =   3960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
         Height          =   375
         Left            =   2640
         TabIndex        =   0
         Top             =   1200
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
         Left            =   2640
         TabIndex        =   1
         Top             =   2040
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
         BorderWidth     =   2
         X1              =   6840
         X2              =   0
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Date To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1440
         TabIndex        =   8
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Date From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1440
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmPdcPmtRepDt"
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
Dim Sqlqry4 As String
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim i

Private Sub cmdAllPdc_Click()
 If ValidateData = True Then
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry = " Delete * from Pdcreport"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        Sqlqry1 = "select * from ppmt_mas where Cheque_Dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Cheque_dt<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        
        If rs.RecordCount <> 0 Then
         rs.MoveFirst
          If IsNull(rs!posting_Dt) = True Then
             i = ""
          Else
             i = DateValue(rs!posting_Dt)
          End If
         Do Until rs.EOF
            Sqlqry2 = "Select * from ppmt_tra where vouc_no=" & rs!VOUC_NO & ""
            Set rs1 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
            rs1.MoveFirst
            Do Until rs1.EOF
            If IsNull(rs1!Description) = True Then
               rs1!Description = ""
            End If
             
             Sqlqry3 = "Insert into pdcreport Values(" & Trim(rs1!VOUC_NO) & ",'" & Trim(rs1!tDate) & "','" _
                        & Trim(rs1!acct_code) & "','" & findfirstfixup(Trim(rs1!acct_name)) & "','" _
                        & findfirstfixup(Trim(rs1!Description)) & "','" & Trim(rs1!tcurrency) & "'," & Trim(rs1!tra_amount) & ",'" & Trim(rs!BANK_NAME) & "','" _
                        & Trim(rs!CHEQUE_NO) & "', '" & Trim(rs!Cheque_Dt) & "' , '" & i & "')"
             ws.BeginTrans
             db.Execute (Sqlqry3)
             ws.CommitTrans
             rs1.MoveNext
            Loop
            
          rs.MoveNext
          Loop
         End If
       End If
     With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\PdcPayRep.rpt"
     .Formulas(0) = "zzz='" & " From " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
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

Private Sub cmdPending_Click()
If ValidateData = True Then
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry = " Delete * from Pdcreport"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        Sqlqry1 = "select * from ppmt_mas where Cheque_Dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Cheque_Dt<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and isnull(Posting_dt)"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        
        If rs.RecordCount <> 0 Then
         rs.MoveFirst
          
         Do Until rs.EOF
            Sqlqry2 = "Select * from ppmt_tra where vouc_no=" & rs!VOUC_NO & ""
            Set rs1 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
            rs1.MoveFirst
            Do Until rs1.EOF
            If IsNull(rs1!Description) = True Then
               rs1!Description = ""
            End If
             
             Sqlqry3 = "Insert into pdcreport Values(" & Trim(rs1!VOUC_NO) & ",'" & Trim(rs1!tDate) & "','" _
                        & Trim(rs1!acct_code) & "','" & findfirstfixup(Trim(rs1!acct_name)) & "','" _
                        & findfirstfixup(Trim(rs1!Description)) & "','" & Trim(rs1!tcurrency) & "'," & Trim(rs1!tra_amount) & ",'" & Trim(rs!BANK_NAME) & "','" _
                        & Trim(rs!CHEQUE_NO) & "', '" & Trim(rs!Cheque_Dt) & "' , '')"
             ws.BeginTrans
             db.Execute (Sqlqry3)
             ws.CommitTrans
             rs1.MoveNext
            Loop
            
          rs.MoveNext
          Loop
         End If
       End If
     With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\PdcPayRep1.rpt"
     .Formulas(0) = "zzz='" & " From " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
     .WindowMaxButton = True
     .WindowState = crptMaximized
     .Action = 1
    End With
End Sub

Private Sub Form_Load()
txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
txtdateto.TextWithMask = Format(Now, "dd/mm/yyyy")
txtdatefrom.TabIndex = 0
' SendKeys " {Home} + {End} "
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdateto.SetFocus
End Sub

Private Sub txtdatefrom_LostFocus()
 If IsDate(txtdatefrom.TextWithMask) = False Then
      MsgBox "Invalid Date From ", vbInformation, "Invalid Entry"
      txtdatefrom.SetFocus
      SendKeys "{Home} + {End}"
 End If
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdPending.SetFocus
End Sub
Private Function ValidateData()
ValidateData = False

If IsDate(txtdatefrom.TextWithMask) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf IsDate(txtdateto.TextWithMask) = False Then
   MsgBox "Invalid To Date", vbInformation, "Invalid Entry"
   txtdateto.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
ValidateData = True
End If
End Function
Private Sub textclear()
 txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
 txtdateto.TextWithMask = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub txtdateto_LostFocus()
 If IsDate(txtdateto.TextWithMask) = False Then
      MsgBox "Invalid Date To ", vbInformation, "Invalid Entry"
      txtdateto.SetFocus
      SendKeys "{Home} + {End}"
 End If
End Sub
