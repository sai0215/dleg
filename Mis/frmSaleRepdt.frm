VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmSaleRepdt 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Invoice List - Date Wise"
   ClientHeight    =   8775
   ClientLeft      =   -45
   ClientTop       =   330
   ClientWidth     =   11985
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
   ScaleHeight     =   8775
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Invoice List - Date Wise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4575
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   6855
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Preview"
         Height          =   975
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Clear"
         Height          =   975
         Left            =   2655
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<<&Back"
         Height          =   975
         Left            =   3870
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3120
         Width           =   1215
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   5400
         Top             =   2160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   960
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
         Left            =   2520
         TabIndex        =   7
         Top             =   1800
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
         BorderWidth     =   2
         X1              =   6840
         X2              =   -240
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Date To"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1320
         TabIndex        =   6
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Date From"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSaleRepdt"
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
Dim rs3 As Recordset
Private Sub cmdBack_Click()
 Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub
Private Sub cmdClear_Click()
 textclear
End Sub

Private Sub cmdDisplay_Click()
  If ValidateData = True Then
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = " Delete * from invrep"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        Sqlqry1 = "Select * from bo_mas where Invoice_date>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Invoice_date<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and status='N'"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        
        If rs.RecordCount <> 0 Then
         rs.MoveFirst
         Do Until rs.EOF
         
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         Sqlqry = " Insert into invrep values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "',' " _
                                     & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                     & Trim(rs!sub_media) & "','" & Trim(rs!tcurrency) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                     & Val(rs!tra_gamount) & "," _
                                     & Val(rs!Tot_free) & "," _
                                     & Val(rs!Tot_barter) & ",'" _
                                     & Val(Trim(rs!disc_percentage)) & "','" _
                                     & Val(Trim(rs!disc_rate)) & "'," _
                                     & Val(Trim(rs!add_discount)) & "," _
                                     & Val(Trim(rs!surcharge)) & "," _
                                     & Val(Trim(rs!tra_namount)) & ",'" & Format(rs!invoice_date, "dd/mm/yyy") & "')"
       ws.BeginTrans
       db.Execute (Sqlqry)
       ws.CommitTrans
       
       
       rs.MoveNext
       Loop
      End If
            
       
     With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\InvrepAgndt.rpt"
     .Formulas(0) = "zzz='" & " From " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
     .WindowMaxButton = True
     .WindowState = crptMaximized
     .Action = 1
     End With
     
   End If
End Sub

Private Sub Form_Load()
txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
txtdateto.TextWithMask = Format(Now, "dd/mm/yyyy")
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
If KeyAscii = 13 Then cmdDisplay.SetFocus
End Sub
Private Function ValidateData()
ValidateData = False

If IsDate(DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy"))) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
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
 txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
 txtdateto.TextWithMask = Format(Now, "dd/mm/YYYY")
End Sub

Private Sub txtdateto_LostFocus()
If IsDate(txtdateto.TextWithMask) = False Then
      MsgBox "Invalid Date To ", vbInformation, "Invalid Entry"
      txtdateto.SetFocus
      SendKeys "{Home} + {End}"
End If
End Sub
