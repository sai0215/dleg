VERSION 5.00
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "PVMASK.OCX"
Begin VB.Form frmPdcRecPosting1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "PDC  Receipts Posting for Accounting Effect"
   ClientHeight    =   8775
   ClientLeft      =   15
   ClientTop       =   345
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "PDC Receipts"
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
      Height          =   8175
      Left            =   480
      TabIndex        =   9
      Top             =   240
      Width           =   10695
      Begin VB.CommandButton cmdGG 
         BackColor       =   &H00C0C0C0&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4080
         Picture         =   "frmPdcRecPosting1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3360
         Width           =   975
      End
      Begin VB.ListBox lstFlex 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1980
         Left            =   1800
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         Top             =   1080
         Width           =   7695
      End
      Begin VB.ListBox lstpost 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1980
         Left            =   1800
         MultiSelect     =   1  'Simple
         TabIndex        =   10
         Top             =   4560
         Width           =   7695
      End
      Begin VB.CommandButton Cmdg 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Selected"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3000
         Picture         =   "frmPdcRecPosting1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdLL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5040
         Picture         =   "frmPdcRecPosting1.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Selected"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6000
         Picture         =   "frmPdcRecPosting1.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3120
         Picture         =   "frmPdcRecPosting1.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7080
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         Picture         =   "frmPdcRecPosting1.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7080
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
         Height          =   855
         Left            =   4320
         Picture         =   "frmPdcRecPosting1.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7080
         Width           =   1215
      End
      Begin PVMaskEditLib.PVMaskEdit txtdate 
         Height          =   375
         Left            =   4680
         TabIndex        =   0
         Top             =   360
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
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
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   10680
         X2              =   0
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   10680
         X2              =   0
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enter Date of Depositing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   2160
         TabIndex        =   13
         Top             =   480
         Width           =   2385
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "For Deposit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "To be Deposited"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmPdcRecPosting1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim i
Dim f
Dim Z
Dim sum As Integer
Dim rs As Recordset
Dim rs1 As Recordset
Dim Sqlqry As String
Dim Sqlqry1 As String
Private Sub CmdBack_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub cmdClear_Click()
textclear
End Sub

Private Sub CmdSave_Click()
Dim X
i = lstpost.ListIndex
If lstpost.ListCount = 0 Then
    MsgBox " Select Entries to be posted"
    lstFlex.SetFocus
    Exit Sub
Else
      For i = 0 To lstpost.ListCount - 1
       Set ws = DBEngine.Workspaces(0)
       Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       Sqlqry = "update prpt_mas1 set posting_dt=#" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "# ," & " status='Y' where vouc_no=" & Val(Mid(lstpost.List(i), 1, 6)) & " and Cheque_Dt=# " & Mid(lstpost.List(i), 8, 10) & " # and cheque_no='" & Val(Mid(lstpost.List(i), 19, 10)) & "' "
       
       ws.BeginTrans
       db.Execute (Sqlqry)
       ws.CommitTrans
   Next
       MsgBox " Entries Posted =" & lstpost.ListCount
       lstFlex.Clear
       lstpost.Clear
       txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
 End If

End Sub

Private Sub Form_Load()
 txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
  
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtdate_LostFocus()

If IsDate(txtdate.TextWithMask) = False Then
      MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
      txtdate.SetFocus
      SendKeys "{Home} + {End}"
    End If
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Select * from prpt_mas1 where IsNull(Posting_Dt)=true  and Cheque_Dt<=#" & DateValue(Format(txtdate.TextWithMask, "dd/mm/yyyy")) & "# and status='N' order by vouc_no"
 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

  
  If rs.RecordCount <> 0 Then
    
   If IsDate(DateValue(Format(txtdate.TextWithMask, "dd/mm/yyyy"))) = True Then
      rs.MoveFirst
      lstFlex.Clear
      lstpost.Clear
      Do Until rs.EOF
       lstFlex.AddItem rs!VOUC_NO & Chr(9) & Format(rs!Cheque_Dt, "dd/mm/yyyy") & "  " & rs!CHEQUE_NO & " " & Trim(rs!BANK_NAME) & " " & rs!tra_amount
       rs.MoveNext
      Loop
   Else
      MsgBox "Invalid Date"
   End If
 End If
End Sub
 
Private Sub Cmdg_Click()
 For i = lstFlex.ListCount - 1 To 0 Step -1
    If lstFlex.Selected(i) Then
       lstpost.AddItem lstFlex.List(i)
       lstFlex.RemoveItem (i)
    End If
Next
End Sub

Private Sub Cmdgg_Click()
 For i = lstFlex.ListCount - 1 To 0 Step -1
         lstpost.AddItem lstFlex.List(i)
         lstFlex.RemoveItem (i)
 Next i

End Sub

Private Sub Cmdl_Click()
 For f = lstpost.ListCount - 1 To 0 Step -1
    
    If lstpost.Selected(f) Then
       lstFlex.AddItem lstpost.Text
       lstpost.RemoveItem (f)
    End If
 Next
End Sub

Private Sub Cmdll_Click()
 For i = lstpost.ListCount - 1 To 0 Step -1
         lstFlex.AddItem lstpost.List(i)
         lstpost.RemoveItem (i)
 Next i
End Sub
Private Sub textclear()
    lstFlex.Clear
    lstpost.Clear
    txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
End Sub
