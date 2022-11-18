VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmPdcPaymentPosting 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Pdc Payment Posting"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "PDC ISSUE - POSTING"
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
      Height          =   8055
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   11655
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   4320
         Picture         =   "frmPdcPaymentPosting.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6960
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   5640
         Picture         =   "frmPdcPaymentPosting.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6960
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   2880
         Picture         =   "frmPdcPaymentPosting.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6960
         Width           =   1455
      End
      Begin VB.CommandButton cmdL 
         BackColor       =   &H80000000&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdLL 
         BackColor       =   &H80000000&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Cmdg 
         BackColor       =   &H80000000&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1800
         Width           =   1095
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
         Height          =   2220
         Left            =   6600
         MultiSelect     =   1  'Simple
         TabIndex        =   6
         Top             =   1440
         Width           =   4695
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
         Height          =   2220
         Left            =   240
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         Top             =   1440
         Width           =   4695
      End
      Begin VB.CommandButton cmdGG 
         BackColor       =   &H80000000&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2160
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2295
         Left            =   240
         TabIndex        =   11
         Top             =   4080
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   4
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
         BackColorFixed  =   8388608
         BackColorBkg    =   8421376
         GridLines       =   2
         AllowUserResizing=   1
      End
      Begin PVMaskEditLib.PVMaskEdit txtdate 
         Height          =   375
         Left            =   5040
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
         BorderColor     =   &H000000C0&
         X1              =   11640
         X2              =   0
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Line Line1 
         X1              =   11640
         X2              =   0
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Total Cheques Due on posting Dt."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   405
         Left            =   480
         TabIndex        =   14
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cheques for Posting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   405
         Left            =   7920
         TabIndex        =   13
         Top             =   1080
         Width           =   2445
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enter Date of Posting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   360
         Width           =   2745
      End
   End
End
Attribute VB_Name = "frmPdcPaymentPosting"
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

Private Sub Flexitems()
With MSFlexGrid1
    .Clear
    .AllowUserResizing = flexResizeColumns
    .Rows = 1
    .Cols = 7
    .Col = 0
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Code"
    .ColAlignment(0) = 0
    .ColWidth(0) = 800
    .ColWidth(1) = 3700
    .ColWidth(2) = 750
    .ColWidth(3) = 3000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 825
    .Col = 1
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Name"
    .Col = 2
    .CellBackColor = RGB(180, 170, 160)
    .Text = "B_Code"
    .Col = 3
    .CellBackColor = RGB(180, 170, 160)
    .Text = "B_Name"
    .Col = 4
    .CellBackColor = RGB(180, 170, 160)
    .Text = "CH_NO"
    .Col = 5
    .CellBackColor = RGB(180, 170, 160)
    .Text = "CH_DT"
    .CellBackColor = RGB(180, 170, 160)
    .Col = 6
    .CellBackColor = RGB(180, 170, 160)
    .Text = "    Amount"
    .Row = 0
    .Col = 1

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
Private Sub CmdSave_Click()
i = lstpost.ListIndex
If lstpost.ListCount = 0 Then
    MsgBox " Select Entries to be posted"
    lstFlex.SetFocus
    Exit Sub
Else
      For i = 0 To lstpost.ListCount - 1
       Set ws = DBEngine.Workspaces(0)
       Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       Sqlqry = "Update PPMT_MAS set posting_dt=#" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "#, " & " Status ='Y' where vouc_no=" & Val(Mid(lstpost.List(i), 1, 6)) & ""
       Sqlqry1 = "Update PPMT_tra set posting_dt=#" & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "# where vouc_no=" & Val(Mid(lstpost.List(i), 1, 6)) & ""
       ws.BeginTrans
       db.Execute (Sqlqry)
       db.Execute (Sqlqry1)
       ws.CommitTrans
      Next
       MsgBox " Entries Posted =" & lstpost.ListCount
       lstFlex.Clear
       lstpost.Clear
       Flexitems
       txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
    
 End If

End Sub
Private Sub Form_Load()
 txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
End Sub
Private Sub txtDate_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CmdSave.SetFocus
End Sub

Private Sub txtdate_LostFocus()

 If IsDate(txtdate.TextWithMask) = False Then
      MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
      txtdate.SetFocus
      SendKeys "{Home} + {End}"
 End If
 
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Select * from PPMT_MAS where IsNull(Posting_Dt) and Cheque_Dt<=#" & DateValue(Format(txtdate.TextWithMask, "dd/mm/yyyy")) & "# order by vouc_no"
 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

 MSFlexGrid1.Clear
 
 If rs.RecordCount <> 0 Then
    
   If IsDate(DateValue(Format(txtdate.TextWithMask, "dd/mm/yyyy"))) = True Then
      rs.MoveFirst
      Flexitems
      lstFlex.Clear
      lstpost.Clear
      Do Until rs.EOF
       MSFlexGrid1.AddItem rs!VOUC_NO & Chr(9) & rs!PAID_TO & Chr(9) & rs!bank_code & Chr(9) & rs!BANK_NAME & Chr(9) & rs!CHEQUE_NO & Chr(9) & rs!Cheque_Dt & Chr(9) & rs!tra_amount
       lstFlex.AddItem rs!VOUC_NO & ":" & rs!BANK_NAME & ":" & rs!tra_amount
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
  MSFlexGrid1.Clear
  txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
End Sub

