VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "PVMASK.OCX"
Begin VB.Form frmPdcReceiptAddition1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pdc Receipt Addition 1"
   ClientHeight    =   8775
   ClientLeft      =   -45
   ClientTop       =   255
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PDC Receipt Addition"
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
      Height          =   8295
      Left            =   480
      TabIndex        =   14
      Top             =   240
      Width           =   11295
      Begin VB.TextBox txtTtlAmount 
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cboCurrency 
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
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtConvRate 
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   7245
         TabIndex        =   29
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtChequeBank 
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
         Height          =   315
         Left            =   2760
         TabIndex        =   8
         Top             =   3720
         Width           =   2895
      End
      Begin VB.ComboBox lstAcctCode 
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
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   4215
      End
      Begin VB.TextBox txtTtlDesc 
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
         Height          =   315
         Left            =   1440
         MaxLength       =   150
         TabIndex        =   5
         Top             =   2760
         Width           =   7575
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
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
         Height          =   735
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7320
         Width           =   975
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFFFF&
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
         Height          =   735
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7320
         Width           =   975
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
         Height          =   735
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7320
         Width           =   975
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   9360
         TabIndex        =   10
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtRecdFrom 
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
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1560
         Width           =   4215
      End
      Begin VB.ComboBox cboBankCode 
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
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox txtChequeNO 
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
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   855
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   8400
         Top             =   1920
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2175
         Left            =   240
         TabIndex        =   15
         Top             =   4320
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   5
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
         BackColorFixed  =   12111599
         BackColorBkg    =   8421376
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PVMaskEditLib.PVMaskEdit txtdate 
         Height          =   375
         Left            =   4200
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
      Begin PVMaskEditLib.PVMaskEdit txtchequedt 
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   3720
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
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   2760
         TabIndex        =   32
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Currency"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label lblConvRate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Conv. Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   5985
         TabIndex        =   30
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   11280
         X2              =   0
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cheque Bank"
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
         Left            =   3120
         TabIndex        =   28
         Top             =   3480
         Width           =   1395
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   11280
         X2              =   0
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Description"
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
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label lblTtlAmt 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   8880
         TabIndex        =   26
         Top             =   6600
         Width           =   690
      End
      Begin VB.Label LblTtlAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   9720
         TabIndex        =   25
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label lblVoucNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1440
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amount"
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
         Left            =   9720
         TabIndex        =   23
         Top             =   3480
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Agency"
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
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date"
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
         Left            =   3600
         TabIndex        =   21
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Payer Name"
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
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Voucher No."
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
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Depositing Bank"
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
         Left            =   6720
         TabIndex        =   18
         Top             =   3480
         Width           =   1725
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Chq #"
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
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Chq Date"
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
         Left            =   1320
         TabIndex        =   16
         Top             =   3480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPdcReceiptAddition1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim Sqlqry3 As String
Dim a
Dim B
Dim X
Dim Y As Long
Dim Z
Dim i
Dim j
Dim fdate As Date
Dim ldate As Date

Private Sub cboBankCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtamount.SetFocus
End Sub

Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtTtlAmount.SetFocus
End Sub
Private Sub cboCurrency_LostFocus()
    If CboCurrency.Text = "USD" Then
     lblConvRate.Visible = True
     txtConvRate.Visible = True
     txtConvRate.Text = ""
     txtConvRate.TabIndex = 3
    Else
     lblConvRate.Visible = False
     txtConvRate.Visible = False
     txtConvRate.Text = 1
     txtConvRate.TabIndex = 14
    End If
End Sub
Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub cmdClear_Click()
     textclear
     lblTtlAmount.Caption = ""
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from DUMPRPT2"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     MSFlexGrid1.Clear
     txtTtlAmount.SetFocus
End Sub

Private Sub CmdSave_Click()
Dim ctype As String
 If ValidateData = True Then
   If Val(txtTtlAmount.Text) = lblTtlAmount Then
         
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   
    Sqlqry1 = "Select * from DUMPRPT2"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount = 0 Then
         MsgBox " Transactions are not recorded"
         Exit Sub
      Else
         rs.MoveFirst
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         Sqlqry3 = "Insert into Prpt_Mas1 values('" & rs!vouc_no & "','" & rs!vouc_type & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(txtRecdFrom) & "','" _
                                     & findfirstfixup(Trim(txtTtlDesc)) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & ",'" _
                                     & Mid(lstAcctCode, 1, 4) & "','" _
                                     & Mid(lstAcctCode, 12, 35) & "','" _
                                     & Trim(rs!bank_code) & "','" _
                                     & Trim(rs!BANK_NAME) & "','" _
                                     & rs!CHEQUE_NO & "','" _
                                     & Trim(rs!cheque_Bank) & "','" _
                                     & Trim(rs!Cheque_Dt) & "',' ','" _
                                     & Val(rs!Amount) & "','N')"

            ws.BeginTrans
            db.Execute (Sqlqry3)
            ws.CommitTrans
          rs.MoveNext
         Loop
       End If
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Update Docu_mas set Doc_no='" & lblVoucNo & "' where Doc_type='PRT'"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     ctype = CboCurrency.Text
     textclear
     lblVoucNo = lblVoucNo + 1
     
   Else
   MsgBox "Total amount is not equal to Total transactions"
   Exit Sub
   End If
  End If
  MsgBox " Record is inserted", vbInformation, "Status"
  Dim X As Integer
  X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
  If X = vbYes Then
   If CboCurrency.Text = "DHS" Then
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\prptVou1.rpt"
        CrystalReport1.SelectionFormula = "{Prpt_mas1.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
        CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtTtlAmount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & CboCurrency.Text & "'"
        CrystalReport1.Action = 1
    Else
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\prptVou1.rpt"
        CrystalReport1.SelectionFormula = "{Prpt_mas1.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
        CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(txtTtlAmount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & CboCurrency.Text & "'"
        CrystalReport1.Action = 1
    End If
  End If
End Sub

Private Sub Form_Load()
 B = 0
 txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
 txtchequedt.TextWithMask = Format(Now, "dd/mm/yyyy")
 CboCurrency.AddItem "DHS"
 CboCurrency.AddItem "USD"
 
 lblConvRate.Visible = False
 txtConvRate.Visible = False
 txtConvRate.Text = 1
 
 AutoIncrementVoucher
 PopulateAcctSuppCust
 PopulateBankCodes
 Flexitems
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
  Sqlqry = "Delete * from DUMPRPT2"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
End Sub

Private Sub AutoIncrementVoucher()
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Select DOC_no from DOCU_MAS WHERE DOC_TYPE='PRT'"
 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
 If rs.RecordCount = 0 Then
   MsgBox "Document type 'PRT' not found"
   Exit Sub
 Else
   lblVoucNo = Val(rs!doc_no) + 1
 End If
End Sub

Private Sub PopulateAcctSuppCust()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry1 = "Select * from Supp_fin order by Supp_no"
Sqlqry2 = "Select * from Agndtls order by AgentName"
Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)

lstAcctCode.Clear

If rs2.RecordCount = 0 Then
    MsgBox "No Records found in the Agency Register"
Else
   rs2.MoveFirst
   Do Until rs2.EOF
      lstAcctCode.AddItem "AGNC" & "    :  " & rs2!agentname
      rs2.MoveNext
   Loop
End If


If rs1.RecordCount = 0 Then
    MsgBox "No Records found in the Supplier Register"
Else
   rs1.MoveFirst
   Do Until rs1.EOF
      lstAcctCode.AddItem rs1!Supp_no & "    :  " & rs1!SUPP_NAME
      rs1.MoveNext
   Loop
End If

    
End Sub

Private Function ValidateData()

ValidateData = False
If txtdate = "" Or IsDate(DateValue(txtdate.TextWithMask)) = False Then
  MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
  txtdate.SetFocus
  SendKeys "{Home} + {End}"
  Exit Function
ElseIf txtConvRate.Text = "" Then
  MsgBox "Enter Convertion Rate - - cannot be zero", vbInformation, "Invalid Entry"
  txtConvRate.SetFocus
  Exit Function
ElseIf txtTtlAmount.Text = "" Or IsNumeric(txtTtlAmount) = False Then
  MsgBox "Invalid Total Amount", vbInformation, "Invalid Entry"
  txtTtlAmount.SetFocus
  Exit Function
ElseIf lstAcctCode.Text = "" Then
  MsgBox "Select Customer Code from list box", vbInformation, "Invalid Entry"
  lstAcctCode.SetFocus
  Exit Function
ElseIf txtTtlDesc.Text = "" Or IsNumeric(txtTtlDesc) = True Then
  MsgBox "Invalid Description", vbInformation, "Invalid Entry"
  txtTtlDesc.SetFocus
  Exit Function
ElseIf txtamount.Text = "" Or IsNumeric(txtamount) = False Then
  MsgBox "Invalid Amount", vbInformation, "Invalid Entry"
  txtamount.SetFocus
  Exit Function
ElseIf txtChequeNO.Text = "" Then
  MsgBox "Invalid Cheque No.", vbInformation, "Invalid Entry"
  txtChequeNO.SetFocus
  Exit Function
ElseIf IsDate(txtchequedt.TextWithMask) = False Then
  MsgBox "Invalid Cheque Date", vbInformation, "Invalid Entry"
  txtchequedt.SetFocus
  Exit Function
ElseIf cboBankCode.Text = "" Or IsNumeric(txtamount) = False Then
  MsgBox "Invalid Amount", vbInformation, "Invalid Entry"
  txtamount.SetFocus
  Exit Function
Else
 ValidateData = True
End If
End Function

Private Sub Flexitems()
With MSFlexGrid1

    .Clear
    .AllowUserResizing = flexResizeColumns
    .Rows = 1
    .Cols = 6
    .Col = 0
    .CellBackColor = RGB(180, 170, 160)
    .CellAlignment = 2
    .Text = "Cheque_No"
    .ColAlignment(0) = 0
    .ColWidth(0) = 1000
    .ColWidth(1) = 1200
    .ColWidth(2) = 2100
    .ColWidth(3) = 1000
    .ColWidth(4) = 4100
    .ColWidth(5) = 900
    .Col = 1
    .CellBackColor = RGB(180, 170, 160)
    .CellAlignment = 0
    .Text = "Chq_Date"
    .Col = 2
    .CellBackColor = RGB(180, 170, 160)
    .CellAlignment = 1
    .Text = "Chq_Bank"
    .Col = 3
    .CellBackColor = RGB(180, 170, 160)
    .CellAlignment = 0
    .Text = "Code"
    .Col = 4
    .CellBackColor = RGB(180, 170, 160)
    .CellAlignment = 1
    .Text = "Dep_Bank"
    .Col = 5
    .CellBackColor = RGB(180, 170, 160)
    .CellAlignment = 2
    .Text = "Amount"
    .Row = 0
    .Col = 1
  
  End With
End Sub

Private Sub lstAcctCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtTtlDesc.SetFocus
End Sub

Private Sub Msflexgrid1_dblclick()
 Dim i
 Dim j
 Dim X
 Dim Y, Z, U As Long
 Dim MchequeNo, MChequeBank, MBankCode, MDepBank
 Dim MchequeDt As Date
 Dim mamount As Currency
 
 X = MSFlexGrid1.Rows
 If X > 1 Then
 If MSFlexGrid1.Row = MSFlexGrid1.TopRow Then
  Exit Sub
 Else
   i = MsgBox(" Are you sure .. ! You want to Remove this transaction", vbInformation + vbYesNo)
    If i = vbYes Then
     With MSFlexGrid1
        j = .Row
        .Col = 0
        txtChequeNO = .Text
        .Col = 1
        txtchequedt = .Text
        .Col = 2
        txtChequeBank = .Text
        .Col = 3
        MBankCode = .Text
        .Col = 4
        MDepBank = .Text
        .Col = 5
        txtamount = .Text
        
        Sqlqry = "Select Bank_code,Bank_name from Bank_mas Where Bank_Code='" & MBankCode & "' "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

        If rs.RecordCount = 0 Then
            MsgBox "No Matching Record found in the Bank Register"
        Else
            cboBankCode.Text = rs!bank_code & " : " & rs!BANK_NAME
        End If
        
                                   
         lblTtlAmount.Caption = Val(lblTtlAmount.Caption) - Val(txtamount)
         Y = Val(lblTtlAmount.Caption)
        .RemoveItem (j)
        
        Sqlqry1 = "Delete * from DumpRPT2 where Cheque_no='" & txtChequeNO & "' and Cheque_dt =#" & DateValue(Format(txtchequedt.TextWithMask, "dd/mm/yyyy")) & "# and Amount =" & Val(txtamount) & ""
        ws.BeginTrans
        db.Execute Sqlqry1
        ws.CommitTrans
        
        
     End With
    End If
   End If
  End If
End Sub

Private Sub txtamount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtChequeNO.SetFocus
End Sub

Private Sub txtAmount_LostFocus()
 Dim cur As String
 Dim con As Currency
 
  If ValidateData = True Then
    If Val(txtamount.Text) > Val(txtTtlAmount.Text) Then
      MsgBox " Entered Amount Greater than Total Amount"
      txtamount.SetFocus
    Exit Sub
    End If
      
    If Val(txtTtlAmount.Text) = Val(lblTtlAmount) Then
      CmdSave.SetFocus
      Exit Sub
    End If
    
    
            cur = ""
            con = 1
    
       If CboCurrency.Text = "USD" Then
          cur = "USD"
          con = Val(Trim(txtConvRate.Text))
           
        Else
          cur = "DHS"
          con = 1
        End If
        
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = " Select * from DUMPRPT2"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If txtamount.Text = 0 Then
      Exit Sub
      txtamount.SetFocus
    End If
    If rs.RecordCount = 0 Then
       Sqlqry = " Insert into DUMPRPT2 values('" & lblVoucNo & "','PRT','" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & Mid(cboBankCode, 1, 6) & "','" _
                                     & Mid(cboBankCode, 10, 35) & "','" _
                                     & txtChequeNO & "',#" _
                                     & Format(txtchequedt.TextWithMask, "dd/mm/yyyy") & "#,'" _
                                     & Trim(txtChequeBank) & "','" _
                                     & Trim(CboCurrency) & "'," _
                                     & Val(Trim(txtConvRate)) & "," _
                                     & Val(Trim(txtamount)) & "," _
                                     & Val(txtamount) * con & ")"

        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        Sqlqry1 = "Select * from DUMPRPT2"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlexGrid1.Clear
            Exit Sub
        Else
            Flexitems
            rs.MoveFirst
            Do Until rs.EOF
              MSFlexGrid1.AddItem rs!CHEQUE_NO & Chr(9) & Format(rs!Cheque_Dt, "DD/MM/YYYY") & Chr(9) & rs!cheque_Bank & Chr(9) & rs!bank_code & Chr(9) & rs!BANK_NAME & Chr(9) & rs!tra_amount
              rs.MoveNext
            Loop
        End If
            lblTtlAmount = Val(txtamount.Text)
            lblTtlAmount.Alignment = 1
            If Val(txtTtlAmount) = Val(txtamount.Text) Then
             CmdSave.SetFocus
            Else
             txtChequeNO.Text = Val(txtChequeNO) + 1
             txtChequeNO.SetFocus
            End If
        fdate = DateValue(Format(txtchequedt.TextWithMask, "dd/mm/yyyy"))
            
      Else
          
         rs.MoveFirst
         
         X = 0
         Do Until rs.EOF
          X = X + rs!tra_amount
          B = B + 1
           If B = 1 Then
            ldate = DateValue(Format(txtchequedt.TextWithMask, "dd/mm/yyyy"))
            B = 2
           End If
          rs.MoveNext
         Loop
      
       If Val(txtTtlAmount.Text) >= X + Val(txtamount.Text) Then
        
      Sqlqry = " Insert into DUMPRPT2 values('" & lblVoucNo & "','PRT','" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & Mid(cboBankCode, 1, 6) & "','" _
                                     & Mid(cboBankCode, 10, 35) & "','" _
                                     & txtChequeNO & "',#" _
                                     & Format(txtchequedt.TextWithMask, "dd/mm/yyyy") & "#,'" _
                                     & Trim(txtChequeBank) & "','" _
                                     & Trim(CboCurrency) & "'," _
                                     & Val(Trim(txtConvRate)) & "," _
                                     & Val(Trim(txtamount)) & "," _
                                     & Val(txtamount) * con & ")"

        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
          
        Sqlqry1 = "Select * from DUMPRPT2"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
           MSFlexGrid1.Clear
           Exit Sub
        Else
           Flexitems
           Y = 0
           rs.MoveFirst
           Do Until rs.EOF
            MSFlexGrid1.AddItem rs!CHEQUE_NO & Chr(9) & Format(rs!Cheque_Dt, "DD/MM/YYYY") & Chr(9) & rs!cheque_Bank & Chr(9) & rs!bank_code & Chr(9) & rs!BANK_NAME & Chr(9) & rs!Amount
            Y = Y + Val(rs!tra_amount)
            rs.MoveNext
            Loop
        End If
             lblTtlAmount = Y
             lblTtlAmount.Alignment = 1
             If Val(txtTtlAmount.Text) = Y Then
               CmdSave.SetFocus
             Else
               txtChequeNO.Text = Val(txtChequeNO) + 1
               a = DateDiff("M", fdate, ldate)
               txtchequedt.Text = Format(DateAdd("M", a, Format(txtchequedt.TextWithMask, "dd/mm/yyyy")), "DD/MM/YYYY")
               txtChequeNO.SetFocus
             End If
         Else
             MsgBox "Entered Amount is more than Total Amount"
             txtamount.SetFocus
             Exit Sub
         End If
       End If
 End If
 
End Sub
Private Sub txtChequeBank_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboBankCode.SetFocus
End Sub
Private Sub txtChequeDt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtChequeBank.SetFocus
End Sub

Private Sub txtchequedt_LostFocus()
Dim a As Date
Dim B As Date

  If IsDate(txtchequedt.TextWithMask) = False Then
      MsgBox "Invalid cheque Date ", vbInformation, "Invalid Entry"
      txtchequedt.SetFocus
      SendKeys "{Home} + {End}"
  End If

a = Format(txtchequedt.TextWithMask, "dd/mm/yyyy")
B = Now()

If DateValue(a) <= DateValue(B) Then
  MsgBox "Cheque date cannot be lesser than or equal to current date"
  txtchequedt.SetFocus
  Exit Sub
End If
End Sub

Private Sub txtChequeNO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtchequedt.SetFocus
End Sub
Private Sub txtConvRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRecdFrom.SetFocus
End Sub
Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboCurrency.SetFocus
End Sub
Private Sub txtdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtamount.SetFocus
End Sub
Private Sub txtpaidto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtTtlDesc.SetFocus
End Sub

Private Sub txtdate_LostFocus()
  If IsDate(txtdate.TextWithMask) = False Then
      MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
      txtdate.SetFocus
      SendKeys "{Home} + {End}"
  End If
End Sub

Private Sub txtrecdfrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstAcctCode.SetFocus
End Sub
Private Sub txtTtlAmount_KeyPress(KeyAscii As Integer)
    If CboCurrency = "USD" Then
     If KeyAscii = 13 Then txtConvRate.SetFocus
    Else
     If KeyAscii = 13 Then txtRecdFrom.SetFocus
    End If
End Sub
Private Sub txtTtlAmount_LostFocus()
txtamount.Text = Val(txtTtlAmount.Text)
End Sub
Private Sub txtTtlDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtChequeNO.SetFocus
End Sub

Private Function textclear()
     txtChequeNO.Text = ""
     txtChequeBank.Text = ""
     txtchequedt.TextWithMask = Format(Now, "dd/mm/yyyy")
     txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
     cboBankCode.Clear
     txtRecdFrom.Text = ""
     txtTtlDesc.Text = ""
     txtamount.Text = ""
     lblTtlAmount.Caption = ""
     B = 0
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from DUMPRPT2"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     Flexitems
     txtTtlAmount.SetFocus
     PopulateBankCodes
End Function

Public Sub PopulateBankCodes()
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Select * from Bank_mas order by bank_code"
 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

 cboBankCode.Clear

  If rs.RecordCount = 0 Then
      MsgBox "No Records found in the Bank Register"
  Else
      rs.MoveFirst
   Do Until rs.EOF
      cboBankCode.AddItem rs!bank_code & " : " & rs!BANK_NAME
      rs.MoveNext
   Loop
  End If

End Sub
