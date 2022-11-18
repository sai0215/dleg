VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmPDCPaymentAddition1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "PDC Issue - Addition"
   ClientHeight    =   8775
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   12060
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
   ScaleWidth      =   12060
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
      Height          =   8655
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   11655
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   3480
         Width           =   2535
         Begin VB.CommandButton cmdSupplier 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Supplier"
            Height          =   375
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdAgency 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Agency"
            Height          =   375
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.TextBox txtConvRate 
         BackColor       =   &H80000018&
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   7800
         TabIndex        =   32
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cboCurrency 
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtChequeNO 
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox cboBankCode 
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtTtlAmount 
         BackColor       =   &H80000018&
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   4680
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtPaidTo 
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   2040
         Width           =   4575
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
         ForeColor       =   &H80000012&
         Height          =   375
         Left            =   9960
         TabIndex        =   12
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtDesc 
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
         ForeColor       =   &H80000012&
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   3840
         Width           =   4215
      End
      Begin VB.ListBox lstAcctCode 
         BackColor       =   &H80000018&
         ForeColor       =   &H00404040&
         Height          =   1500
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   3840
         Width           =   5055
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
         Left            =   5520
         MaskColor       =   &H0080C0FF&
         Picture         =   "frmPDCPaymentAddition1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7680
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFF80&
         Caption         =   "<<&Back"
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
         Left            =   6720
         MaskColor       =   &H0080C0FF&
         Picture         =   "frmPDCPaymentAddition1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7680
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Save"
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
         Left            =   4320
         MaskColor       =   &H0080C0FF&
         Picture         =   "frmPDCPaymentAddition1.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7680
         Width           =   1215
      End
      Begin VB.TextBox txtTtlDesc 
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1680
         MaxLength       =   150
         TabIndex        =   7
         Top             =   2640
         Width           =   9735
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   9120
         Top             =   2880
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1575
         Left            =   240
         TabIndex        =   17
         Top             =   5400
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   5
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   9613530
         BackColorBkg    =   8421376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PVMaskEditLib.PVMaskEdit txtdate 
         Height          =   375
         Left            =   4680
         TabIndex        =   0
         Top             =   240
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
         Left            =   4680
         TabIndex        =   4
         Top             =   1440
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
      Begin VB.Label lblConvRate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Conv. Rate"
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   6540
         TabIndex        =   33
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Currency"
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   675
         TabIndex        =   31
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cheque Dt."
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   3420
         TabIndex        =   30
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cheque No."
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   390
         TabIndex        =   29
         Top             =   1560
         Width           =   1230
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Bank Name"
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   6375
         TabIndex        =   28
         Top             =   1560
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Voucher No."
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   285
         TabIndex        =   27
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Payee Name"
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   195
         TabIndex        =   26
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Total Amount"
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   3240
         TabIndex        =   25
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Date"
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   3990
         TabIndex        =   24
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   5520
         TabIndex        =   23
         Top             =   3480
         Width           =   1320
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   10080
         TabIndex        =   22
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label lblVoucNo 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   1095
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
         Height          =   375
         Left            =   9960
         TabIndex        =   20
         Top             =   6960
         Width           =   1455
      End
      Begin VB.Label lblTtlAmt 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   8280
         TabIndex        =   19
         Top             =   7080
         Width           =   1530
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Description"
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   375
         TabIndex        =   18
         Top             =   2760
         Width           =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404080&
         BorderWidth     =   2
         X1              =   11640
         X2              =   0
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   0
         X2              =   11640
         Y1              =   7560
         Y2              =   7560
      End
   End
End
Attribute VB_Name = "frmPDCPaymentAddition1"
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
Dim X
Dim y As Long
Dim Z
Dim i
Dim j
Private Sub cboBankCode_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtPaidTo.SetFocus
End Sub

Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtTtlAmount.SetFocus
End Sub

Private Sub cboCurrency_LostFocus()
    If cboCurrency.Text = "USD" Then
     lblConvRate.Visible = True
     txtConvRate.Visible = True
     txtConvRate.Text = ""
     txtConvRate.TabIndex = 3
    Else
     lblConvRate.Visible = False
     txtConvRate.Visible = False
     txtConvRate.Text = 1
     txtConvRate.TabIndex = 17
    End If
End Sub
Private Sub cmdBack_Click()
 Unload Me
End Sub
Private Sub cmdClear_Click()
     textclear
     lblTtlAmount.Caption = ""
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from dumppmt1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     MSFlexGrid1.Clear
     txtTtlAmount.SetFocus
End Sub

Private Sub CmdSave_Click()
Dim TTAmt As Currency
Dim ctype As String

If ValidateData = True Then
 
            cur = ""
            con = 1
    
       If cboCurrency.Text = "USD" Then
          cur = "USD"
          con = Val(Trim(txtConvRate.Text))
           
        Else
          cur = "DHS"
          con = 1
        End If
    
   If Val(txtTtlAmount.Text) = Val(lblTtlAmount.Caption) Then
         
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
               Sqlqry = " Insert into ppmt_mas values('" & lblVoucNo & "','PPT','" _
                                             & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                             & UCase(Trim(txtPaidTo)) & "','" _
                                             & Mid(cboBankCode, 1, 6) & "','" _
                                             & Mid(cboBankCode, 10, 25) & "','" _
                                             & UCase(Trim(txtChequeNO)) & "','" _
                                             & Format(txtchequedt.TextWithMask, "dd/mm/yyyy") & "','','" _
                                             & findfirstfixup(UCase(Trim(txtTtlDesc))) & "','" _
                                             & Trim(cboCurrency) & "'," _
                                             & Val(Trim(txtConvRate)) & "," _
                                             & Val(txtTtlAmount) & "," _
                                             & Val(txtTtlAmount) * con & ",'N')"
               ws.BeginTrans
               db.Execute (Sqlqry)
               ws.CommitTrans
        
            Sqlqry1 = "Select * from dumppmt1"
            Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
             If rs.RecordCount = 0 Then
                 MsgBox " Transactions are not recorded"
                 Exit Sub
              Else
                 rs.MoveFirst
                 Do Until rs.EOF
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 Sqlqry3 = "Insert into ppmt_tra values('" & rs!vouc_no & "','" & rs!vouc_type & "','" _
                                             & Trim(rs!tDate) & "','" _
                                             & rs!acct_code & "','" _
                                             & findfirstfixup(rs!acct_name) & "','" _
                                             & findfirstfixup(rs!Description) & "','" _
                                             & Format(txtchequedt.TextWithMask, "dd/mm/yyyy") & "','','" _
                                             & rs!tcurrency & "'," _
                                             & rs!tconvertion & "," _
                                             & rs!tra_amount & "," _
                                             & rs!Amount & ")"
        
                    ws.BeginTrans
                    db.Execute (Sqlqry3)
                    ws.CommitTrans
                  rs.MoveNext
                 Loop
               End If
     
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = "Update docu_mas set doc_no='" & lblVoucNo & "' where doc_type='PPT'"
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
         ctype = cboCurrency.Text
        textclear
        lblVoucNo = lblVoucNo + 1
   Else
        MsgBox "Total amount is not equal to entered amount"
        Exit Sub
   End If
  
   End If
    MsgBox " Record is inserted", vbInformation, "Status"
  
  Dim X
   X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
  If X = vbYes Then
   If ctype = "DHS" Then
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\ppmtVou.rpt"
        CrystalReport1.SelectionFormula = "{ppmt_tra.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
        CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtTtlAmount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
     Else
             CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\ppmtVou.rpt"
        CrystalReport1.SelectionFormula = "{ppmt_tra.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
        CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(txtTtlAmount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1

     End If
  End If
 
End Sub

Private Sub cmdAgency_Click()

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Agndtls order by agentname"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    
    lstAcctCode.Clear
    
    If rs.RecordCount = 0 Then
        MsgBox "No Records found in the Agency Register"
    Else
       rs.MoveFirst
       Do Until rs.EOF
          lstAcctCode.AddItem rs!agentname
          rs.MoveNext
       Loop
    End If
    
    LSTSUP = 0
    LSTAGN = 1
End Sub

Private Sub cmdSupplier_Click()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Supp_fin order by Supp_no"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    
    lstAcctCode.Clear
    
    If rs.RecordCount = 0 Then
        MsgBox "No Records found in the Suppliers Register"
    Else
       rs.MoveFirst
       Do Until rs.EOF
          lstAcctCode.AddItem rs!Supp_no & "  :  " & rs!Supp_name
          rs.MoveNext
       Loop
    End If
    LSTSUP = 1
    LSTAGN = 0
End Sub

Private Sub Form_Load()
 txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
 txtchequedt.TextWithMask = Format(Now, "dd/mm/yyyy")
 cboCurrency.AddItem "DHS"
 cboCurrency.AddItem "USD"
 
 lblConvRate.Visible = False
 txtConvRate.Visible = False
 
 LSTSUP = 1
 LSTAGN = 0
 
 AutoIncrementVoucher
 PopulateAcctSuppCust
 PopulateBankCodes
 Flexitems
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
  Sqlqry = "delete * from dumppmt1"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
    
End Sub

Private Sub AutoIncrementVoucher()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select DOC_no from DOCU_MAS WHERE DOC_TYPE='PPT'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
       MsgBox "Document type 'PPT' not found"
       Exit Sub
    Else
       lblVoucNo = Val(rs!doc_no) + 1
    End If
End Sub

Private Sub PopulateAcctSuppCust()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Supp_fin order by Supp_no"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    
    lstAcctCode.Clear
    
    If rs.RecordCount = 0 Then
        MsgBox "No Records found in the Suppliers Register"
    Else
       rs.MoveFirst
       Do Until rs.EOF
          lstAcctCode.AddItem rs!Supp_no & "  :  " & rs!Supp_name
          rs.MoveNext
       Loop
    End If
    LSTSUP = 1
    LSTAGN = 0
End Sub

Private Function ValidateData()

ValidateData = False
If txtdate = "" Or IsDate(txtdate.TextWithMask) = False Then
  MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
  txtdate.SetFocus
  SendKeys "{Home} + {End}"
  Exit Function
ElseIf txtTtlAmount.Text = "" Or IsNumeric(txtTtlAmount) = False Then
  MsgBox "Invalid Total Amount", vbInformation, "Invalid Entry"
  txtTtlAmount.SetFocus
  Exit Function
ElseIf txtPaidTo.Text = "" Or IsNumeric(txtTtlAmount) = False Then
  MsgBox "Invalid Name of Payee", vbInformation, "Invalid Entry"
  txtPaidTo.SetFocus
  Exit Function
ElseIf lstAcctCode.SelCount = 0 Then
  MsgBox "Select Account/Supplier/Customer Code from list box", vbInformation, "Invalid Entry"
  lstAcctCode.SetFocus
  Exit Function
ElseIf txtdesc.Text = "" Or IsNumeric(txtdesc) = True Then
  MsgBox "Invalid Description", vbInformation, "Invalid Entry"
  txtdesc.SetFocus
  Exit Function
ElseIf txtAmount.Text = "" Or IsNumeric(txtAmount) = False Then
  MsgBox "Invalid Amount", vbInformation, "Invalid Entry"
  txtAmount.SetFocus
  Exit Function
ElseIf txtChequeNO.Text = "" Then
  MsgBox "Invalid Cheque No.", vbInformation, "Invalid Entry"
  txtChequeNO.SetFocus
  Exit Function
ElseIf txtchequedt.Text = "" Or IsDate(txtchequedt.TextWithMask) = False Then
  MsgBox "Invalid Cheque Date", vbInformation, "Invalid Entry"
  txtchequedt.SetFocus
  Exit Function
ElseIf txtConvRate.Text = "" Then
  MsgBox "Enter Convertion Rate - - cannot be zero", vbInformation, "Invalid Entry"
  txtConvRate.SetFocus
  Exit Function
ElseIf cboBankCode.Text = "" Or IsNumeric(txtAmount) = False Then
  MsgBox "Invalid Amount", vbInformation, "Invalid Entry"
  txtAmount.SetFocus
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
    .Cols = 4
    .Col = 0
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Account Code"
    .ColAlignment(0) = 0
    .ColWidth(0) = 1100
    .ColWidth(1) = 3250
    .ColWidth(2) = 6000
    .ColWidth(3) = 900
    .Col = 1
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Account Name"
    .Col = 2
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Description"
    .Col = 3
    .CellBackColor = RGB(180, 170, 160)
    .Text = "Amount"
    .Row = 0
    .Col = 1
  
  End With
End Sub
Private Sub lstAcctCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdesc.SetFocus
End Sub
Private Sub Msflexgrid1_dblclick()
 Dim i
 Dim j
 Dim X
 Dim y, Z, U
 Dim txtaccode, txtacname
 
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
        txtaccode = .Text
        If .Text = "AGNC" Then
               
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = "Select * from Agndtls order by agentname"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            
            lstAcctCode.Clear
            
            If rs.RecordCount = 0 Then
                MsgBox "No Records found in the Agency Register"
            Else
               rs.MoveFirst
               Do Until rs.EOF
                  lstAcctCode.AddItem rs!agentname
                  rs.MoveNext
               Loop
            End If
            
            LSTSUP = 0
            LSTAGN = 1

        Else
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = "Select * from Supp_fin order by Supp_no"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            
            lstAcctCode.Clear
            
            If rs.RecordCount = 0 Then
                MsgBox "No Records found in the Suppliers Register"
            Else
               rs.MoveFirst
               Do Until rs.EOF
                  lstAcctCode.AddItem rs!Supp_no & "  :  " & rs!Supp_name
                  rs.MoveNext
               Loop
            End If
            LSTSUP = 1
            LSTAGN = 0
       End If
        .Col = 1
        txtacname = .Text
        .Col = 2
        txtdesc = .Text
        .Col = 3
        txtAmount = .Text
                            
          Sqlqry1 = "Select Supp_no,Supp_name from supp_Fin where supp_no='" & Trim(txtaccode) & "' order by supp_no"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                lstAcctCode.Text = rs1!Supp_no & "  :  " & rs1!Supp_name
               Else
                 Sqlqry2 = "Select AGENTNAME from AGNDTLS where AGENTNAME='" & Trim(txtacname) & "' order by aGENTNAME"
                 Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                 If rs2.RecordCount <> 0 Then
                  lstAcctCode.Text = rs2!agentname
                 Else
                  MsgBox "Selected Code not found in Account/Supplier/Agency list"
                 End If
                End If
      
            
        lblTtlAmount.Caption = Val(lblTtlAmount.Caption) - Val(txtAmount)
        
        .RemoveItem (j)
        
        Sqlqry1 = "Delete * from dumppmt1 where Acct_Code='" & txtaccode & "' and description ='" & txtdesc & "' and amount =" & Val(txtAmount) & ""
        ws.BeginTrans
        db.Execute Sqlqry1
        ws.CommitTrans
        
        
     End With
    End If
   End If
  End If
End Sub
Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtAmount_LostFocus()
   
   Dim accd As String
   Dim acname As String
   
  If LSTSUP = 1 Then
     accd = Mid(lstAcctCode, 1, 4)
     acname = Mid(lstAcctCode, 10, 35)
  Else
     accd = "AGNC"
     acname = Trim(lstAcctCode.Text)
  End If
  
 
    cur = ""
    con = 1
 
  If cboCurrency.Text = "USD" Then
      cur = "USD"
      con = Val(Trim(txtConvRate.Text))
       
  Else
      cur = "DHS"
      con = 1
  End If
 
   If ValidateData = True Then
    If Val(txtAmount.Text) > Val(txtTtlAmount.Text) Then
      MsgBox " Entered Amount Greater than Total Amount"
      txtAmount.SetFocus
    Exit Sub
    End If
      
       
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = " select * from dumppmt1"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If txtAmount.Text = 0 Then
      Exit Sub
      txtAmount.SetFocus
    End If
    If rs.RecordCount = 0 Then
       Sqlqry = " Insert into dumppmt1 values('" & lblVoucNo & "','PPT','" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & accd & "','" _
                                     & findfirstfixup(acname) & "','" _
                                     & findfirstfixup(UCase(Trim(txtdesc))) & "','" _
                                     & Trim(cboCurrency) & "'," _
                                     & Val(Trim(txtConvRate)) & "," _
                                     & Val(Trim(txtAmount)) & "," _
                                     & Val(txtAmount) * con & ")"

        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        Sqlqry1 = "select * from dumppmt1"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
            MSFlexGrid1.Clear
            Exit Sub
        Else
            Flexitems
            rs.MoveFirst
            Do Until rs.EOF
              MSFlexGrid1.AddItem rs!acct_code & Chr(9) & rs!acct_name & Chr(9) & rs!Description & Chr(9) & rs!tra_amount
              rs.MoveNext
            Loop
        End If
            lblTtlAmount = Val(txtAmount.Text)
            lblTtlAmount.Alignment = 1
            If Val(txtTtlAmount) = Val(txtAmount.Text) Then
            cmdSave.SetFocus
            Else
            lstAcctCode.SetFocus
            End If
      Else
        rs.MoveFirst
        X = 0
         Do Until rs.EOF
          X = X + rs!tra_amount
          rs.MoveNext
         Loop
      
       If Val(txtTtlAmount.Text) >= X + Val(txtAmount.Text) Then
       Sqlqry = " Insert into dumppmt1 values('" & lblVoucNo & "','PPT','" _
                                     & Format(txtdate, "dd/mm/yyyy") & "','" _
                                     & accd & "','" _
                                     & findfirstfixup(acname) & "','" _
                                     & findfirstfixup(UCase(Trim(txtdesc))) & "','" _
                                     & Trim(cboCurrency) & "'," _
                                     & Val(Trim(txtConvRate)) & "," _
                                     & Val(Trim(txtAmount)) & "," _
                                     & Val(txtAmount) * Val(txtConvRate) & ")"
   
   ws.BeginTrans
          db.Execute (Sqlqry)
          ws.CommitTrans
          
          Sqlqry1 = "Select * from dumppmt1"
          Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
           
           If rs.RecordCount = 0 Then
             MSFlexGrid1.Clear
             Exit Sub
           Else
             Flexitems
             y = 0
             rs.MoveFirst
             Do Until rs.EOF
               MSFlexGrid1.AddItem rs!acct_code & Chr(9) & rs!acct_name & Chr(9) & rs!Description & Chr(9) & rs!tra_amount
               y = y + rs!tra_amount
               rs.MoveNext
             Loop
           End If
             
             lblTtlAmount = y
             lblTtlAmount.Alignment = 1
             If Val(txtTtlAmount.Text) = y Then
               cmdSave.SetFocus
             Else
               lstAcctCode.SetFocus
             End If
         Else
             MsgBox "Entered Amount is more than Total Amount"
             txtAmount.SetFocus
             Exit Sub
         End If
       End If
 End If
 txtdesc.Text = Trim(txtTtlDesc.Text)
End Sub
Private Sub txtChequeDt_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboBankCode.SetFocus
End Sub
Private Sub txtchequedt_LostFocus()
    If IsDate(txtchequedt.TextWithMask) = False Then
          MsgBox "Invalid Cheque Date ", vbInformation, "Invalid Entry"
          txtchequedt.SetFocus
          SendKeys "{Home} + {End}"
    End If
End Sub
Private Sub txtChequeNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtchequedt.SetFocus
End Sub

Private Sub txtConvRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtChequeNO.SetFocus
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboCurrency.SetFocus
End Sub

Private Sub txtdate_LostFocus()
If IsDate(txtdate.TextWithMask) = False Then
      MsgBox "Invalid Date ", vbInformation, "Invalid Entry"
      txtdate.SetFocus
      SendKeys "{Home} + {End}"
End If
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAmount.SetFocus
End Sub

Private Sub txtpaidto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtTtlDesc.SetFocus
End Sub

Private Sub txtTtlAmount_KeyPress(KeyAscii As Integer)
If cboCurrency = "USD" Then
 If KeyAscii = 13 Then txtConvRate.SetFocus
Else
 If KeyAscii = 13 Then txtChequeNO.SetFocus
End If
End Sub

Private Sub txtTtlAmount_LostFocus()
txtAmount.Text = Val(txtTtlAmount.Text)
End Sub

Private Sub txtTtlDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstAcctCode.SetFocus
txtdesc.Text = Trim(txtTtlDesc.Text)
End Sub

Private Function textclear()
     txtChequeNO.Text = ""
     txtchequedt.Text = ""
     txtConvRate.Text = ""
     txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
     cboBankCode.Clear
     txtPaidTo.Text = ""
     txtTtlDesc.Text = ""
     txtdesc.Text = ""
     txtAmount.Text = ""
     lblTtlAmount.Caption = ""
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from dumppmt1"
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
      MsgBox "No Records found in the Bank Master"
 Else
      rs.MoveFirst
      Do Until rs.EOF
      cboBankCode.AddItem rs!bank_code & " : " & rs!BANK_NAME
      rs.MoveNext
 Loop
 End If

End Sub

Private Sub txtTtlDesc_LostFocus()
txtdesc.Text = Trim(txtTtlDesc.Text)
End Sub
