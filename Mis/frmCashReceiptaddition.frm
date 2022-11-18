VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmCashReceiptAddition 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   Caption         =   "Cash Receipt Addition"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   240
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cash Receipt Addition"
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
      Height          =   8415
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   11655
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
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtConvRate 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   375
         Left            =   7845
         TabIndex        =   26
         Top             =   1320
         Width           =   855
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   3360
         Width           =   3615
         Begin VB.CommandButton cmdAgency 
            BackColor       =   &H00C0FFC0&
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
            Height          =   375
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmdSupplier 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Supplier"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdAcCode 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Gen. A/C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   0
            Width           =   1095
         End
      End
      Begin PVMaskEditLib.PVMaskEdit txtdate 
         Height          =   375
         Left            =   4680
         TabIndex        =   0
         Top             =   480
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
      Begin VB.TextBox txtTtlAmount 
         BackColor       =   &H80000018&
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
         Height          =   405
         Left            =   4680
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtPaidTo 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   2040
         Width           =   4575
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10080
         TabIndex        =   9
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H80000018&
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
         Left            =   4560
         TabIndex        =   8
         Top             =   3720
         Width           =   5295
      End
      Begin VB.ListBox lstAcctCode 
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
         Height          =   1500
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   3720
         Width           =   4335
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
         Left            =   4560
         Picture         =   "frmCashReceiptaddition.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7440
         Width           =   975
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
         Height          =   855
         Left            =   5520
         Picture         =   "frmCashReceiptaddition.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7440
         Width           =   975
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
         Left            =   3600
         Picture         =   "frmCashReceiptaddition.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7440
         Width           =   975
      End
      Begin VB.TextBox txtTtlDesc 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         MaxLength       =   150
         TabIndex        =   4
         Top             =   2640
         Width           =   9615
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   7320
         Top             =   7800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         Top             =   5280
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   5
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483647
         BackColorFixed  =   -2147483647
         BackColorSel    =   -2147483624
         BackColorBkg    =   8421376
         TextStyle       =   1
         TextStyleFixed  =   1
      End
      Begin VB.Label lblcurtype 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   9480
         TabIndex        =   30
         Top             =   6960
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   3285
         TabIndex        =   29
         Top             =   1440
         Width           =   1380
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   720
         TabIndex        =   28
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label lblConvRate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   6585
         TabIndex        =   27
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         X1              =   11640
         X2              =   0
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00400040&
         BorderWidth     =   2
         X1              =   11640
         X2              =   0
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   300
         TabIndex        =   23
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         Left            =   4080
         TabIndex        =   22
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Description"
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
         Height          =   345
         Left            =   4560
         TabIndex        =   21
         Top             =   3360
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Amount "
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
         Left            =   10200
         TabIndex        =   20
         Top             =   3360
         Width           =   1020
      End
      Begin VB.Label lblVoucNo 
         Alignment       =   2  'Center
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
         ForeColor       =   &H80000001&
         Height          =   495
         Left            =   1680
         TabIndex        =   19
         Top             =   480
         Width           =   1215
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
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   10080
         TabIndex        =   18
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Label lblTtlAmt 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
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
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   8040
         TabIndex        =   17
         Top             =   6960
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Description "
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
         Left            =   405
         TabIndex        =   16
         Top             =   2760
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmCashReceiptAddition"
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
Dim y
Dim Z
Dim i
Dim j

Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtTtlAmount.SetFocus
End Sub

Private Sub cboCurrency_LostFocus()
  If cboCurrency.Text = "USD" Then
     lblConvRate.Visible = True
     txtConvRate.Visible = True
     lblcurtype.Caption = "USD"
     txtConvRate.Text = ""
     txtConvRate.TabIndex = 3
     
    Else
     lblConvRate.Visible = False
     txtConvRate.Visible = False
     lblcurtype.Caption = "DHS"
     txtConvRate.Text = 1
     txtConvRate.TabIndex = 16
    End If
End Sub

Private Sub cmdAcCode_Click()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Acct_mas order by acct_code"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    
    lstAcctCode.Clear
    
    If rs.RecordCount = 0 Then
        MsgBox "No Records found in the Accounts Master"
    Else
       rs.MoveFirst
       Do Until rs.EOF
          lstAcctCode.AddItem rs!acct_code & "  :  " & rs!acct_name
          rs.MoveNext
       Loop
    End If
    LSTSUP = 0
    LSTAC = 1
    LSTAGN = 0
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
    LSTAC = 0
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
    LSTAC = 0
    LSTAGN = 0
End Sub

Private Sub Msflexgrid1_dblclick()
 Dim i As Long
 Dim j As Long
 Dim X As Long
 Dim y, Z, U As Long
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
        If Len(.Text) = 6 Then
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = "Select * from Acct_mas order by acct_code"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            
            lstAcctCode.Clear
            
            If rs.RecordCount = 0 Then
                MsgBox "No Records found in the Accounts Master"
            Else
               rs.MoveFirst
               Do Until rs.EOF
                  lstAcctCode.AddItem rs!acct_code & "  :  " & rs!acct_name
                  rs.MoveNext
               Loop
            End If
            LSTSUP = 0
            LSTAC = 1
            LSTAGN = 0
            
          ElseIf .Text = "AGNC" Then
     
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
            LSTAC = 0
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
            LSTAC = 0
            LSTAGN = 0
       End If
        .Col = 1
        txtacname = .Text
        .Col = 2
        txtdesc = .Text
        .Col = 3
        txtAmount = .Text
                   
        Sqlqry = "Select acct_Code,acct_name from Acct_mas where acct_code='" & txtaccode & "' order by acct_code"
         Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
                  lstAcctCode.Text = rs!acct_code & "  :  " & rs!acct_name
           Else
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
            End If
            
            
          lblTtlAmount.Caption = Round(Val(lblTtlAmount.Caption) - Val(txtAmount), 2)
        
        .RemoveItem (j)
        
        Sqlqry1 = "Delete * from dumCrpt1 where Acct_Code='" & txtaccode & "' and description ='" & txtdesc & "' and Tra_amount =" & Val(txtAmount) & ""
        ws.BeginTrans
        db.Execute Sqlqry1
        ws.CommitTrans
        
        
     End With
    End If
   End If
  End If
End Sub
Private Sub cmdBack_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub
Private Sub cmdClear_Click()
     txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
     txtTtlAmount.Text = ""
     txtPaidTo.Text = ""
     txtTtlDesc.Text = ""
     lblcurtype.Caption = ""
     cboCurrency.ListIndex = -1
     lstAcctCode.Clear
     txtdesc.Text = ""
     txtAmount.Text = ""
     lblTtlAmount.Caption = ""
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from dumCrpt1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     MSFlexGrid1.Clear
     txtTtlAmount.SetFocus
End Sub
Private Sub CmdSave_Click()
Dim TTAmt As Currency
Dim ctype As String
cur = ""
cod = ""
con = 1

 If ValidateData = True Then
   If Val(txtTtlAmount.Text) = lblTtlAmount Then
    If cboCurrency.Text = "USD" Then
      cur = "USD"
      cod = "103002"
      con = Val(Trim(txtConvRate.Text))
       
    Else
      cur = "DHS"
      cod = "103001"
      con = 1
    End If
    
         
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       Sqlqry = " Insert into crpt_mas values('" & lblVoucNo & "','CRT','" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & UCase(Trim(txtPaidTo)) & "','" _
                                     & findfirstfixup(UCase(Trim(txtTtlDesc))) & "','" _
                                     & cur & "'," _
                                     & con & "," _
                                     & Trim(txtTtlAmount) & "," & Val(Trim(txtTtlAmount)) * con & ", '" & Val(cod) & "','N')"
       ws.BeginTrans
       db.Execute (Sqlqry)
       ws.CommitTrans
        
    Sqlqry1 = "Select * from dumCrpt1"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount = 0 Then
         MsgBox " Transactions are not recorded"
         Exit Sub
      Else
         rs.MoveFirst
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         Sqlqry3 = "Insert into crpt_tra values('" & rs!vouc_no & "','" & rs!vouc_type & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & rs!acct_code & "','" _
                                     & findfirstfixup(rs!acct_name) & "','" _
                                     & findfirstfixup(rs!Description) & "','" _
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
     Sqlqry = "Update docu_mas set doc_no='" & lblVoucNo & "' where doc_type='CRT'"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
      ctype = cboCurrency.Text
     textclear
     lblVoucNo = lblVoucNo + 1
     
   Else
   MsgBox "Total Amount is not equal to Entered Amount"
   Exit Sub
   End If
  End If
  MsgBox " Record is inserted", vbInformation, "Status"
  textclear
  
  Dim X
  X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
  If X = vbYes Then
   If ctype = "DHS" Then
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\CrptVou.rpt"
        CrystalReport1.SelectionFormula = "{crpt_tra.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
        CrystalReport1.Formulas(0) = "xxx1='" & Inwords(Val(txtTtlAmount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    Else
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\CrptVou.rpt"
        CrystalReport1.SelectionFormula = "{crpt_tra.Vouc_no}=" & Val(lblVoucNo.Caption) - 1 & ""
        CrystalReport1.Formulas(0) = "xxx1='" & inwordsusd(Val(txtTtlAmount)) & " Only" & "'"
        CrystalReport1.Formulas(1) = "curtype='" & cboCurrency.Text & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1

    End If
    
  End If
End Sub
Private Sub Form_Load()
 txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
 cboCurrency.AddItem "DHS"
 cboCurrency.AddItem "USD"
 lblConvRate.Visible = False
 txtConvRate.Visible = False
 
 LSTSUP = 0
 LSTAC = 1
 LSTAGN = 0
 
 AutoIncrementVoucher
 
 
 PopulateAcctSuppCust
 
 Flexitems
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
  Sqlqry = "delete * from dumCrpt1"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
 
End Sub
Private Sub AutoIncrementVoucher()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select DOC_no from DOCU_MAS WHERE DOC_TYPE='CRT'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
       MsgBox "Document type 'CRT' not found"
       Exit Sub
    Else
       lblVoucNo = Val(rs!doc_no) + 1
    End If
End Sub
Private Sub PopulateAcctSuppCust()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Acct_mas order by acct_code"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    
    lstAcctCode.Clear
    
    If rs.RecordCount = 0 Then
        MsgBox "No Records found in the Accounts Master"
    Else
       rs.MoveFirst
       Do Until rs.EOF
          lstAcctCode.AddItem rs!acct_code & "  :  " & rs!acct_name
          rs.MoveNext
       Loop
    End If
    
End Sub

Private Function ValidateData()
ValidateData = False
If txtdate.TextWithMask = "" Or IsDate(txtdate.TextWithMask) = False Then
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
Private Sub txtAmount_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtAmount_LostFocus()

Dim accd As String
Dim acname As String


  If ValidateData = True Then
     If Val(txtAmount.Text) > Val(txtTtlAmount.Text) Then
      MsgBox " Entered Amount Greater than Total Amount"
      txtAmount.SetFocus
      Exit Sub
     End If
      
  If LSTAC = 1 Then
     accd = Mid(lstAcctCode, 1, 6)
     acname = Mid(lstAcctCode, 12, 35)
  ElseIf LSTSUP = 1 Then
    accd = Mid(lstAcctCode, 1, 4)
    acname = Mid(lstAcctCode, 10, 35)
  Else
    accd = "AGNC"
    acname = Trim(lstAcctCode.Text)
  End If
  
 
    cur = ""
    cod = ""
    con = 1
 
  If cboCurrency.Text = "USD" Then
      cur = "USD"
      cod = "103002"
      con = Val(Trim(txtConvRate.Text))
       
  Else
      cur = "DHS"
      cod = "103001"
      con = 1
  End If
 
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = " select * from dumCrpt1"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
    If txtAmount.Text = 0 Then
      Exit Sub
      txtAmount.SetFocus
    End If
    If rs.RecordCount = 0 Then
       Sqlqry = " Insert into dumCrpt1 values('" & lblVoucNo & "','CRT','" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & accd & "','" _
                                     & findfirstfixup(acname) & "','" _
                                     & findfirstfixup(UCase(Trim(txtdesc))) & "','" _
                                     & Trim(cboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(txtAmount.Text) & "," _
                                     & Val(txtAmount) * Val(txtConvRate) & ")"

        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        
        Sqlqry1 = "select * from dumCrpt1"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount = 0 Then
           ' MSFlexGrid1.Clear
            Exit Sub
        Else
            Flexitems
            rs.MoveFirst
          Do Until rs.EOF
           MSFlexGrid1.AddItem rs!acct_code & Chr(9) & rs!acct_name & Chr(9) & rs!Description & Chr(9) & rs!tra_amount
           rs.MoveNext
          Loop
        End If
        lblTtlAmount = Round(Val(txtAmount.Text), 2)
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
         Sqlqry = " Insert into dumCrpt1 values('" & lblVoucNo & "','CRT','" _
                                     & Format(txtdate.TextWithMask, "dd/mm/yyyy") & "','" _
                                     & accd & "','" _
                                     & findfirstfixup(acname) & "','" _
                                     & findfirstfixup(UCase(Trim(txtdesc))) & "','" _
                                     & Trim(cboCurrency.Text) & "'," _
                                     & Val(txtConvRate.Text) & "," _
                                     & Val(txtAmount.Text) & "," _
                                     & Val(txtAmount) * Val(txtConvRate) & ")"

        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
          Sqlqry1 = "Select * from dumCrpt1"
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
             lblTtlAmount = Round(y, 2)
             lblTtlAmount.Alignment = 1
             If y = Val(txtTtlAmount.Text) Then
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

Private Sub txtConvRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPaidTo.SetFocus
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboCurrency.SetFocus
End Sub

Private Sub txtdate_LostFocus()
If IsDate(txtdate.TextWithMask) = False Then
   MsgBox "Invalid Date", vbInformation, "Invalid Entry"
   txtdate.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAmount.SetFocus
End Sub
Private Sub txtpaidto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtTtlDesc.SetFocus
End Sub
Private Sub txtTtlAmount_KeyPress(KeyAscii As Integer)
If cboCurrency.Text = "USD" Then
 If KeyAscii = 13 Then txtConvRate.SetFocus
Else
 If KeyAscii = 13 Then txtPaidTo.SetFocus
End If
End Sub
Private Sub txtTtlAmount_LostFocus()
txtAmount.Text = Val(txtTtlAmount.Text)
End Sub
Private Sub txtTtlDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstAcctCode.SetFocus
End Sub

Private Function textclear()
     txtPaidTo.Text = ""
     txtTtlDesc.Text = ""
     txtdesc.Text = ""
     txtAmount.Text = ""
     txtConvRate.Text = ""
     lblTtlAmount.Caption = ""
     txtdate.TextWithMask = Format(Now, "dd/mm/yyyy")
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Delete * from dumCrpt1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     Flexitems
     txtTtlAmount.SetFocus
End Function

Private Sub txtTtlDesc_LostFocus()
txtdesc.Text = Trim(txtTtlDesc.Text)
End Sub
