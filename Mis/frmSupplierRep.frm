VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "PVMASK.OCX"
Begin VB.Form frmSupplierRep 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Statement of Supplier  "
   ClientHeight    =   8775
   ClientLeft      =   15
   ClientTop       =   270
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Statement of Supplier"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   600
      Width           =   7215
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   6360
         Top             =   5400
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
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
         Picture         =   "frmSupplierRep.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5160
         Width           =   1695
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
         Height          =   975
         Left            =   2640
         Picture         =   "frmSupplierRep.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Print Preview"
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
         Left            =   1080
         Picture         =   "frmSupplierRep.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5160
         Width           =   1575
      End
      Begin VB.ListBox lstSuppliers 
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
         ForeColor       =   &H00404040&
         Height          =   3060
         Left            =   480
         TabIndex        =   0
         Top             =   480
         Width           =   6375
      End
      Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   3720
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
         Left            =   3240
         TabIndex        =   8
         Top             =   4320
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
         X1              =   7200
         X2              =   0
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label3 
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
         Height          =   240
         Left            =   1920
         TabIndex        =   6
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Height          =   240
         Left            =   1920
         TabIndex        =   5
         Top             =   3840
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmSupplierRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim Opbal As Currency
Dim opcrpt As Currency
Dim opcpmt As Currency
Dim Opbrpt As Currency
Dim Opbpmt As Currency
Dim Opprpt As Currency
Dim Opppmt As Currency
Dim opjbd As Currency
Dim Opjbc As Currency
Dim opdbntdb As Currency
Dim opdbntcr As Currency
Dim Opcrntdb As Currency
Dim Opcrntcr As Currency
Dim Opcrsl As Currency
Dim Opcrpr As Currency
Dim Ttlopbal As Currency
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim Sqlqry3 As String
Dim Sqlqry4 As String
Dim Sqlqry5 As String
Dim Sqlqry6 As String
Dim Sqlqry7 As String
Dim Sqlqry8 As String
Dim Sqlqry9 As String
Dim Sqlqry10 As String
Dim SQLQRY11 As String
Dim SQLQRY12 As String
Dim Sqlqry13 As String
Dim Sqlqry14 As String
Dim Sqlqry15 As String
Dim sqlqry16 As String
Dim sqlqry17 As String
Dim sqlqry18 As String
Dim sqlqry19 As String
Dim sqlqry20 As String
Dim Sqlqry21 As String
Dim Sqlqry22 As String
Dim Sqlqry23 As String
Dim Sqlqry24 As String
Dim Sqlqry25 As String
Dim Sqlqry26 As String
Dim Sqlqry27 As String
Dim Sqlqry28 As String
Dim Sqlqry29 As String
Dim Sqlqry30 As String
Dim Sqlqry31 As String
Dim Sqlqry32 As String
Dim Sqlqry33 As String
Dim Sqlqry34 As String
Dim Sqlqry35 As String
Dim Sqlqry36 As String
Dim Sqlqry37 As String
Dim Sqlqry38 As String
Dim Sqlqry39 As String
Dim Sqlqry40 As String
Dim Sqlqry41 As String
Dim Sqlqry42 As String
Dim Sqlqry43 As String
Dim Sqlqry44 As String
Dim Sqlqry45 As String
Dim Sqlqry46 As String
Dim Sqlqry47 As String
Dim Sqlqry48 As String
Dim Sqlqry49 As String
Dim Sqlqry50 As String
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim rs3 As Recordset
Dim rs4 As Recordset
Dim rs5 As Recordset
Dim rs6 As Recordset
Dim rs7 As Recordset
Dim rs8 As Recordset
Dim rs9 As Recordset
Dim rs10 As Recordset
Dim rs11 As Recordset
Dim rs12 As Recordset
Dim rs13 As Recordset
Dim rs14 As Recordset
Dim rs15 As Recordset
Dim rs16 As Recordset
Dim rs17 As Recordset
Dim rs18 As Recordset
Dim rs19 As Recordset
Dim rs20 As Recordset
Dim rs21 As Recordset
Dim rs22 As Recordset
Dim rs23 As Recordset
Dim rs24 As Recordset
Dim rs25 As Recordset
Dim rs26 As Recordset
Dim rs27 As Recordset
Dim rs28 As Recordset
Dim rs29 As Recordset
Dim rs30 As Recordset
Dim rs31 As Recordset
Dim rs32 As Recordset
Dim rs33 As Recordset
Dim rs34 As Recordset
Dim rs35 As Recordset
Dim vaddr
Dim vcity
Dim vcountry
Dim vtel
Dim vfax
Dim yy
Dim ttlpdc As Currency
Private Sub CmdBack_Click()
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
        
        Sqlqry = " Delete * from SuppReport"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        
        Opbal = 0
        opcrpt = 0
        opcpmt = 0
        Opbrpt = 0
        Opbpmt = 0
        Opprpt = 0
        Opppmt = 0
        opjbd = 0
        Opjbc = 0
        opdbntdb = 0
        opdbntcr = 0
        Opcrntdb = 0
        Opcrntcr = 0
        Opcrsl = 0
        Opcrpr = 0
        Ttlopbal = 0
        ttlpdc = 0
             
    ' Opening Balance from Supplier Finance
        Sqlqry = " select * from Supp_Fin where Supp_no='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
          MsgBox " Supplier Code not found in Supp_Fin"
          Exit Sub
        Else
          rs.MoveFirst
          Opbal = rs!open_bal
          vaddr = rs!Address
          vcity = rs!city
          vcountry = rs!country
          vtel = rs!telephone
          vfax = rs!fax
          yy = Trim(lstSuppliers.Text)
        End If
        
    ' Cash Receipt before From date
        Sqlqry1 = "select sum(amount) from crpt_tra where tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then opcrpt = rs1.Fields(0)
        
    ' Cash Payment before From date
        Sqlqry2 = "select sum(Amount) from cpmt_Tra where tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs2.Fields(0)) = False Then opcpmt = rs2.Fields(0)
                     
    ' Bank Receipt before From date
        Sqlqry3 = "select sum(amount) from brpt_tra where tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs3 = db.OpenRecordset(Sqlqry3, dbOpenDynaset)
        If IsNull(rs3.Fields(0)) = False Then Opbrpt = rs3.Fields(0)
                
    ' Bank Payment before From date
        Sqlqry4 = "select sum(Amount) from bpmt_tra where tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs4 = db.OpenRecordset(Sqlqry4, dbOpenDynaset)
        If IsNull(rs4.Fields(0)) = False Then Opbpmt = rs4.Fields(0)
        
    ' Pdc Receipts before From date
        Sqlqry6 = "Select sum(amount) from Prpt_mas1 where Cheque_Dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_dt) and acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs6 = db.OpenRecordset(Sqlqry6, dbOpenDynaset)
        If IsNull(rs6.Fields(0)) = False Then Opprpt = rs6.Fields(0)
        
     ' Pdc Payments before From Date
        Sqlqry8 = "Select sum(amount) from Ppmt_tra where Cheque_Dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(posting_dt) and acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs8 = db.OpenRecordset(Sqlqry8, dbOpenDynaset)
        If IsNull(rs8.Fields(0)) = False Then Opppmt = rs8.Fields(0)
         
     ' Journal Debit Amount before From Date
        Sqlqry9 = "select sum(damount) from Jrnl_tra where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstSuppliers, 1, 4) & "' and dc_code ='D'"
        Set rs9 = db.OpenRecordset(Sqlqry9, dbOpenDynaset)
        If IsNull(rs9.Fields(0)) = False Then opjbd = rs9.Fields(0)
                  
     ' Journal Credit Amount before From Date
        Sqlqry10 = "select sum(camount) from Jrnl_tra where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstSuppliers, 1, 4) & "' and dc_code ='C'"
        Set rs10 = db.OpenRecordset(Sqlqry10, dbOpenDynaset)
        If IsNull(rs10.Fields(0)) = False Then Opjbc = rs10.Fields(0)
        
     ' Debit note (credit) before From Date
        SQLQRY11 = "select sum(amount) from debt_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs11 = db.OpenRecordset(SQLQRY11, dbOpenDynaset)
        If IsNull(rs11.Fields(0)) = False Then opdbntcr = rs11.Fields(0)
        
      ' Debit note (debit)  before From Date
        SQLQRY12 = "select sum(amount) from debt_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and cust_no='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs12 = db.OpenRecordset(SQLQRY12, dbOpenDynaset)
        If IsNull(rs12.Fields(0)) = False Then opdbntdb = rs12.Fields(0)
        
      ' Credit note (debit) before From Date
        Sqlqry13 = "select sum(amount) from crdt_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs13 = db.OpenRecordset(Sqlqry13, dbOpenDynaset)
        If IsNull(rs13.Fields(0)) = False Then Opcrntdb = rs13.Fields(0)
                 
       ' Credit note (credit)  before From Date
        Sqlqry14 = "select sum(amount) from crdt_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and supp_no='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs14 = db.OpenRecordset(Sqlqry14, dbOpenDynaset)
        If IsNull(rs14.Fields(0)) = False Then Opcrntcr = rs14.Fields(0)
        
                  
                                     
      ' Opening balance credit Purchases  before From Date
        sqlqry16 = "select sum(gamount) from crpr_mas where tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and supp_no='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs16 = db.OpenRecordset(sqlqry16, dbOpenDynaset)
        If IsNull(rs16.Fields(0)) = False Then Opcrpr = rs16.Fields(0)
        
        
        Ttlopbal = Opbal - opcrpt + opcpmt - Opbrpt + Opbpmt - Opprpt + Opppmt + opjbd - Opjbc _
                    + opdbntdb - opdbntcr _
                    + Opcrntdb - Opcrntcr + Opcrsl - Opcrpr
                    
        
        If Ttlopbal < 0 Then
           sqlqry17 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "'," & 0 & ",'','" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','Opening Balance'," & 0 & "," & Abs(Ttlopbal) & ")"
           ws.BeginTrans
           db.Execute (sqlqry17)
           ws.CommitTrans
        Else
           sqlqry17 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "'," & 0 & ",'','" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','Opening Balance'," & Ttlopbal & "," & 0 & ")"
           ws.BeginTrans
           db.Execute (sqlqry17)
           ws.CommitTrans
        End If
          
        ' Cash Receipt after From date and before to date
        sqlqry18 = "select * from crpt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs18 = db.OpenRecordset(sqlqry18, dbOpenDynaset)
        If rs18.RecordCount <> 0 Then
          rs18.MoveFirst
         Do Until rs18.EOF
          sqlqry19 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs18!VOUC_NO & "','" & rs18!vouc_type & "','" & Trim(rs18!tdate) & "','" & Trim(rs18!Description) & "'," & 0 & "," & Trim(rs18!amount) & ")"
          ws.BeginTrans
          db.Execute (sqlqry19)
          ws.CommitTrans
          rs18.MoveNext
         Loop
        End If
        
        ' Cash Payment after From date and before to date
        sqlqry20 = "select * from Cpmt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs19 = db.OpenRecordset(sqlqry20, dbOpenDynaset)
        If rs19.RecordCount <> 0 Then
         rs19.MoveFirst
         Do Until rs19.EOF
          Sqlqry21 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs19!VOUC_NO & "','" & rs19!vouc_type & "','" & Trim(rs19!tdate) & "','" & Trim(rs19!Description) & "'," & Trim(rs19!amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry21)
          ws.CommitTrans
          rs19.MoveNext
         Loop
        End If
                
        ' Bank Receipt after From date and before to date
        Sqlqry22 = "select * from brpt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs20 = db.OpenRecordset(Sqlqry22, dbOpenDynaset)
        If rs20.RecordCount <> 0 Then
         rs20.MoveFirst
         Do Until rs20.EOF
          Sqlqry23 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs20!VOUC_NO & "','" & rs20!vouc_type & "','" & Trim(rs20!tdate) & "','" & Trim(rs20!Description) & "'," & 0 & "," & Trim(rs20!amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry23)
          ws.CommitTrans
          rs20.MoveNext
         Loop
        End If
        
        ' Bank Payment after From date and before to date
        Sqlqry24 = "select * from bpmt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs21 = db.OpenRecordset(Sqlqry24, dbOpenDynaset)
        If rs21.RecordCount <> 0 Then
         rs21.MoveFirst
         Do Until rs21.EOF
          Sqlqry25 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs21!VOUC_NO & "','" & rs21!vouc_type & "','" & Trim(rs21!tdate) & "','" & Trim(rs21!Description) & "'," & Trim(rs21!amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry25)
          ws.CommitTrans
          rs21.MoveNext
         Loop
        End If
        
       ' Pdc Receipts after From date and before to date
        Sqlqry26 = "select * from prpt_mas1 where Cheque_Dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and not isnull(Posting_Dt) and acct_code ='" & Mid(lstSuppliers, 1, 4) & "' "
        Set rs22 = db.OpenRecordset(Sqlqry26, dbOpenDynaset)
        If rs22.RecordCount <> 0 Then
         rs22.MoveFirst
          Do Until rs22.EOF
               Sqlqry28 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs23!VOUC_NO & "','" & rs22!vouc_type & "','" & Trim(rs22!Cheque_Dt) & "','" & Trim(rs22!Description) & "'," & 0 & "," & Trim(rs22!amount) & ")"
               ws.BeginTrans
               db.Execute (Sqlqry28)
               ws.CommitTrans
             
           rs22.MoveNext
          Loop
        End If
        
        ' Pdc Payments after From date and before to date
        Sqlqry29 = "select * from Ppmt_mas where Cheque_Dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#  and not Isnull(Posting_Dt) "
        Set rs24 = db.OpenRecordset(Sqlqry29, dbOpenDynaset)
        If rs24.RecordCount <> 0 Then
         rs24.MoveFirst
         Do Until rs24.EOF
            Sqlqry30 = "Select * from Ppmt_tra where Vouc_no=" & Val(rs24!VOUC_NO) & " and acct_code ='" & Mid(lstSuppliers, 1, 4) & "'"
            Set rs25 = db.OpenRecordset(Sqlqry30, dbOpenDynaset)
             If rs25.RecordCount <> 0 Then
               rs25.MoveFirst
                Do Until rs25.EOF
                  Sqlqry31 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs25!VOUC_NO & "','" & rs25!vouc_type & "','" & Trim(rs24!Cheque_Dt) & "','" & Trim(rs25!Description) & "'," & Trim(rs25!amount) & "," & 0 & ")"
                    ws.BeginTrans
                    db.Execute (Sqlqry31)
                    ws.CommitTrans
                 rs25.MoveNext
                Loop
              End If
          rs24.MoveNext
          Loop
         End If
         
       ' Journal Debit after From date and before to date
        Sqlqry32 = "select * from jrnl_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstSuppliers, 1, 4) & "' and dc_code='D' "
        Set rs26 = db.OpenRecordset(Sqlqry32, dbOpenDynaset)
         If rs26.RecordCount <> 0 Then
          rs26.MoveFirst
          Do Until rs26.EOF
          Sqlqry33 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs26!VOUC_NO & "','" & rs26!vouc_type & "','" & Trim(rs26!tdate) & "','" & Trim(rs26!Description) & "'," & Trim(rs26!damount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry33)
          ws.CommitTrans
          rs26.MoveNext
         Loop
        End If
        
        ' Journal Credit after From date and before to date
        Sqlqry33 = "select * from Jrnl_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstSuppliers, 1, 4) & "' and dc_code='C'"
        Set rs27 = db.OpenRecordset(Sqlqry33, dbOpenDynaset)
        If rs27.RecordCount <> 0 Then
         rs27.MoveFirst
        Do Until rs27.EOF
          Sqlqry34 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs27!VOUC_NO & "','" & rs27!vouc_type & "','" & Trim(rs27!tdate) & "','" & Trim(rs27!Description) & "'," & 0 & "," & Trim(rs27!camount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry34)
          ws.CommitTrans
          rs27.MoveNext
         Loop
        End If
        
     ' DebitNote - credit after From date and before to date
        Sqlqry35 = "select * from debt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs28 = db.OpenRecordset(Sqlqry35, dbOpenDynaset)
        If rs28.RecordCount <> 0 Then
          rs28.MoveFirst
         Do Until rs28.EOF
          Sqlqry36 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs28!VOUC_NO & "','" & rs28!vouc_type & "','" & Trim(rs28!tdate) & "','" & Trim(rs28!Description) & "'," & 0 & "," & Trim(rs28!amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry36)
          ws.CommitTrans
          rs28.MoveNext
         Loop
        End If
        
        ' DebitNote - debit after From date and before to date
        Sqlqry37 = "select * from debt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Cust_no='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs29 = db.OpenRecordset(Sqlqry37, dbOpenDynaset)
        If rs29.RecordCount <> 0 Then
          rs29.MoveFirst
         Do Until rs29.EOF
          Sqlqry38 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs29!VOUC_NO & "','" & rs29!vouc_type & "','" & Trim(rs29!tdate) & "','" & Trim(rs29!Description) & "'," & Trim(rs29!amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry38)
          ws.CommitTrans
          rs29.MoveNext
         Loop
        End If
        
        ' CreditNote - Credit after From date and before to date
        Sqlqry38 = "select * from crdt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Supp_no='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs30 = db.OpenRecordset(Sqlqry38, dbOpenDynaset)
         If rs30.RecordCount <> 0 Then
           rs30.MoveFirst
         Do Until rs30.EOF
           Sqlqry39 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs30!VOUC_NO & "','" & rs30!vouc_type & "','" & Trim(rs30!tdate) & "','" & Trim(rs30!Description) & "'," & 0 & "," & Trim(rs30!amount) & ")"
           ws.BeginTrans
           db.Execute (Sqlqry39)
           ws.CommitTrans
           rs30.MoveNext
         Loop
        End If
        
        ' Credit Note - Debit after From date and before to date
        Sqlqry40 = "select * from crdt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs31 = db.OpenRecordset(Sqlqry40, dbOpenDynaset)
        If rs31.RecordCount <> 0 Then
         rs31.MoveFirst
        Do Until rs31.EOF
          Sqlqry41 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs31!VOUC_NO & "','" & rs31!vouc_type & "','" & Trim(rs31!tdate) & "','" & Trim(rs31!Description) & "'," & Trim(rs31!amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry41)
          ws.CommitTrans
          rs31.MoveNext
         Loop
        End If
        
      
        
        ' Credit Purchase after From date and before to date
        Sqlqry44 = "select * from crpr_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Supp_no='" & Mid(lstSuppliers, 1, 4) & "'"
        Set rs33 = db.OpenRecordset(Sqlqry44, dbOpenDynaset)
        If rs33.RecordCount <> 0 Then
         rs33.MoveFirst
        Do Until rs33.EOF
          Sqlqry45 = "Insert into SuppReport values('" & Mid(lstSuppliers, 1, 4) & "','" & findfirstfixup(Mid(lstSuppliers, 8, 30)) & "','" & rs33!VOUC_NO & "','" & rs33!vouc_type & "','" & Trim(rs33!tdate) & "','Purchase'," & 0 & "," & Trim(rs33!gamount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry45)
          ws.CommitTrans
          rs33.MoveNext
         Loop
        End If
       
       'Total Amount of Pdc Issued With out Posting entries
        Sqlqry46 = "select * from Ppmt_mas where Isnull(Posting_Dt) "
        Set rs34 = db.OpenRecordset(Sqlqry46, dbOpenDynaset)
        If rs34.RecordCount <> 0 Then
         rs34.MoveFirst
         Do Until rs34.EOF
            Sqlqry47 = "Select * from Ppmt_tra where Vouc_no=" & Val(rs34!VOUC_NO) & " and acct_code ='" & Mid(lstSuppliers, 1, 4) & "'"
            Set rs35 = db.OpenRecordset(Sqlqry47, dbOpenDynaset)
             If rs35.RecordCount <> 0 Then
               rs35.MoveFirst
                Do Until rs35.EOF
                  ttlpdc = ttlpdc + rs35!amount
                  rs35.MoveNext
                Loop
              End If
          rs34.MoveNext
          Loop
         End If
     
        
    With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\supreport.rpt"
     .Formulas(0) = "zzz='" & " As on " & Trim(txtdateto.TextWithMask) & "'"
     .Formulas(1) = "Address='" & Trim(vaddr) & "'"
     .Formulas(2) = "City='" & "City : " & Trim(vcity) & "   Country : " & Trim(vcountry) & "'"
     .Formulas(3) = "TelFax='" & " Tel : " & Mid(vtel, 1, 15) & "    Fax : " & Mid(vfax, 1, 15) & "'"
     .Formulas(4) = "Pdc =" & ttlpdc & ""
     .WindowState = crptMaximized
     .Action = 1
    End With
   
   Else
        MsgBox "Improper Dates", vbDefaultButton1, "Invalid entry"
        Exit Sub
  End If
End Sub
Private Sub Form_Load()
 PopulateSuppliers
 txtdatefrom.TextWithMask = Format(Now, "DD/mm/yyyy")
 txtdateto.TextWithMask = Format(Now, "DD/mm/yyyy")
End Sub
Private Sub PopulateSuppliers()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select * from Supp_fin order by Supp_name"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

lstSuppliers.Clear

 If rs.RecordCount = 0 Then
      MsgBox "No Records found in the Supplier Master"
 Else
      rs.MoveFirst
    Do Until rs.EOF
      lstSuppliers.AddItem rs!Supp_no & " : " & rs!Supp_name
      rs.MoveNext
   Loop
 End If

End Sub

Private Sub lstSuppliers_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdatefrom.SetFocus
End Sub
Private Sub txtdatefrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtdateto.SetFocus
End Sub

Private Sub txtdatefrom_LostFocus()
If IsDate(txtdatefrom.TextWithMask) = False Then
      MsgBox "Invalid Date from", vbInformation, "Invalid Entry"
      txtdatefrom.SetFocus
      SendKeys "{Home} + {End}"
    End If
End Sub

Private Sub txtdateto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdDisplay.SetFocus
End Sub
Private Function ValidateData()
ValidateData = False

If IsDate(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf lstSuppliers.SelCount = 0 Then
  MsgBox "Select Supplier", vbInformation, "Invalid Entry"
  lstSuppliers.SetFocus
  SendKeys " {Home} + {end} "
  Exit Function
ElseIf IsDate(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid To Date", vbInformation, "Invalid Entry"
   txtdateto.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
ValidateData = True
End If
End Function
Private Sub textclear()
 lstSuppliers.ListIndex = -1
 txtdatefrom.TextWithMask = Format(Now, "DD/mm/yyyy")
 txtdateto.TextWithMask = Format(Now, "DD/mm/YYYY")
End Sub

Private Sub txtdateto_LostFocus()
If IsDate(txtdateto.TextWithMask) = False Then
      MsgBox "Invalid Date to", vbInformation, "Invalid Entry"
      txtdateto.SetFocus
      SendKeys "{Home} + {End}"
    End If
End Sub
