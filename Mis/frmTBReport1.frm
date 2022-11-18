VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmTBReport1 
   Caption         =   "Trial Balance1"
   ClientHeight    =   5805
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   7695
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   4815
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   1920
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<<&Back"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "&Display"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtDateFrom 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trial Balance As on"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   2400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enter Date "
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmTBReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ws As Workspace
Dim db As Database
Dim Opbal As Currency
Dim Opcrptcr As Currency
Dim Opcrptdb As Currency
Dim Opcpmtcr As Currency
Dim Opcpmtdb As Currency
Dim Opbrptcr As Currency
Dim Opbrptdb As Currency
Dim Opbpmtdb As Currency
Dim Opbpmtcr As Currency
Dim Opprptcr As Currency
Dim Opprptdb As Currency
Dim Opppmtcr As Currency
Dim Opppmtdb As Currency
Dim Opjrnldb As Currency
Dim Opjrnlcr As Currency
Dim opdbntcr As Currency
Dim opdbntdb As Currency
Dim Opcrntcr As Currency
Dim Opcrntdb As Currency
Dim Opcrslcr As Currency
Dim Opcrsldb As Currency
Dim Opcrprdb As Currency
Dim OpcaslCr As Currency
Dim Opcasldb As Currency
Dim Opcaprcr As Currency
Dim Opcaprdb As Currency
Dim opcpmt As Currency
Dim opcrpt As Currency
Dim Opcasl As Currency
Dim Opbrpta As Currency
Dim Opbpmta As Currency
Dim Ttlopbal As Currency
Dim TTLCUST As Currency
Dim Advclbal As Currency
Dim VehClbal As Currency
Dim ProClbal As Currency
Dim PurClbal As Currency
Dim SalClbal As Currency
Dim GoodClbal As Currency
Dim X As Currency
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
Dim Sqlqry51 As String
Dim Sqlqry52 As String
Dim Sqlqry53 As String
Dim Sqlqry54 As String
Dim Sqlqry55 As String
Dim Sqlqry56 As String
Dim Sqlqry57 As String
Dim Sqlqry58 As String
Dim Sqlqry59 As String
Dim Sqlqry60 As String
Dim Sqlqry61 As String
Dim Sqlqry62 As String
Dim Sqlqry63 As String
Dim Sqlqry64 As String
Dim Sqlqry65 As String
Dim Sqlqry66 As String
Dim Sqlqry67 As String
Dim Sqlqry68 As String
Dim Sqlqry69 As String
Dim Sqlqry70 As String
Dim Sqlqry71 As String
Dim Sqlqry72 As String
Dim Sqlqry73 As String
Dim Sqlqry74 As String
Dim Sqlqry75 As String
Dim Sqlqry76 As String
Dim Sqlqry77 As String
Dim Sqlqry78 As String
Dim Sqlqry79 As String
Dim Sqlqry80 As String
Dim Sqlqry81 As String
Dim Sqlqry82 As String
Dim Sqlqry83 As String
Dim Sqlqry84 As String
Dim Sqlqry85 As String
Dim Sqlqry86 As String
Dim Sqlqry87 As String
Dim Sqlqry88 As String
Dim Sqlqry89 As String
Dim Sqlqry90 As String
Dim Sqlqry91 As String
Dim Sqlqry92 As String
Dim Sqlqry93 As String
Dim Sqlqry94 As String
Dim Sqlqry95 As String
Dim Sqlqry96 As String
Dim Sqlqry97 As String
Dim Sqlqry98 As String
Dim Sqlqry99 As String
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
Dim rs36 As Recordset
Dim rs37 As Recordset
Dim rs38 As Recordset
Dim rs39 As Recordset
Dim rs40 As Recordset
Dim rs41 As Recordset
Dim rs42 As Recordset
Dim rs43 As Recordset
Dim rs44 As Recordset
Dim rs45 As Recordset
Dim rs46 As Recordset
Dim rs47 As Recordset
Dim rs48 As Recordset
Dim rs49 As Recordset
Dim rs50 As Recordset
Dim rs51 As Recordset
Dim rs52 As Recordset
Dim rs53 As Recordset
Dim rs54 As Recordset
Dim rs55 As Recordset
Dim rs56 As Recordset
Dim rs57 As Recordset
Dim rs58 As Recordset
Dim rs59 As Recordset
Dim rs60 As Recordset
Dim rs61 As Recordset
Dim rs62 As Recordset
Dim rs63 As Recordset
Dim rs64 As Recordset
Dim rs65 As Recordset
Dim rs66 As Recordset
Dim rs67 As Recordset
Dim rs68 As Recordset
Dim rs69 As Recordset
Dim rs70 As Recordset
Dim rs71 As Recordset
Dim rs72 As Recordset
Dim rs73 As Recordset
Dim rs74 As Recordset
Dim rs75 As Recordset
Dim rs76 As Recordset
Dim rs77 As Recordset
Dim rs78 As Recordset
Dim rs79 As Recordset
Dim rs80 As Recordset
Dim rs81 As Recordset
Dim rs82 As Recordset
Dim opsal As Currency
Dim oprec As Currency
Dim oppur As Currency
Dim oppay As Currency
Dim Opbrpt As Currency
Dim Opprpt As Currency
Dim Opjdb As Currency
Dim Opbpmt As Currency
Dim Opppmt As Currency
Dim Opjcr As Currency
Dim opjbd As Currency
Dim Opjbc As Currency
Dim Opcrsl As Currency
Dim Opcrpr As Currency
Dim Ttlsupp As Currency
Dim Tpdc As Currency
Private Sub cmdBack_Click()
 Unload Me
End Sub
Private Sub cmdClear_Click()
 textclear
End Sub
 Private Sub cmdDisplay_Click()
  If ValidateData = True Then
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase("c:\uday\udfin.mdb")
    doclosnil
    Sqlqry = " select * from acct_mas order by acct_code"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount = 0 Then
       MsgBox " Selected Code not found in Account Master"
       Exit Sub
     Else
      rs.MoveFirst
      Do Until rs.EOF
        Opbal = 0
        Opcrptcr = 0
        Opcrptdb = 0
        Opcpmtcr = 0
        Opcpmtdb = 0
        Opbrptcr = 0
        Opbrptdb = 0
        Opbpmtdb = 0
        Opbpmtcr = 0
        Opprptcr = 0
        Opprptdb = 0
        Opppmtcr = 0
        Opppmtdb = 0
        Opjrnldb = 0
        Opjrnlcr = 0
        opdbntcr = 0
        opdbntdb = 0
        Opcrntcr = 0
        Opcrntdb = 0
        Opcrslcr = 0
        Opcrsldb = 0
        Opcrprdb = 0
        OpcaslCr = 0
        Opcasldb = 0
        Opcaprcr = 0
        Opcaprdb = 0
        Ttlopbal = 0
        TTLCUST = 0
        Ttlsupp = 0
        Opcasl = 0
        Tpdc = 0
        opsal = 0
        oprec = 0
        oppur = 0
        oppay = 0
        Opbrpt = 0
        Opbpmt = 0
        opcpmt = 0
        opcrpt = 0
        Opprpt = 0
        Opppmt = 0
        Opjdb = 0
        Opjcr = 0
        Opcrsl = 0
        Opcrpr = 0
        If IsNull(rs!open_bal) = True Then
         Opbal = 0
        Else
        Opbal = rs!open_bal
        End If
     If Val(rs!acct_code) = 103001 Then
         docashbal
     ElseIf Val(rs!acct_code) >= 103101 And Val(rs!acct_code) < 103200 Then
         dobankbal
     Else
         
        ' Opening Balance from Cash Receipt(credit) Transactions
        Sqlqry1 = " select * from crpt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs1.RecordCount <> 0 Then
          rs1.MoveFirst
          Do Until rs1.EOF
           Opcrptcr = Opcrptcr + rs1!amount
           rs1.MoveNext
          Loop
        End If
        
        ' Opening Balance from Cash Receipt(debit) master
        Sqlqry2 = " select * from crpt_mas where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Cash_code='" & rs!acct_code & "'"
        Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If rs2.RecordCount <> 0 Then
          rs2.MoveFirst
          Do Until rs2.EOF
           Opcrptdb = Opcrptdb + rs2!ttl_amount
           rs2.MoveNext
          Loop
        End If
        
        ' Opening Balance from Cash Payment(debit) Transactions
        Sqlqry3 = " select * from cpmt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs3 = db.OpenRecordset(Sqlqry3, dbOpenDynaset)
        If rs3.RecordCount <> 0 Then
          rs3.MoveFirst
          Do Until rs3.EOF
          Opcpmtdb = Opcpmtdb + rs3!amount
          rs3.MoveNext
          Loop
        End If
        
        ' Opening Balance from Cash Payment(credit) Transactions
        Sqlqry4 = " select * from cpmt_mas where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and cash_code='" & rs!acct_code & "'"
        Set rs4 = db.OpenRecordset(Sqlqry4, dbOpenDynaset)
        If rs4.RecordCount <> 0 Then
          rs4.MoveFirst
          Do Until rs4.EOF
           Opcpmtcr = Opcpmtcr + rs4!ttl_amount
           rs4.MoveNext
          Loop
        End If
        
        ' Opening Balance from Bank Receipt(credit) transaction before from date.
        Sqlqry5 = "select * from brpt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "'"
        Set rs5 = db.OpenRecordset(Sqlqry5, dbOpenDynaset)
        If rs5.RecordCount <> 0 Then
         rs5.MoveFirst
         Do Until rs5.EOF
          Opbrptcr = Opbrptcr + rs5!amount
          rs5.MoveNext
         Loop
        End If
        
        ' Opening Balance from Bank Receipt(debit) transaction before from date.
        Sqlqry6 = "select * from brpt_mas where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and bank_code='" & rs!acct_code & "'"
        Set rs6 = db.OpenRecordset(Sqlqry6, dbOpenDynaset)
        If rs6.RecordCount <> 0 Then
          rs6.MoveFirst
         Do Until rs6.EOF
          Opbrptdb = Opbrptdb + rs6!ttl_amount
          rs6.MoveNext
         Loop
        End If
        
        ' Opening Balance from Bank Payment(debit) before From date
        Sqlqry7 = "select * from bpmt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs7 = db.OpenRecordset(Sqlqry7, dbOpenDynaset)
        If rs7.RecordCount <> 0 Then
         rs7.MoveFirst
         Do Until rs7.EOF
          Opbpmtdb = Opbpmtdb + rs7!amount
          rs7.MoveNext
         Loop
        End If
                
        ' Opening Balance from Bank Payment(credit) before From date
        Sqlqry8 = "select * from bpmt_mas where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and bank_code='" & rs!acct_code & "'"
        Set rs8 = db.OpenRecordset(Sqlqry8, dbOpenDynaset)
        If rs8.RecordCount <> 0 Then
         rs8.MoveFirst
         Do Until rs8.EOF
          Opbpmtcr = Opbpmtcr + rs8!ttl_amount
          rs8.MoveNext
         Loop
        End If
        
       ' Opening Pdc Receipts before From date
        Sqlqry9 = "select * from prpt_mas where Cheque_Dt<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and not isnull(posting_dt)"
        Set rs9 = db.OpenRecordset(Sqlqry9, dbOpenDynaset)
        If rs9.RecordCount <> 0 Then
        rs9.MoveFirst
        Do Until rs9.EOF
          If rs9!bank_code = rs!acct_code Then
            Opprptdb = Opprptdb + rs9!ttl_amount
          End If
          
          Sqlqry10 = "Select * from Prpt_tra where vouc_no=" & Val(rs9!VOUC_NO) & " and acct_code='" & rs!acct_code & "'"
          Set rs10 = db.OpenRecordset(Sqlqry10, dbOpenDynaset)
          If rs10.RecordCount <> 0 Then
           rs10.MoveFirst
            Do Until rs10.EOF
             Opprptcr = Opprptcr + rs10!amount
             rs10.MoveNext
            Loop
           End If
         rs9.MoveNext
        Loop
        End If
        
        ' Opening Balance Pdc Payments before From Date
        SQLQRY11 = "select * from Ppmt_mas where Cheque_Dt<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and not isnull(posting_Dt)"
        Set rs11 = db.OpenRecordset(SQLQRY11, dbOpenDynaset)
         If rs11.RecordCount <> 0 Then
          rs11.MoveFirst
           Do Until rs11.EOF
            If rs!acct_code = rs11!bank_code Then
              Opppmtcr = Opppmtcr + rs11!ttl_amount
            End If
           
             SQLQRY12 = "Select * from Prpt_tra where vouc_no=" & Val(rs11!VOUC_NO) & " and acct_code='" & rs!acct_code & "'"
             Set rs12 = db.OpenRecordset(SQLQRY12, dbOpenDynaset)
             If rs12.RecordCount <> 0 Then
              rs12.MoveFirst
               Do Until rs12.EOF
                Opppmtdb = Opppmtdb + rs12!amount
                rs12.MoveNext
               Loop
             End If
          rs11.MoveNext
          Loop
        End If
        
       ' Journal Debit Amount before From Date
        Sqlqry13 = "select * from Jrnl_tra where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "' and dc_code ='D'"
        Set rs13 = db.OpenRecordset(Sqlqry13, dbOpenDynaset)
        If rs13.RecordCount <> 0 Then
         rs13.MoveFirst
        Do Until rs13.EOF
         Opjrnldb = Opjrnldb + rs13!damount
         rs13.MoveNext
        Loop
        End If
         
       ' Journal Credit Amount before From Date
        Sqlqry14 = "select * from Jrnl_tra where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "' and dc_code ='C'"
        Set rs14 = db.OpenRecordset(Sqlqry14, dbOpenDynaset)
        If rs14.RecordCount <> 0 Then
         rs14.MoveFirst
        Do Until rs14.EOF
         Opjrnlcr = Opjrnlcr + rs14!camount
         rs14.MoveNext
        Loop
        End If
         
       ' opening balance  debit note (credit) before From Date
        Sqlqry15 = "select * from debt_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "'"
        Set rs15 = db.OpenRecordset(Sqlqry15, dbOpenDynaset)
        If rs15.RecordCount <> 0 Then
         rs15.MoveFirst
        Do Until rs15.EOF
         opdbntcr = opdbntcr + rs15!amount
         rs15.MoveNext
        Loop
        End If
         
       ' Opening balance debit note (debit)  before From Date
        sqlqry16 = "select * from debt_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and cust_no='" & rs!acct_code & "'"
        Set rs16 = db.OpenRecordset(sqlqry16, dbOpenDynaset)
        If rs16.RecordCount <> 0 Then
         rs16.MoveFirst
        Do Until rs16.EOF
         opdbntdb = opdbntdb + rs16!amount
         rs16.MoveNext
        Loop
        End If
        
       ' Opening balance  Credit note (debit) before From Date
        sqlqry17 = "select * from crdt_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "'"
        Set rs17 = db.OpenRecordset(sqlqry17, dbOpenDynaset)
        If rs17.RecordCount <> 0 Then
         rs17.MoveFirst
        Do Until rs17.EOF
         Opcrntdb = Opcrntdb + rs17!amount
         rs17.MoveNext
        Loop
        End If
         
       ' Opening balance credit note (credit)  before From Date
        sqlqry18 = "select * from crdt_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and supp_no='" & rs!acct_code & "'"
        Set rs18 = db.OpenRecordset(sqlqry18, dbOpenDynaset)
        If rs18.RecordCount <> 0 Then
         rs18.MoveFirst
        Do Until rs18.EOF
         Opcrntcr = Opcrntcr + rs18!amount
         rs18.MoveNext
        Loop
        End If
                  
         ' Opening balance credit Sales  before From Date
        sqlqry19 = "select * from crsl_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs19 = db.OpenRecordset(sqlqry19, dbOpenDynaset)
        If rs19.RecordCount <> 0 Then
         rs19.MoveFirst
        Do Until rs19.EOF
         Opcrslcr = Opcrslcr + rs19!namount
         rs19.MoveNext
        Loop
        End If
                           
        ' Opening balance credit Purchases  before From Date
        sqlqry20 = "select * from crpr_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs20 = db.OpenRecordset(sqlqry20, dbOpenDynaset)
        If rs20.RecordCount <> 0 Then
         rs20.MoveFirst
        Do Until rs20.EOF
         Opcrprdb = Opcrprdb + rs20!namount
         rs20.MoveNext
        Loop
        End If
        
        ' Opening balance Cash Sales(credit)  before From Date
        Sqlqry21 = "select * from casl_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs21 = db.OpenRecordset(Sqlqry21, dbOpenDynaset)
        If rs21.RecordCount <> 0 Then
         rs21.MoveFirst
        Do Until rs21.EOF
         OpcaslCr = OpcaslCr + rs21!namount
         rs21.MoveNext
        Loop
        End If
                 
         ' Opening balance Cash Sales(debit)  before From Date
        Sqlqry22 = "select * from casl_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and cash_code='" & rs!acct_code & "'"
        Set rs22 = db.OpenRecordset(Sqlqry22, dbOpenDynaset)
        If rs22.RecordCount <> 0 Then
         rs22.MoveFirst
        Do Until rs22.EOF
         Opcasldb = Opcasldb + rs22!namount
         rs22.MoveNext
        Loop
        End If
        
        ' Opening balance Cash Purchases(debit)  before From Date
        Sqlqry23 = "select * from capr_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs23 = db.OpenRecordset(Sqlqry23, dbOpenDynaset)
        If rs23.RecordCount <> 0 Then
         rs23.MoveFirst
        Do Until rs23.EOF
         Opcaprdb = Opcaprdb + rs23!namount
         rs23.MoveNext
        Loop
        End If
                 
         ' Opening balance Cash Purchases(Credit)  before From Date
        Sqlqry24 = "select * from capr_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and cash_code='" & rs!acct_code & "'"
        Set rs24 = db.OpenRecordset(Sqlqry24, dbOpenDynaset)
        If rs24.RecordCount <> 0 Then
         rs24.MoveFirst
        Do Until rs24.EOF
         Opcaprcr = Opcaprcr + rs24!namount
         rs24.MoveNext
        Loop
        End If
                  
         Ttlopbal = Opbal - Opcrptcr + Opcrptdb - Opcpmtcr + Opcpmtdb - Opbrptcr + Opbrptdb + Opbpmtdb - Opbpmtcr + Opprptdb - Opprptcr + Opppmtdb _
                 - Opppmtcr + Opjrnldb - Opjrnlcr + opdbntdb - opdbntcr + Opcrntdb - Opcrntcr - Opcrslcr + Opcrprdb - OpcaslCr + Opcasldb - Opcaprcr + Opcasldb
                
        
       End If
                
          Sqlqry25 = "Update acct_mas set close_bal =" & Ttlopbal & " where Acct_code ='" & rs!acct_code & "'"
          ws.BeginTrans
          db.Execute (Sqlqry25)
          ws.CommitTrans
       
       rs.MoveNext
       Loop
       DOCUSTPOST
       DOSUPPPOST
       SalClbal = 0
       PurClbal = 0
       ProClbal = 0
       VehClbal = 0
       Advclbal = 0
        Sqlqry80 = " Select * from  acct_mas where acct_code >='105000' and acct_code<='105499' "
        Set rs73 = db.OpenRecordset(Sqlqry80, dbOpenDynaset)
        If rs73.RecordCount <> 0 Then
          rs73.MoveFirst
          Do Until rs73.EOF
          Advclbal = Advclbal + rs73!Close_bal
          rs73.MoveNext
          Loop
        End If
          Sqlqry81 = " Update acct_mas set close_bal=" & Advclbal & " where acct_code ='105000'"
            ws.BeginTrans
            db.Execute (Sqlqry81)
            ws.CommitTrans
          Sqlqry82 = " update acct_mas set close_bal =" & 0 & " where acct_code>='105001' and acct_code<='105499'"
            ws.BeginTrans
            db.Execute (Sqlqry82)
            ws.CommitTrans
            
        Sqlqry83 = " Select * from  acct_mas where acct_code >='417000' and acct_code<'417100' "
        Set rs74 = db.OpenRecordset(Sqlqry83, dbOpenDynaset)
        If rs74.RecordCount <> 0 Then
          rs74.MoveFirst
           Do Until rs74.EOF
          VehClbal = VehClbal + rs74!Close_bal
          rs74.MoveNext
          Loop
        End If
         Sqlqry84 = " Update acct_mas set close_bal=" & VehClbal & " where acct_code ='417000'"
         ws.BeginTrans
         db.Execute (Sqlqry84)
         ws.CommitTrans
         Sqlqry85 = " update acct_mas set close_bal =" & 0 & " where acct_code>='417001' and acct_code<'417100'"
         ws.BeginTrans
         db.Execute (Sqlqry85)
         ws.CommitTrans
         
        Sqlqry86 = " Select * from  acct_mas where acct_code >='204000' and acct_code<'204100' "
        Set rs75 = db.OpenRecordset(Sqlqry86, dbOpenDynaset)
        If rs75.RecordCount <> 0 Then
          rs75.MoveFirst
          Do Until rs75.EOF
          ProClbal = ProClbal + rs75!Close_bal
          rs75.MoveNext
          Loop
        End If
         Sqlqry87 = " Update acct_mas set close_bal=" & ProClbal & " where acct_code ='204000'"
         ws.BeginTrans
         db.Execute (Sqlqry87)
         ws.CommitTrans
         Sqlqry88 = " update acct_mas set close_bal =" & 0 & " where acct_code>='204001' and acct_code<'204100'"
         ws.BeginTrans
         db.Execute (Sqlqry88)
         ws.CommitTrans
         
        Sqlqry89 = " Select * from  acct_mas where acct_code >='401000' and acct_code<'403000' "
        Set rs76 = db.OpenRecordset(Sqlqry89, dbOpenDynaset)
        If rs76.RecordCount <> 0 Then
          rs76.MoveFirst
          Do Until rs76.EOF
          PurClbal = PurClbal + rs76!Close_bal
          rs76.MoveNext
          Loop
        End If
         Sqlqry90 = " Update acct_mas set close_bal=" & PurClbal & " where acct_code ='401000'"
         ws.BeginTrans
         db.Execute (Sqlqry90)
         ws.CommitTrans
         Sqlqry91 = " update acct_mas set close_bal =" & 0 & " where acct_code>='401001' and acct_code<'403000'"
         ws.BeginTrans
         db.Execute (Sqlqry91)
         ws.CommitTrans
         
        Sqlqry92 = " Select * from  acct_mas where acct_code >='301000' and acct_code<'302000' "
        Set rs77 = db.OpenRecordset(Sqlqry92, dbOpenDynaset)
        If rs77.RecordCount <> 0 Then
          rs77.MoveFirst
          Do Until rs77.EOF
          SalClbal = SalClbal + rs77!Close_bal
          rs77.MoveNext
          Loop
        End If
         Sqlqry93 = " Update acct_mas set close_bal=" & SalClbal & " where acct_code ='301000'"
         ws.BeginTrans
         db.Execute (Sqlqry93)
         ws.CommitTrans
         Sqlqry94 = " update acct_mas set close_bal =" & 0 & " where acct_code>='301001' and acct_code<'302000'"
         ws.BeginTrans
         db.Execute (Sqlqry94)
         ws.CommitTrans
         
         
     End If
     
     With CrystalReport1
       .DataFiles(0) = "C:\uday\udfin.mdb"
       .ReportFileName = "C:\uday\TrialBalance.rpt"
       .Formulas(0) = "xxx='" & txtdatefrom & "'"
       .WindowMaxButton = True
       .WindowState = crptMaximized
       .Action = 1
     End With

     
   Else
        MsgBox "Improper Dates", vbDefaultButton1, "Invalid entry"
        Exit Sub
  End If
  
  
End Sub
Private Sub Form_Load()
txtdatefrom.Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdDisplay.SetFocus
End Sub

Private Function ValidateData()
 ValidateData = False

If IsDate(txtdatefrom) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
   ValidateData = True
End If

End Function
Private Sub textclear()
 txtdatefrom.Text = Format(Now, "dd/mm/yyyy")
End Sub
Private Sub docashbal()
      
       'Opbal = 0
       Opbpmt = 0
       Opbrpt = 0
       opsal = 0
       oprec = 0
       oppur = 0
       oppay = 0
       Opjdb = 0
       Opjcr = 0
       Ttlopbal = 0
       ' Total Amount of Sales before From date
        Sqlqry26 = "select * from casl_mas where Cash_code='103001' and date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "#"
        Set rs25 = db.OpenRecordset(Sqlqry26, dbOpenDynaset)
        If rs25.RecordCount <> 0 Then
          rs25.MoveFirst
         Do Until rs25.EOF
          opsal = opsal + rs25!namount
          rs25.MoveNext
         Loop
        End If
        
        ' Total Amount of Purchase before From date
        Sqlqry27 = "select * from capr_mas where Cash_code='103001' and date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "#"
        Set rs26 = db.OpenRecordset(Sqlqry27, dbOpenDynaset)
        If rs26.RecordCount <> 0 Then
         rs26.MoveFirst
         Do Until rs26.EOF
          oppur = oppur + rs26!namount
          rs26.MoveNext
         Loop
        End If
        
       ' Total cash Receipts before From date
        Sqlqry28 = "select * from crpt_mas where Cash_code='103001' and date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "#"
        Set rs27 = db.OpenRecordset(Sqlqry28, dbOpenDynaset)
        If rs27.RecordCount <> 0 Then
         rs27.MoveFirst
        Do Until rs27.EOF
          oprec = oprec + rs27!ttl_amount
          rs27.MoveNext
        Loop
        End If
        
        ' Total cash Payments before To date
        Sqlqry29 = "select * from cpmt_mas where Cash_code ='103001' and date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "#"
        Set rs28 = db.OpenRecordset(Sqlqry29, dbOpenDynaset)
        If rs28.RecordCount <> 0 Then
         rs28.MoveFirst
        Do Until rs28.EOF
         oppay = oppay + rs28!ttl_amount
         rs28.MoveNext
        Loop
        End If
        
        ' Bank Payment  debit Before before From Date
        sqlqry16 = "select * from bpmt_tra where Acct_Code ='103001' and  date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "#"
        Set rs11 = db.OpenRecordset(sqlqry16, dbOpenDynaset)
        If rs11.RecordCount <> 0 Then
         rs11.MoveFirst
        Do Until rs11.EOF
         Opbpmt = Opbpmt + rs11!amount
         rs11.MoveNext
        Loop
        End If
         
         ' Bank Receipt Credit Before before From Date
        sqlqry17 = "select * from brpt_tra where Acct_Code ='103001' and  date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "#"
        Set rs12 = db.OpenRecordset(sqlqry17, dbOpenDynaset)
        If rs12.RecordCount <> 0 Then
         rs12.MoveFirst
        Do Until rs12.EOF
         Opbrpt = Opbrpt + rs12!amount
         rs12.MoveNext
        Loop
        End If
         
         ' Journal Debit Amount before From Date
        Sqlqry78 = "select * from Jrnl_tra where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='103001' and dc_code ='D'"
        Set rs71 = db.OpenRecordset(Sqlqry78, dbOpenDynaset)
        If rs71.RecordCount <> 0 Then
         rs71.MoveFirst
        Do Until rs71.EOF
         Opjdb = Opjdb + rs71!damount
         rs71.MoveNext
        Loop
        End If
         
       ' Journal Credit Amount before From Date
        Sqlqry79 = "select * from Jrnl_tra where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='103001' and dc_code ='C'"
        Set rs72 = db.OpenRecordset(Sqlqry79, dbOpenDynaset)
        If rs72.RecordCount <> 0 Then
         rs72.MoveFirst
        Do Until rs72.EOF
         Opjcr = Opjcr + rs72!camount
         rs72.MoveNext
        Loop
        End If
              
         
         Ttlopbal = Opbal + Opbpmt - Opbrpt + opsal + oprec - oppur - oppay + Opjdb - Opjcr
      
End Sub

Public Sub dobankbal()
               
        Ttlopbal = 0
        ' Opbal = 0
        Opcasl = 0
        opcpmt = 0
        opcrpt = 0
        Opbrpt = 0
        Opprpt = 0
        Opbrpta = 0
        Opjdb = 0
        Opbpmt = 0
        Opbpmta = 0
        Opppmt = 0
        Opjcr = 0
        
        Sqlqry28 = "select * from casl_mas where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and cash_code='" & rs!acct_code & "'"
        Set rs13 = db.OpenRecordset(Sqlqry28, dbOpenDynaset)
        If rs13.RecordCount <> 0 Then
          rs13.MoveFirst
         Do Until rs13.EOF
          Opcasl = Opcasl + rs13!namount
          rs13.MoveNext
         Loop
        End If
        
        ' Bank Receipt bank code  before From date
        Sqlqry30 = "select * from brpt_mas where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and bank_code='" & rs!acct_code & "'"
        Set rs29 = db.OpenRecordset(Sqlqry30, dbOpenDynaset)
        If rs29.RecordCount <> 0 Then
          rs29.MoveFirst
         Do Until rs29.EOF
          Opbrpt = Opbrpt + rs29!ttl_amount
          rs29.MoveNext
         Loop
        End If
        
         ' Bank Receipt Acct Code before From date
        Sqlqry30 = "select * from brpt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "'"
        Set rs29 = db.OpenRecordset(Sqlqry30, dbOpenDynaset)
        If rs29.RecordCount <> 0 Then
          rs29.MoveFirst
         Do Until rs29.EOF
          Opbrpta = Opbrpta + rs29!amount
          rs29.MoveNext
         Loop
        End If
        
        ' Bank Payment Bank Code before From date
        Sqlqry31 = "select * from bpmt_mas where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and bank_code='" & rs!acct_code & "'"
        Set rs30 = db.OpenRecordset(Sqlqry31, dbOpenDynaset)
        If rs30.RecordCount <> 0 Then
         rs30.MoveFirst
         Do Until rs30.EOF
          Opbpmt = Opbpmt + rs30!ttl_amount
          rs30.MoveNext
         Loop
        End If
        
        ' Bank Payment Account Code before From date
        Sqlqry31 = "select * from bpmt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs30 = db.OpenRecordset(Sqlqry31, dbOpenDynaset)
        If rs30.RecordCount <> 0 Then
         rs30.MoveFirst
         Do Until rs30.EOF
          Opbpmta = Opbpmta + rs30!amount
          rs30.MoveNext
         Loop
        End If
                    
      ' Cash receipt (credit) before From date
        sqlqry20 = "select * from crpt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs13 = db.OpenRecordset(sqlqry20, dbOpenDynaset)
        If rs13.RecordCount <> 0 Then
         rs13.MoveFirst
         Do Until rs13.EOF
          opcrpt = opcrpt + rs13!amount
          rs13.MoveNext
         Loop
        End If
        
     ' Cash Payment (Debit) before From date
        Sqlqry21 = "select * from cpmt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs!acct_code & "'"
        Set rs14 = db.OpenRecordset(Sqlqry21, dbOpenDynaset)
        If rs14.RecordCount <> 0 Then
         rs14.MoveFirst
         Do Until rs14.EOF
          opcpmt = opcpmt + rs14!amount
          rs14.MoveNext
         Loop
        End If
        
       ' Pdc Receipts before From date
        Sqlqry32 = "select * from prpt_mas where Cheque_dt<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and bank_code='" & rs!acct_code & "' and not isnull(posting_dt)"
        Set rs31 = db.OpenRecordset(Sqlqry32, dbOpenDynaset)
        If rs31.RecordCount <> 0 Then
         rs31.MoveFirst
        Do Until rs31.EOF
          Opprpt = Opprpt + rs31!ttl_amount
          rs31.MoveNext
        Loop
        End If
        
        ' Pdc Payments before From Date
        Sqlqry33 = "select * from Ppmt_mas where Cheque_Dt<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and bank_code='" & rs!acct_code & "' and not isnull(posting_Dt)"
        Set rs32 = db.OpenRecordset(Sqlqry33, dbOpenDynaset)
        If rs32.RecordCount <> 0 Then
         rs32.MoveFirst
        Do Until rs32.EOF
         Opppmt = Opppmt + rs32!ttl_amount
         rs32.MoveNext
        Loop
        End If
         
       ' Journal Debit Amount before From Date
        Sqlqry34 = "select * from Jrnl_tra where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "' and dc_code ='D'"
        Set rs33 = db.OpenRecordset(Sqlqry34, dbOpenDynaset)
        If rs33.RecordCount <> 0 Then
         rs33.MoveFirst
        Do Until rs33.EOF
         Opjdb = Opjdb + rs33!damount
         rs33.MoveNext
        Loop
        End If
         
       ' Journal Credit Amount before From Date
        Sqlqry35 = "select * from Jrnl_tra where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs!acct_code & "' and dc_code ='C'"
        Set rs34 = db.OpenRecordset(Sqlqry35, dbOpenDynaset)
        If rs34.RecordCount <> 0 Then
         rs34.MoveFirst
        Do Until rs34.EOF
         Opjcr = Opjcr + rs34!camount
         rs34.MoveNext
        Loop
        End If
         
        Ttlopbal = Opbal + Opcasl + opcpmt - opcrpt + Opbrpt - Opbrpta + Opprpt + Opjdb - Opbpmt + Opbpmta - Opppmt - Opjcr
End Sub
Private Sub DOCUSTPOST()
    
   TTLCUST = 0
   Sqlqry36 = " select * from Cust_Fin order by cust_no"
   Set rs35 = db.OpenRecordset(Sqlqry36, dbOpenDynaset)
   If rs35.RecordCount = 0 Then
     MsgBox " Customer Code not found in Cust_Fin"
     Exit Sub
    Else
     rs35.MoveFirst
      Do Until rs35.EOF
         
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
                Tpdc = 0
                
        Opbal = rs35!open_bal
       ' Cash Receipt before From date
        Sqlqry37 = "select * from crpt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs35!cust_no & "'"
        Set rs36 = db.OpenRecordset(Sqlqry37, dbOpenDynaset)
        If rs36.RecordCount <> 0 Then
          rs36.MoveFirst
        Do Until rs36.EOF
          opcrpt = opcrpt + rs36!amount
          rs36.MoveNext
        Loop
        End If
        
        ' Cash Payment before From date
        Sqlqry38 = "select * from cpmt_Tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs35!cust_no & "'"
        Set rs37 = db.OpenRecordset(Sqlqry38, dbOpenDynaset)
        If rs37.RecordCount <> 0 Then
          rs37.MoveFirst
         Do Until rs37.EOF
          opcpmt = opcpmt + rs37!amount
          rs37.MoveNext
         Loop
        End If
                
        ' Bank Receipt before From date
        Sqlqry39 = "select * from brpt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs35!cust_no & "'"
        Set rs38 = db.OpenRecordset(Sqlqry39, dbOpenDynaset)
        If rs38.RecordCount <> 0 Then
          rs38.MoveFirst
         Do Until rs38.EOF
          Opbrpt = Opbrpt + rs38!amount
          rs38.MoveNext
         Loop
        End If
        
        ' Bank Payment before From date
        Sqlqry40 = "select * from bpmt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs35!cust_no & "'"
        Set rs39 = db.OpenRecordset(Sqlqry40, dbOpenDynaset)
        If rs39.RecordCount <> 0 Then
         rs39.MoveFirst
         Do Until rs39.EOF
          Opbpmt = Opbpmt + rs39!amount
          rs39.MoveNext
         Loop
        End If
        
       ' Pdc Receipts before From date
        Sqlqry41 = "select * from prpt_mas where Cheque_Dt<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and not isnull(posting_dt)"
        Set rs40 = db.OpenRecordset(Sqlqry41, dbOpenDynaset)
        If rs40.RecordCount <> 0 Then
          rs40.MoveFirst
         Do Until rs40.EOF
           Sqlqry42 = "Select * from Prpt_tra where Vouc_no=" & Val(rs40!VOUC_NO) & " and acct_code='" & rs35!cust_no & "'"
           Set rs41 = db.OpenRecordset(Sqlqry42, dbOpenDynaset)
           If rs41.RecordCount <> 0 Then
            rs41.MoveFirst
             Do Until rs41.EOF
              Opprpt = Opprpt + rs41!amount
              rs41.MoveNext
             Loop
           End If
          rs40.MoveNext
         Loop
        End If
        
        ' Pdc Payments before From Date
        Sqlqry43 = "select * from Ppmt_mas where Cheque_Dt<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and not isnull(posting_Dt)"
        Set rs42 = db.OpenRecordset(Sqlqry43, dbOpenDynaset)
        If rs42.RecordCount <> 0 Then
         rs42.MoveFirst
         Do Until rs42.EOF
          Sqlqry44 = "Select * from Ppmt_tra where Vouc_no=" & Val(rs42!VOUC_NO) & " and acct_code='" & rs35!cust_no & "'"
          Set rs43 = db.OpenRecordset(Sqlqry44, dbOpenDynaset)
           If rs43.RecordCount <> 0 Then
            rs43.MoveFirst
             Do Until rs43.EOF
              Opppmt = Opppmt + rs43!amount
              rs43.MoveNext
             Loop
           End If
         rs42.MoveNext
         Loop
        End If
         
       ' Journal Debit Amount before From Date
        Sqlqry45 = "select * from Jrnl_tra where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs35!cust_no & "' and dc_code ='D'"
        Set rs44 = db.OpenRecordset(Sqlqry45, dbOpenDynaset)
        If rs44.RecordCount <> 0 Then
         rs44.MoveFirst
        Do Until rs44.EOF
         opjbd = opjbd + rs44!damount
         rs44.MoveNext
        Loop
        End If
         
       ' Journal Credit Amount before From Date
        Sqlqry46 = "select * from Jrnl_tra where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs35!cust_no & "' and dc_code ='C'"
        Set rs45 = db.OpenRecordset(Sqlqry46, dbOpenDynaset)
        If rs45.RecordCount <> 0 Then
         rs45.MoveFirst
        Do Until rs45.EOF
         Opjbc = Opjbc + rs45!camount
         rs45.MoveNext
        Loop
        End If
         
        ' Opening balance  debit note (credit) before From Date
        Sqlqry47 = "select * from debt_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs35!cust_no & "'"
        Set rs46 = db.OpenRecordset(Sqlqry47, dbOpenDynaset)
        If rs46.RecordCount <> 0 Then
         rs46.MoveFirst
        Do Until rs46.EOF
         opdbntcr = opdbntcr + rs46!amount
         rs46.MoveNext
        Loop
        End If
         
       ' Opening balance debit note (debit)  before From Date
        Sqlqry48 = "select * from debt_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and cust_no='" & rs35!cust_no & "'"
        Set rs47 = db.OpenRecordset(Sqlqry48, dbOpenDynaset)
        If rs47.RecordCount <> 0 Then
         rs47.MoveFirst
        Do Until rs47.EOF
         opdbntdb = opdbntdb + rs47!amount
         rs47.MoveNext
        Loop
        End If
        
       ' Opening balance  Credit note (debit) before From Date
        Sqlqry49 = "select * from crdt_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs35!cust_no & "'"
        Set rs48 = db.OpenRecordset(Sqlqry49, dbOpenDynaset)
        If rs48.RecordCount <> 0 Then
         rs48.MoveFirst
        Do Until rs48.EOF
         Opcrntdb = Opcrntdb + rs48!amount
         rs48.MoveNext
        Loop
        End If
         
       ' Opening balance credit note (credit)  before From Date
        Sqlqry50 = "select * from crdt_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and supp_no='" & rs35!cust_no & "'"
        Set rs49 = db.OpenRecordset(Sqlqry50, dbOpenDynaset)
        If rs49.RecordCount <> 0 Then
         rs49.MoveFirst
        Do Until rs49.EOF
         Opcrntcr = Opcrntcr + rs49!amount
         rs49.MoveNext
        Loop
        End If
                  
         ' Opening balance credit Sales  before From Date
        Sqlqry51 = "select * from crsl_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and cust_no='" & rs35!cust_no & "'"
        Set rs50 = db.OpenRecordset(Sqlqry51, dbOpenDynaset)
        If rs50.RecordCount <> 0 Then
         rs50.MoveFirst
        Do Until rs50.EOF
         Opcrsl = Opcrsl + rs50!namount
         rs50.MoveNext
        Loop
        End If
                           
        ' Opening balance credit Purchases  before From Date
        Sqlqry52 = "select * from crpr_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and supp_no='" & rs35!cust_no & "'"
        Set rs51 = db.OpenRecordset(Sqlqry52, dbOpenDynaset)
        If rs51.RecordCount <> 0 Then
         rs51.MoveFirst
        Do Until rs51.EOF
         Opcrpr = Opcrpr + rs51!namount
         rs51.MoveNext
        Loop
        End If
         
        Ttlopbal = Opbal - opcrpt + opcpmt - Opbrpt + Opbpmt - Opprpt + Opppmt + opjbd - Opjbc + opdbntdb - opdbntcr _
                    + Opcrntdb - Opcrntcr + Opcrsl - Opcrpr
        TTLCUST = TTLCUST + Ttlopbal
     rs35.MoveNext
     Loop
         
       ' Pending Post Dated Cheques Received
      
        Sqlqry53 = "select * from prpt_mas where isnull(posting_dt) AND CHEQUE_dT>=#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# "
        Set rs52 = db.OpenRecordset(Sqlqry53, dbOpenDynaset)
        If rs52.RecordCount <> 0 Then
          rs52.MoveFirst
         Do Until rs52.EOF
             Tpdc = Tpdc + rs52!ttl_amount
             rs52.MoveNext
         Loop
        End If
          TTLCUST = TTLCUST - Tpdc
          
        Sqlqry54 = "Update acct_mas set close_bal=" & TTLCUST & " where acct_code ='102000'"
        ws.BeginTrans
        db.Execute (Sqlqry54)
        ws.CommitTrans
       ' 103501 = Bills Receivable
       
        Sqlqry54 = "SELECT * FROM ACCT_MAS WHERE ACCT_CODE='103501'"
        Set rs54 = db.OpenRecordset(Sqlqry54, dbOpenDynaset)
         If rs54.RecordCount <> 0 Then
          rs54.MoveFirst
          If IsNull(rs54!open_bal) = True Then
            X = 0
          Else
            X = rs54!open_bal
          End If
         End If
         
        
              
        Sqlqry55 = "Update acct_mas set Close_bal=" & X + Tpdc & " where acct_code='103501'"
        ws.BeginTrans
        db.Execute (Sqlqry55)
        ws.CommitTrans
       
     End If
     
           
            
End Sub
Private Sub DOSUPPPOST()
   Ttlsupp = 0
   Sqlqry57 = " select * from Supp_Fin order by Supp_no"
   Set rs54 = db.OpenRecordset(Sqlqry57, dbOpenDynaset)
   If rs54.RecordCount = 0 Then
     MsgBox " Supplier Code not found in Supp_Fin"
     Exit Sub
    Else
     rs54.MoveFirst
      Do Until rs54.EOF
         
                Opbal = 0
                Opbrpt = 0
                Opbpmt = 0
                Opprpt = 0
                Opppmt = 0
                opjbd = 0
                Opjbc = 0
                opcrpt = 0
                opcpmt = 0
                opdbntdb = 0
                opdbntcr = 0
                Opcrntdb = 0
                Opcrntcr = 0
                Opcrsl = 0
                Opcrpr = 0
                Ttlopbal = 0
                Tpdc = 0
                
                
         Opbal = rs54!open_bal
       ' Cash Receipt before From date
        Sqlqry58 = "select * from crpt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs54!Supp_no & "'"
        Set rs55 = db.OpenRecordset(Sqlqry58, dbOpenDynaset)
        If rs55.RecordCount <> 0 Then
          rs55.MoveFirst
        Do Until rs55.EOF
          opcrpt = opcrpt + rs55!amount
          rs55.MoveNext
         Loop
        End If
        
        ' Cash Payment before From date
        Sqlqry59 = "select * from cpmt_Tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs54!Supp_no & "'"
        Set rs56 = db.OpenRecordset(Sqlqry59, dbOpenDynaset)
        If rs56.RecordCount <> 0 Then
          rs56.MoveFirst
         Do Until rs56.EOF
          opcpmt = opcpmt + rs56!amount
          rs56.MoveNext
         Loop
        End If
                
        ' Bank Receipt before From date
        Sqlqry60 = "select * from brpt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and acct_code='" & rs54!Supp_no & "'"
        Set rs57 = db.OpenRecordset(Sqlqry60, dbOpenDynaset)
        If rs57.RecordCount <> 0 Then
          rs57.MoveFirst
         Do Until rs57.EOF
          Opbrpt = Opbrpt + rs57!amount
          rs57.MoveNext
         Loop
        End If
        
        ' Bank Payment before From date
        Sqlqry61 = "select * from bpmt_tra where date< #" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs54!Supp_no & "'"
        Set rs58 = db.OpenRecordset(Sqlqry61, dbOpenDynaset)
        If rs58.RecordCount <> 0 Then
         rs58.MoveFirst
         Do Until rs58.EOF
          Opbpmt = Opbpmt + rs58!amount
          rs58.MoveNext
         Loop
        End If
        
       ' Pdc Receipts before From date
        Sqlqry62 = "select * from prpt_mas where Cheque_dt<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and not isnull(posting_dt)"
        Set rs59 = db.OpenRecordset(Sqlqry62, dbOpenDynaset)
        If rs59.RecordCount <> 0 Then
          rs59.MoveFirst
         Do Until rs59.EOF
           Sqlqry63 = "Select * from Prpt_tra where Vouc_no=" & Val(rs59!VOUC_NO) & " and acct_code='" & rs54!Supp_no & "'"
           Set rs60 = db.OpenRecordset(Sqlqry63, dbOpenDynaset)
           If rs60.RecordCount <> 0 Then
            rs60.MoveFirst
             Do Until rs60.EOF
              Opprpt = Opprpt + rs60!amount
              rs60.MoveNext
             Loop
           End If
          rs59.MoveNext
         Loop
        End If
        
        ' Pdc Payments before From Date
        Sqlqry64 = "select * from Ppmt_mas where Cheque_dt<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and not isnull(posting_Dt)"
        Set rs61 = db.OpenRecordset(Sqlqry64, dbOpenDynaset)
        If rs61.RecordCount <> 0 Then
         rs61.MoveFirst
         Do Until rs61.EOF
          Sqlqry65 = "Select * from Ppmt_tra where Vouc_no=" & Val(rs61!VOUC_NO) & " and acct_code='" & rs54!Supp_no & "'"
          Set rs62 = db.OpenRecordset(Sqlqry65, dbOpenDynaset)
           If rs62.RecordCount <> 0 Then
            rs62.MoveFirst
             Do Until rs62.EOF
              Opppmt = Opppmt + rs62!amount
              rs62.MoveNext
             Loop
           End If
         rs61.MoveNext
         Loop
        End If
         
         
       ' Journal Debit Amount before From Date
        Sqlqry66 = "select * from Jrnl_tra where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs54!Supp_no & "' and dc_code ='D'"
        Set rs63 = db.OpenRecordset(Sqlqry66, dbOpenDynaset)
        If rs63.RecordCount <> 0 Then
         rs63.MoveFirst
        Do Until rs63.EOF
         opjbd = opjbd + rs63!damount
         rs63.MoveNext
        Loop
        End If
         
       ' Journal Credit Amount before From Date
        Sqlqry67 = "select * from Jrnl_tra where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs54!Supp_no & "' and dc_code ='C'"
        Set rs64 = db.OpenRecordset(Sqlqry67, dbOpenDynaset)
        If rs64.RecordCount <> 0 Then
         rs64.MoveFirst
        Do Until rs64.EOF
         Opjbc = Opjbc + rs64!camount
         rs64.MoveNext
        Loop
        End If
         
        ' opening balance  debit note (credit) before From Date
        Sqlqry68 = "select * from debt_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs54!Supp_no & "'"
        Set rs65 = db.OpenRecordset(Sqlqry68, dbOpenDynaset)
        If rs65.RecordCount <> 0 Then
         rs65.MoveFirst
        Do Until rs65.EOF
         opdbntcr = opdbntcr + rs65!amount
         rs65.MoveNext
        Loop
        End If
         
       ' Opening balance debit note (debit)  before From Date
        Sqlqry69 = "select * from debt_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and cust_no='" & rs54!Supp_no & "'"
        Set rs66 = db.OpenRecordset(Sqlqry69, dbOpenDynaset)
        If rs66.RecordCount <> 0 Then
         rs66.MoveFirst
        Do Until rs66.EOF
         opdbntdb = opdbntdb + rs66!amount
         rs66.MoveNext
        Loop
        End If
        
       ' Opening balance  Credit note (debit) before From Date
        Sqlqry70 = "select * from crdt_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and Acct_code='" & rs54!Supp_no & "'"
        Set rs67 = db.OpenRecordset(Sqlqry70, dbOpenDynaset)
        If rs67.RecordCount <> 0 Then
         rs67.MoveFirst
        Do Until rs67.EOF
         Opcrntdb = Opcrntdb + rs67!amount
         rs67.MoveNext
        Loop
        End If
         
       ' Opening balance credit note (credit)  before From Date
        Sqlqry71 = "select * from crdt_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and supp_no='" & rs54!Supp_no & "'"
        Set rs68 = db.OpenRecordset(Sqlqry71, dbOpenDynaset)
        If rs68.RecordCount <> 0 Then
         rs68.MoveFirst
        Do Until rs68.EOF
         Opcrntcr = Opcrntcr + rs68!amount
         rs68.MoveNext
        Loop
        End If
                  
         ' Opening balance credit Sales  before From Date
        Sqlqry72 = "select * from crsl_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and cust_no='" & rs54!Supp_no & "'"
        Set rs69 = db.OpenRecordset(Sqlqry72, dbOpenDynaset)
        If rs69.RecordCount <> 0 Then
         rs69.MoveFirst
        Do Until rs69.EOF
         Opcrsl = Opcrsl + rs69!namount
         rs69.MoveNext
        Loop
        End If
                           
        ' Opening balance credit Purchases  before From Date
        Sqlqry73 = "select * from crpr_mas where date<#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# and supp_no='" & rs54!Supp_no & "'"
        Set rs70 = db.OpenRecordset(Sqlqry73, dbOpenDynaset)
        If rs70.RecordCount <> 0 Then
         rs70.MoveFirst
        Do Until rs70.EOF
         Opcrpr = Opcrpr + rs70!namount
         rs70.MoveNext
        Loop
        End If
         
        Ttlopbal = Opbal - opcrpt + opcpmt - Opbrpt + Opbpmt - Opprpt + Opppmt + opjbd - Opjbc + opdbntdb - opdbntcr _
                    + Opcrntdb - Opcrntcr + Opcrsl - Opcrpr
        
        Ttlsupp = Ttlsupp + Ttlopbal
     
     rs54.MoveNext
     Loop
       ' Pending Post Dated Cheques Paid
        Sqlqry74 = "select * from ppmt_mas where isnull(posting_dt) AND CHEQUE_DT>=#" & DateValue(Format(txtdatefrom, "dd/mm/yyyy")) & "# "
        Set rs71 = db.OpenRecordset(Sqlqry74, dbOpenDynaset)
        If rs71.RecordCount <> 0 Then
          rs71.MoveFirst
         Do Until rs71.EOF
             Tpdc = Tpdc + rs71!ttl_amount
             rs71.MoveNext
         Loop
        End If
          Ttlsupp = Ttlsupp + Tpdc
          
        Sqlqry75 = "Update acct_mas set close_bal=" & Ttlsupp & " where acct_code ='202000'"
        ws.BeginTrans
        db.Execute (Sqlqry75)
        ws.CommitTrans
        X = 0
        Sqlqry75 = "SELECT * FROM ACCT_MAS WHERE ACCT_CODE='106001'"
        Set rs71 = db.OpenRecordset(Sqlqry75, dbOpenDynaset)
         If rs71.RecordCount <> 0 Then
          rs71.MoveFirst
          If IsNull(rs71!open_bal) = True Then
            X = 0
          Else
            X = rs71!open_bal
          End If
         End If
        Sqlqry76 = "Update acct_mas set Close_bal=" & X - Tpdc & " where acct_code='106001'"
        ws.BeginTrans
        db.Execute (Sqlqry76)
        ws.CommitTrans
        
     End If
 
End Sub
Private Sub doclosnil()
   Sqlqry77 = "Update acct_mas set close_bal=" & 0 & ""
        ws.BeginTrans
        db.Execute (Sqlqry77)
        ws.CommitTrans
  
End Sub
