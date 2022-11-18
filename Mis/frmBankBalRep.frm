VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmBankBalRep 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bank Balance Report"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   225
   ClientWidth     =   12060
   FillColor       =   &H00000040&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bank Account"
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
      Height          =   7455
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   8055
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Preview"
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
         Left            =   2160
         Picture         =   "frmBankBalRep.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6480
         Width           =   1095
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
         Height          =   735
         Left            =   3240
         Picture         =   "frmBankBalRep.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6480
         Width           =   1095
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
         Height          =   735
         Left            =   4320
         Picture         =   "frmBankBalRep.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6480
         Width           =   1095
      End
      Begin VB.ListBox lstBankCodes 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   6855
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   6120
         Top             =   6240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         Top             =   4800
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
         ForeColor       =   0
         HighlightColor  =   -2147483640
         Alignment       =   1
      End
      Begin PVMaskEditLib.PVMaskEdit txtdateto 
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   5400
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
         ForeColor       =   0
         HighlightColor  =   -2147483640
         Alignment       =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404080&
         BorderWidth     =   3
         X1              =   8040
         X2              =   0
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter Date From"
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
         Height          =   240
         Left            =   1920
         TabIndex        =   3
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter Date To"
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
         Height          =   240
         Left            =   1920
         TabIndex        =   2
         Top             =   5520
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmBankBalRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim Opbaldhs As Currency
Dim opsaldhs As Currency
Dim Opbrptdhs As Currency
Dim Opbrptadhs As Currency
Dim Opbpmtdhs As Currency
Dim Opbpmtadhs As Currency
Dim Opprptdhs As Currency
Dim Opppmtdhs As Currency
Dim Opjdbdhs As Currency
Dim Opjcrdhs As Currency
Dim Ttlopbaldhs As Currency
Dim opcpmtdhs As Currency
Dim opcrptdhs As Currency
Dim opdbntcrdhs As Currency
Dim Opbalusd As Currency
Dim opsalusd As Currency
Dim Opbrptusd As Currency
Dim Opbrptausd As Currency
Dim Opbpmtusd As Currency
Dim Opbpmtausd As Currency
Dim Opprptusd As Currency
Dim Opppmtusd As Currency
Dim Opjdbusd As Currency
Dim Opjcrusd As Currency
Dim opdbntcrusd As Currency
Dim Ttlopbalusd As Currency
Dim opcpmtusd As Currency
Dim opcrptusd As Currency

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
Dim Opcrntdbdhs As Currency
Dim Opcrntdbusd As Currency
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
        
        Sqlqry = " Delete * from bankreport"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
      Opbaldhs = 0
      Opbrptdhs = 0
      Opbpmtdhs = 0
      Opprptdhs = 0
      Opppmtdhs = 0
      opsaldhs = 0
      Opjdbdhs = 0
      Opjcrdhs = 0
      opcpmtdhs = 0
      opcrptdhs = 0
      Opbpmtadhs = 0
      Opbrptadhs = 0
      opdbntcrdhs = 0
      Opcrntdbdhs = 0
      Ttlopbaldhs = 0
      Opbalusd = 0
      Opbrptusd = 0
      Opbpmtusd = 0
      Opprptusd = 0
      Opppmtusd = 0
      opsalusd = 0
      Opjdbusd = 0
      Opjcrusd = 0
      opcpmtusd = 0
      opcrptusd = 0
      Opbpmtausd = 0
      Opbrptausd = 0
      opdbntcrusd = 0
      Opcrntdbusd = 0
      Ttlopbalusd = 0
       
               
        ' Op. from Account Master
        Sqlqry = " select * from bank_mas where bank_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount = 0 Then
          MsgBox " Bank Code not found in Account Master"
          Exit Sub
        Else
          rs.MoveFirst
          Opbaldhs = rs!Open_baldhs
          Opbalusd = rs!open_balUSD
        End If
        
       ' Sqlqry1 = "select Sum(tra_amount) from casl_mas where Cash_code='" & Mid(lstBankCodes, 1, 6) & "' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "#"
       ' Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
       ' If IsNull(rs1.Fields(0)) = False Then opsal = rs1.Fields(0)
        
      ' Bank Receipt Bank Code before From date
        Sqlqry1 = "select sum(tra_amount) from brpt_mas where tcurrency='DHS' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Opbrptdhs = rs1.Fields(0)
        
      ' Bank Receipt Account Code before From date
        Sqlqry1 = "select sum(tra_amount) from brpt_tra where tcurrency='DHS' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Opbrptadhs = rs1.Fields(0)
        
              
      ' Debit note Credit before From date
        Sqlqry1 = "select Sum(tra_amount) from debt_mas where tcurrency='DHS' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then opdbntcrdhs = rs1.Fields(0)
        
      ' Credit note Debit before From date
        Sqlqry1 = "select Sum(tra_amount) from Crdt_mas where tcurrency='DHS' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Opcrntdbdhs = rs1.Fields(0)
        
      ' Bank Payment before From date
        Sqlqry2 = "select sum(tra_amount) from bpmt_mas where tcurrency='DHS' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs2.Fields(0)) = False Then Opbrptadhs = rs2.Fields(0)
           
      ' Bank Payment before From date
        Sqlqry2 = "select sum(tra_amount) from bpmt_tra where tcurrency='DHS' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs2.Fields(0)) = False Then Opbpmtadhs = rs2.Fields(0)
                     
      ' Cash receipt (credit) before From date
        sqlqry20 = "select sum(tra_amount) from crpt_tra where tcurrency='DHS' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs13 = db.OpenRecordset(sqlqry20, dbOpenDynaset)
        If IsNull(rs13.Fields(0)) = False Then opcrptdhs = rs13.Fields(0)
        
      ' Cash Payment (Debit) before From date
        Sqlqry21 = "select sum(tra_amount) from cpmt_tra where tcurrency='DHS' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs14 = db.OpenRecordset(Sqlqry21, dbOpenDynaset)
        If IsNull(rs14.Fields(0)) = False Then opcpmtdhs = rs14.Fields(0)
        
      ' Pdc Receipts before From date
        Sqlqry3 = "select sum(tra_amount) from prpt_mas1 where tcurrency='DHS' and Cheque_Dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstBankCodes, 1, 6) & "' and not isnull(posting_dt)"
        Set rs3 = db.OpenRecordset(Sqlqry3, dbOpenDynaset)
        If IsNull(rs3.Fields(0)) = False Then Opprptdhs = rs3.Fields(0)
                  
       ' Pdc Payments before From Date
        Sqlqry4 = "select sum(tra_amount) from Ppmt_mas where tcurrency='DHS' and Cheque_dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstBankCodes, 1, 6) & "' and not isnull(posting_Dt)"
        Set rs4 = db.OpenRecordset(Sqlqry4, dbOpenDynaset)
        If IsNull(rs4.Fields(0)) = False Then Opppmtdhs = rs4.Fields(0)
                
       ' Journal Debit Amount before From Date
        Sqlqry14 = "select sum(tra_damount) from Jrnl_tra where tcurrency='DHS' and tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Val(Mid(lstBankCodes, 1, 6)) & "' and dc_code ='D'"
        Set rs9 = db.OpenRecordset(Sqlqry14, dbOpenDynaset)
        If IsNull(rs9.Fields(0)) = False Then Opjdbdhs = rs9.Fields(0)
         
       ' Journal Credit Amount before From Date
        Sqlqry15 = "select sum(tra_camount) from Jrnl_tra where tcurrency='DHS' and tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Val(Mid(lstBankCodes, 1, 6)) & "' and dc_code ='C'"
        Set rs10 = db.OpenRecordset(Sqlqry15, dbOpenDynaset)
        If IsNull(rs10.Fields(0)) = False Then Opjcrdhs = rs10.Fields(0)
        
        
        Ttlopbaldhs = Opbaldhs + opcpmtdhs + Opcrntdbdhs - opdbntcrdhs - opcrptdhs + Opbrptdhs - Opbrptadhs + Opprptdhs + Opjdbdhs - Opbpmtdhs + Opbpmtadhs - Opppmtdhs - Opjcrdhs
        
        Sqlqry5 = "Insert into Bankreport values(" & 0 & ",'','DHS','" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','Opening Balance'," & Trim(Ttlopbaldhs) & "," & 0 & ")"
        ws.BeginTrans
        db.Execute (Sqlqry5)
        ws.CommitTrans
        
        
        ' Bank Receipt Bank Code before From date
        Sqlqry1 = "select sum(tra_amount) from brpt_mas where tcurrency='USD' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Opbrptusd = rs1.Fields(0)
        
      ' Bank Receipt Account Code before From date
        Sqlqry1 = "select sum(tra_amount) from brpt_tra where tcurrency='USD' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Opbrptausd = rs1.Fields(0)
        
              
      ' Debit note Credit before From date
        Sqlqry1 = "select Sum(tra_amount) from debt_mas where tcurrency='USD' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then opdbntcrusd = rs1.Fields(0)
        
      ' Credit note Debit before From date
        
        
        Sqlqry1 = "select Sum(tra_amount) from Crdt_mas where tcurrency='USD' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then Opcrntdbusd = rs1.Fields(0)
        
      ' Bank Payment before From date
        Sqlqry2 = "select sum(tra_amount) from bpmt_mas where tcurrency='USD' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs2.Fields(0)) = False Then Opbrptausd = rs2.Fields(0)
           
      ' Bank Payment before From date
        Sqlqry2 = "select sum(tra_amount) from bpmt_tra where tcurrency='USD' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
        If IsNull(rs2.Fields(0)) = False Then Opbpmtausd = rs2.Fields(0)
                     
      ' Cash receipt (credit) before From date
        sqlqry20 = "select sum(tra_amount) from crpt_tra where tcurrency='USD' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs13 = db.OpenRecordset(sqlqry20, dbOpenDynaset)
        If IsNull(rs13.Fields(0)) = False Then opcrptusd = rs13.Fields(0)
        
      ' Cash Payment (Debit) before From date
        Sqlqry21 = "select sum(tra_amount) from cpmt_tra where tcurrency='USD' and tdate< #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs14 = db.OpenRecordset(Sqlqry21, dbOpenDynaset)
        If IsNull(rs14.Fields(0)) = False Then opcpmtusd = rs14.Fields(0)
        
      ' Pdc Receipts before From date
        Sqlqry3 = "select sum(tra_amount) from prpt_mas1 where tcurrency='USD' and Cheque_Dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstBankCodes, 1, 6) & "' and not isnull(posting_dt)"
        Set rs3 = db.OpenRecordset(Sqlqry3, dbOpenDynaset)
        If IsNull(rs3.Fields(0)) = False Then Opprptusd = rs3.Fields(0)
                  
       ' Pdc Payments before From Date
        Sqlqry4 = "select sum(tra_amount) from Ppmt_mas where tcurrency='USD' and Cheque_dt<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstBankCodes, 1, 6) & "' and not isnull(posting_Dt)"
        Set rs4 = db.OpenRecordset(Sqlqry4, dbOpenDynaset)
        If IsNull(rs4.Fields(0)) = False Then Opppmtusd = rs4.Fields(0)
                
       ' Journal Debit Amount before From Date
        Sqlqry14 = "select sum(tra_damount) from Jrnl_tra where tcurrency='USD' and tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Val(Mid(lstBankCodes, 1, 6)) & "' and dc_code ='D'"
        Set rs9 = db.OpenRecordset(Sqlqry14, dbOpenDynaset)
        If IsNull(rs9.Fields(0)) = False Then Opjdbusd = rs9.Fields(0)
         
       ' Journal Credit Amount before From Date
        Sqlqry15 = "select sum(tra_camount) from Jrnl_tra where tcurrency='USD' and tdate<#" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Val(Mid(lstBankCodes, 1, 6)) & "' and dc_code ='C'"
        Set rs10 = db.OpenRecordset(Sqlqry15, dbOpenDynaset)
        If IsNull(rs10.Fields(0)) = False Then Opjcrusd = rs10.Fields(0)
        
        
        Ttlopbalusd = Opbalusd + opcpmtusd + Opcrntdbusd - opdbntcrusd - opcrptusd + Opbrptusd - Opbrptausd + Opprptusd + Opjdbusd - Opbpmtusd + Opbpmtausd - Opppmtusd - Opjcrusd
        
        Sqlqry5 = "Insert into Bankreport values(" & 0 & ",'','USD','" & Format(txtdatefrom.TextWithMask, "dd/mm/yyyy") & "','Opening Balance'," & Trim(Ttlopbalusd) & "," & 0 & ")"
        ws.BeginTrans
        db.Execute (Sqlqry5)
        ws.CommitTrans
        
        
        'sqlqry20 = "select * from casl_mas where Cash_code='" & Mid(lstBankCodes, 1, 6) & "' and tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "#"
        'Set rs13 = db.OpenRecordset(sqlqry20, dbOpenDynaset)
        'If rs13.RecordCount <> 0 Then
        '  rs13.MoveFirst
        ' Do Until rs13.EOF
        '  Sqlqry21 = "Insert into bankreport values('" & rs13!VOUC_NO & "','" & rs13!vouc_type & "','" & Trim(rs13!tDate) & "','Cash Sales'," & Trim(rs13!namount) & "," & 0 & ")"
        '  ws.BeginTrans
        '  db.Execute (Sqlqry21)
        '  ws.CommitTrans
        '  rs13.MoveNext
        ' Loop
        'End If
        
        ' Bank Receipt bank code after From date and before to date
        Sqlqry6 = "select * from brpt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs5 = db.OpenRecordset(Sqlqry6, dbOpenDynaset)
        If rs5.RecordCount <> 0 Then
         rs5.MoveFirst
         Do Until rs5.EOF
          Sqlqry7 = "Insert into bankreport values('" & rs5!VOUC_NO & "','" & rs5!vouc_type & "','" & rs5!tcurrency & "','" & Trim(rs5!tDate) & "','" & Trim(rs5!Description) & "'," & Trim(rs5!tra_amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry7)
          ws.CommitTrans
          rs5.MoveNext
         Loop
        End If
        
          ' Bank Receipt  Account code after From date and before to date
        Sqlqry6 = "select * from brpt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs5 = db.OpenRecordset(Sqlqry6, dbOpenDynaset)
        If rs5.RecordCount <> 0 Then
         rs5.MoveFirst
         Do Until rs5.EOF
          Sqlqry7 = "Insert into bankreport values('" & rs5!VOUC_NO & "','" & rs5!vouc_type & "','" & rs5!tcurrency & "','" & Trim(rs5!tDate) & "','" & Trim(rs5!Description) & "'," & 0 & "," & Trim(rs5!tra_amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry7)
          ws.CommitTrans
          rs5.MoveNext
         Loop
        End If
        
        ' Debit Note Credit after From date and before to date
        Sqlqry6 = "select * from debt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs5 = db.OpenRecordset(Sqlqry6, dbOpenDynaset)
        If rs5.RecordCount <> 0 Then
         rs5.MoveFirst
         Do Until rs5.EOF
          Sqlqry7 = "Insert into bankreport values('" & rs5!VOUC_NO & "','" & rs5!vouc_type & "','" & rs5!tcurrency & "','" & Trim(rs5!tDate) & "','" & Trim(rs5!Description) & "'," & 0 & "," & Trim(rs5!tra_amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry7)
          ws.CommitTrans
          rs5.MoveNext
         Loop
        End If
        
          ' Credit Note Debit after From date and before to date
        Sqlqry6 = "select * from crdt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs5 = db.OpenRecordset(Sqlqry6, dbOpenDynaset)
        If rs5.RecordCount <> 0 Then
         rs5.MoveFirst
         Do Until rs5.EOF
          Sqlqry7 = "Insert into bankreport values('" & rs5!VOUC_NO & "','" & rs5!vouc_type & "','" & rs5!tcurrency & "','" & Trim(rs5!tDate) & "','" & Trim(rs5!Description) & "'," & Trim(rs5!tra_amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry7)
          ws.CommitTrans
          rs5.MoveNext
         Loop
        End If
        
        ' Bank Payment Bank Code after From date and before to date
        Sqlqry8 = "select * from bpmt_mas where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs6 = db.OpenRecordset(Sqlqry8, dbOpenDynaset)
        If rs6.RecordCount <> 0 Then
         rs6.MoveFirst
         Do Until rs6.EOF
          Sqlqry9 = "Insert into bankreport values('" & rs6!VOUC_NO & "','" & rs6!vouc_type & "','" & rs6!tcurrency & "','" & Trim(rs6!tDate) & "','" & Trim(rs6!Description) & "'," & 0 & "," & Trim(rs6!tra_amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry9)
          ws.CommitTrans
          rs6.MoveNext
         Loop
        End If
          
        ' Bank Payment Acct Code after From date and before to date
        Sqlqry8 = "select * from bpmt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs6 = db.OpenRecordset(Sqlqry8, dbOpenDynaset)
        If rs6.RecordCount <> 0 Then
         rs6.MoveFirst
         Do Until rs6.EOF
          Sqlqry9 = "Insert into bankreport values('" & rs6!VOUC_NO & "','" & rs6!vouc_type & "','" & rs6!tcurrency & "','" & Trim(rs6!tDate) & "','" & Trim(rs6!Description) & "'," & Trim(rs6!tra_amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry9)
          ws.CommitTrans
          rs6.MoveNext
         Loop
        End If
               
        
        ' Cash Payment (Debit) after From date and before to date
        Sqlqry22 = "select * from cpmt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs15 = db.OpenRecordset(Sqlqry22, dbOpenDynaset)
        If rs15.RecordCount <> 0 Then
         rs15.MoveFirst
         Do Until rs15.EOF
          Sqlqry23 = "Insert into bankreport values('" & rs15!VOUC_NO & "','" & rs15!vouc_type & "','" & rs15!tcurrency & "','" & Trim(rs15!tDate) & "','" & Trim(rs15!Description) & "'," & Trim(rs15!tra_amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (Sqlqry23)
          ws.CommitTrans
          rs15.MoveNext
         Loop
        End If
        
        ' Cash Receipt (credit) after From date and before to date
        Sqlqry24 = "select * from crpt_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstBankCodes, 1, 6) & "'"
        Set rs16 = db.OpenRecordset(Sqlqry24, dbOpenDynaset)
        If rs16.RecordCount <> 0 Then
         rs16.MoveFirst
         Do Until rs16.EOF
          Sqlqry25 = "Insert into bankreport values('" & rs16!VOUC_NO & "','" & rs16!vouc_type & "','" & rs16!tcurrency & "','" & Trim(rs16!tDate) & "','" & Trim(rs16!Description) & "'," & 0 & "," & Trim(rs16!tra_amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry25)
          ws.CommitTrans
          rs16.MoveNext
         Loop
        End If
        
        
        ' Pdc Receipts after From date and before to date
        Sqlqry10 = "select * from prpt_mas1 where Cheque_dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and Cheque_Dt<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstBankCodes, 1, 6) & "' and not isnull(Posting_Dt) "
        Set rs7 = db.OpenRecordset(Sqlqry10, dbOpenDynaset)
        If rs7.RecordCount <> 0 Then
         rs7.MoveFirst
         Do Until rs7.EOF
          SQLQRY11 = "Insert into bankreport values('" & rs7!VOUC_NO & "','" & rs7!vouc_type & "','" & rs7!tcurrency & "','" & Trim(rs7!Cheque_Dt) & "','" & Trim(rs7!Description) & "'," & Trim(rs7!tra_amount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (SQLQRY11)
          ws.CommitTrans
          rs7.MoveNext
         Loop
        End If
        
        ' Pdc Payments after From date and before to date
        SQLQRY12 = "select * from Ppmt_mas where Cheque_dt>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and bank_code='" & Mid(lstBankCodes, 1, 6) & "' and not Isnull(Posting_Dt) "
        Set rs8 = db.OpenRecordset(SQLQRY12, dbOpenDynaset)
        If rs8.RecordCount <> 0 Then
         rs8.MoveFirst
         Do Until rs8.EOF
          Sqlqry13 = "Insert into bankreport values('" & rs8!VOUC_NO & "','" & rs8!vouc_type & "','" & rs8!tcurrency & "','" & Trim(rs8!Cheque_Dt) & "','" & Trim(rs8!Description) & "'," & 0 & "," & Trim(rs8!tra_amount) & ")"
          ws.BeginTrans
          db.Execute (Sqlqry13)
          ws.CommitTrans
          rs8.MoveNext
         Loop
        End If
   
       ' Journal Debit after From date and before to date
        sqlqry16 = "select * from jrnl_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstBankCodes, 1, 6) & "' and dc_code='D' "
        Set rs11 = db.OpenRecordset(sqlqry16, dbOpenDynaset)
         If rs11.RecordCount <> 0 Then
          rs11.MoveFirst
          Do Until rs11.EOF
          sqlqry17 = "Insert into bankreport values('" & rs11!VOUC_NO & "','" & rs11!vouc_type & "','" & rs11!tcurrency & "','" & Trim(rs11!tDate) & "','" & Trim(rs11!Description) & "'," & Trim(rs11!tra_damount) & "," & 0 & ")"
          ws.BeginTrans
          db.Execute (sqlqry17)
          ws.CommitTrans
          rs11.MoveNext
         Loop
        End If
        
        ' Journal Credit after From date and before to date
        sqlqry18 = "select * from Jrnl_tra where tdate>= #" & DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) & "# and tdate<=#" & DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy")) & "# and Acct_code='" & Mid(lstBankCodes, 1, 6) & "' and dc_code='C'"
        Set rs12 = db.OpenRecordset(sqlqry18, dbOpenDynaset)
        If rs12.RecordCount <> 0 Then
         rs12.MoveFirst
        Do Until rs12.EOF
          sqlqry19 = "Insert into bankreport values('" & rs12!VOUC_NO & "','" & rs12!vouc_type & "','" & rs12!tcurrency & "','" & Trim(rs12!tDate) & "','" & Trim(rs12!Description) & "'," & 0 & "," & Trim(rs12!tra_Camount) & ")"
          ws.BeginTrans
          db.Execute (sqlqry19)
          ws.CommitTrans
          rs12.MoveNext
         Loop
        End If
   
    With CrystalReport1
     .DataFiles(0) = App.Path & "\misov.mdb"
     .ReportFileName = App.Path & "\BankReport.rpt"
    ' .ReportFileName = App.Path & "\bankreporttest.rpt"
     .Formulas(0) = "zzz='" & " From " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
     .Formulas(1) = "yyy='" & Mid(lstBankCodes, 8, 30) & "'"
     .WindowState = crptMaximized
     .Action = 1
    End With
   
   Else
        MsgBox "Improper Dates", vbDefaultButton1, "Invalid entry"
        Exit Sub
  End If
  
End Sub

Private Sub Form_Load()
 PopulateBankCodes
 txtdatefrom.TextWithMask = Format(Now, "DD/mm/yyyy")
 txtdateto.TextWithMask = Format(Now, "DD/mm/yyyy")
End Sub

Private Sub PopulateBankCodes()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select * from Bank_mas order by bank_code"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

lstBankCodes.Clear

 If rs.RecordCount = 0 Then
      MsgBox "No Records found in the Bank Master"
 Else
      rs.MoveFirst
   Do Until rs.EOF
      lstBankCodes.AddItem rs!bank_code & " : " & rs!BANK_NAME
      rs.MoveNext
   Loop
 End If

End Sub
Private Sub lstBankCodes_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdatefrom.SetFocus
End Sub
Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdateto.SetFocus
End Sub

Private Sub txtdatefrom_LostFocus()
 If IsDate(txtdatefrom.TextWithMask) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdDisplay.SetFocus
End Sub
Private Function ValidateData()
ValidateData = False

If IsDate(txtdatefrom.TextWithMask) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf lstBankCodes.SelCount = 0 Then
  MsgBox "Select Bank", vbInformation, "Invalid Entry"
  lstBankCodes.SetFocus
  SendKeys " {Home} + {end} "
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
 lstBankCodes.ListIndex = -1
 txtdatefrom.TextWithMask = Format(Now, "DD/mm/yyyy")
 txtdateto.TextWithMask = Format(Now, "DD/mm/YYYY")
End Sub

Private Sub txtdateto_LostFocus()
    If IsDate(txtdateto.TextWithMask) = False Then
       MsgBox "Invalid To Date", vbInformation, "Invalid Entry"
       txtdateto.SetFocus
       SendKeys " {Home} + {End} "
    End If
End Sub
