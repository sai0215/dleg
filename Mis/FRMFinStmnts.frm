VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmFinStmnts 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Financial Statements"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Financial statements As On"
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
      Height          =   4455
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   7695
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Height          =   3495
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   6615
         Begin VB.CommandButton cmdBack 
            BackColor       =   &H00FFFF80&
            Caption         =   "<<Bac&k"
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
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2280
            Width           =   1575
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
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   2280
            Width           =   1575
         End
         Begin VB.CommandButton cmdDisplay 
            BackColor       =   &H00FFFF80&
            Caption         =   "&P and L A/C"
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
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00FFFF80&
            Caption         =   "&Balance Sheet"
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
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2280
            Width           =   1695
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   5760
            Top             =   720
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   262150
         End
         Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
            Height          =   495
            Left            =   2640
            TabIndex        =   0
            Top             =   840
            Width           =   1455
            _Version        =   65541
            _ExtentX        =   2566
            _ExtentY        =   873
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
            X1              =   0
            X2              =   6600
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1800
            TabIndex        =   7
            Top             =   960
            Width           =   600
         End
      End
   End
   Begin VB.Label lblWait 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Please Wait -- -- -- It's In Processing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   6600
      Width           =   7335
   End
End
Attribute VB_Name = "frmFinStmnts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim rs As Recordset
Dim rs1 As Recordset
Dim salrec As Currency
Dim admnexp As Currency
Dim depreciation As Currency
Dim plac As Currency
Dim purchases As Currency
Dim cbal As Currency
Dim bbal As Currency
Dim br As Currency
Dim bp As Currency
Dim deposits As Currency
Dim advances As Currency
Dim othcurasset As Currency
Dim othcurliability As Currency
Dim sdbtrs As Currency
Dim scrtrs As Currency
Dim fixassets As Currency
Dim curassets As Currency
Dim capital As Currency
Dim reserves As Currency
Dim opexpenses As Currency
Dim opincome As Currency
     
Private Sub cmdBack_Click()
 Unload Me
End Sub

Private Sub CmdPrint_Click()

lblWait.Visible = True
 lblWait.Caption = "Please Wait -- -- -- TB In Process "
 
   Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " Delete * from plac"
        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
          
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " Delete * from balancesheet"
        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
   If ValidateData = True Then
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    
      
 salrec = 0
 admnexp = 0
 depreciation = 0
 plac = 0
 purchases = 0
 fixassets = 0
 curassets = 0
 capital = 0
 reserves = 0
 opexpenses = 0
 opincome = 0
 cbal = 0
 bbal = 0
 br = 0
 bp = 0
 deposits = 0
 advances = 0
 othcurasset = 0
 othcurliability = 0
 sdbtrs = 0
 scrtrs = 0
      
    '  Sales Recovery
    
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 301000 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then salrec = -(rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 401000 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then purchases = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code>'" & 401000 & "' and acct_code<=' " & 417000 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then admnexp = (rs1.Fields(0))
        
        plac = salrec - purchases - admnexp
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 103001 & "' and acct_code='" & 103002 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then cbal = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code>='" & 103101 & "' and acct_code<='" & 103500 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then bbal = (rs1.Fields(0))
        
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 103501 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then br = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 103601 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then deposits = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 105000 & "' and acct_code='" & 105500 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then advances = (rs1.Fields(0))
        
        
      '  Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 202100 & "'"
       ' Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
       ' If IsNull(rs1.Fields(0)) = False Then bp = -(rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code>='" & 101200 & "' and acct_code<'" & 102000 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then fixassets = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 102000 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then sdbtrs = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code>='" & 105600 & "' and acct_code<'" & 107000 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then othcurasset = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 202000 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then scrtrs = -(rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 201001 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then capital = -(rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 204000 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then reserves = -(rs1.Fields(0))
                
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 200260 & "' and acct_code='" & 200270 & "' and acct_code='" & 201002 & "' and acct_code='" & 202100 & "' and acct_code>='" & 202200 & "' and acct_code<='" & 203999 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then othcurliability = -(rs1.Fields(0))
        
            ' sales recovery
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into plac values('','Sales Recovery'," & 0 & "," & salrec & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
                 
        ' Purchases recovery
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into plac values('Purchases',''," & purchases & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
                 
                 
        ' admn expenses
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into plac values('Administrative Expenses',''," & admnexp & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
                 
                     
        ' Profit and loss account
        If plac > 0 Then
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into plac values('Profit',''," & plac & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
       Else
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into plac values('','Loss'," & 0 & "," & plac & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
      End If
        
        
        ' Capital
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet values('Capital',''," & capital & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
        
        ' Reserves
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet values('Reserves',''," & reserves & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
        
        ' Fixassets
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Fixed Assets'," & 0 & "," & fixassets & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
       ' Sundry creditors
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('Sundry Creditors',''," & scrtrs & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
        
     ' Other income
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('Interest, Discount & Other income',''," & othcurliability & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
        
        ' Profit and loss account
        If plac > 0 Then
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet values('Profit',''," & plac & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
       Else
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet values('','Loss'," & 0 & "," & plac & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
      End If
        
     ' Sundry Debtors
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Sundry Debtors '," & 0 & "," & sdbtrs & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
     
     
     ' cash
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Cash on hand'," & 0 & "," & cbal & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
     ' bank
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Bank balance in DHS.'," & 0 & "," & bbal & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
     ' b/r
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Bills Receivable'," & 0 & "," & br & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
     
     ' Deposits
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Deposits'," & 0 & "," & deposits & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
     ' advances
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Advances'," & 0 & "," & advances & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
       
     
     
     ' Other Curassets
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Others and Prepaid Expenses '," & 0 & "," & othcurasset & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
     
     
    
     
     With CrystalReport1
       .DataFiles(0) = App.Path & "\misov.mdb"
       .ReportFileName = App.Path & "\balancesheet.rpt"
       .Formulas(0) = "xxx1='" & txtdatefrom.TextWithMask & "'"
       .WindowMaxButton = True
       .WindowState = crptMaximized
       .Action = 1
     End With
        
        
    End If
    
    End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub
Private Sub cmdClear_Click()
 textclear
End Sub
Private Sub cmdDisplay_Click()
 lblWait.Visible = True
 lblWait.Caption = "Please Wait -- -- -- TB In Process "
 
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " Delete * from plac"
        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
          
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " Delete * from balancesheet"
        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
        
   If ValidateData = True Then
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    
      
 salrec = 0
 admnexp = 0
 depreciation = 0
 plac = 0
 purchases = 0
 fixassets = 0
 curassets = 0
 capital = 0
 reserves = 0
 opexpenses = 0
 opincome = 0
 cbal = 0
 bbal = 0
 br = 0
 bp = 0
 deposits = 0
 advances = 0
 othcurasset = 0
 othcurliability = 0
 sdbtrs = 0
 scrtrs = 0
      
    '  Sales Recovery
    
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 301000 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then salrec = -(rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 401000 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then purchases = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code>'" & 401000 & "' and acct_code<=' " & 417000 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then admnexp = (rs1.Fields(0))
        
        plac = salrec - purchases - admnexp
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 103001 & "' and acct_code='" & 103002 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then cbal = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code>='" & 103101 & "' and acct_code<='" & 103500 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then bbal = (rs1.Fields(0))
        
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 103501 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then br = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 103601 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then deposits = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 105000 & "' and acct_code='" & 105500 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then advances = (rs1.Fields(0))
        
        
      '  Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 202100 & "'"
       ' Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
       ' If IsNull(rs1.Fields(0)) = False Then bp = -(rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code>='" & 101200 & "' and acct_code<'" & 102000 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then fixassets = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 102000 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then sdbtrs = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code>='" & 105600 & "' and acct_code<'" & 107000 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then othcurasset = (rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 202000 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then scrtrs = -(rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 201001 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then capital = -(rs1.Fields(0))
        
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 204000 & "'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then reserves = -(rs1.Fields(0))
                
        Sqlqry1 = " select sum(close_bal) from acct_mas where acct_code='" & 200260 & "' and acct_code='" & 200270 & "' and acct_code='" & 201002 & "' and acct_code='" & 202100 & "' and acct_code>='" & 202200 & "' and acct_code<='" & 203999 & "' "
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then othcurliability = -(rs1.Fields(0))
        
        ' sales recovery
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into plac values('','Sales Recovery'," & 0 & "," & salrec & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
                 
        ' Purchases recovery
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into plac values('Purchases',''," & purchases & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
                 
                 
        ' admn expenses
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into plac values('Administrative Expenses',''," & admnexp & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
                 
                     
        ' Profit and loss account
        If plac > 0 Then
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into plac values('Profit',''," & plac & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
       Else
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into plac values('','Loss'," & 0 & "," & plac & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
      End If
        
        
        ' Capital
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet values('Capital',''," & capital & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
        
        ' Reserves
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet values('Reserves',''," & reserves & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
        
        ' Fixassets
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Fixed Assets'," & 0 & "," & fixassets & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
       ' Sundry creditors
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('Sundry Creditors',''," & scrtrs & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
        
     ' Other income
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('Interest, Discount & Other income',''," & othcurliability & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
        
        ' Profit and loss account
        If plac > 0 Then
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet values('Profit',''," & plac & "," & 0 & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
       Else
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet values('','Loss'," & 0 & "," & plac & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
      End If
        
     ' Sundry Debtors
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Sundry Debtors '," & 0 & "," & sdbtrs & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
     
     
     ' cash
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Cash on hand'," & 0 & "," & cbal & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
     ' bank
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Bank balance in DHS.'," & 0 & "," & bbal & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
     ' b/r
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Bills Receivable'," & 0 & "," & br & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
     
     ' Deposits
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Deposits'," & 0 & "," & deposits & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
     ' advances
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Advances'," & 0 & "," & advances & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
       
     
     
     ' Other Curassets
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry = " Insert into balancesheet  values('','Others and Prepaid Expenses '," & 0 & "," & othcurasset & ")"
                 ws.BeginTrans
                 db.Execute (Sqlqry)
                 ws.CommitTrans
     
     
     
     
     With CrystalReport1
       .DataFiles(0) = App.Path & "\misov.mdb"
       .ReportFileName = App.Path & "\plac.rpt"
       .Formulas(0) = "xxx1='" & txtdatefrom.TextWithMask & "'"
       .WindowMaxButton = True
       .WindowState = crptMaximized
       .Action = 1
     End With
        
        
    End If
End Sub

Private Sub Form_Load()
    txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
    lblWait.Visible = False
    
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry1 = " delete * from plac "
                         
                 ws.BeginTrans
                 db.Execute (Sqlqry1)
                 ws.CommitTrans
                 
          Set ws = DBEngine.Workspaces(0)
          Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
          Sqlqry1 = " delete * from balancesheet "
                         
                 ws.BeginTrans
                 db.Execute (Sqlqry1)
                 ws.CommitTrans
                 
                     
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdDisplay.SetFocus
End Sub

Private Function ValidateData()
 ValidateData = False
If IsDate(txtdatefrom.TextWithMask) = False Then
   MsgBox "Invalid From Date", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
   ValidateData = True
End If
End Function

Private Sub textclear()
  txtdatefrom.TextWithMask = Format(Now(), "dd/mm/yyyy")
End Sub


Private Sub txtdatefrom_LostFocus()
If IsDate(txtdatefrom.TextWithMask) = False Then
      MsgBox "Invalid Date from ", vbInformation, "Invalid Entry"
      txtdatefrom.SetFocus
      SendKeys "{Home} + {End}"
    End If
End Sub
