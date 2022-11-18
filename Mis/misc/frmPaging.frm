VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPaging 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Paging"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9840
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Close DB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CmdBack 
      BackColor       =   &H0080C0FF&
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
      Height          =   615
      Left            =   6960
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CMDVis 
      BackColor       =   &H0080C0FF&
      Caption         =   "&View DB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7440
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=udacc"
      OLEDBString     =   "DSN=udacc"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select sub_media,issue_no,tdate,page,Comments,Product,Agency,Mat_code,Tra_amount from dumBo_traPaging "
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPaging.frx":0000
      Height          =   4935
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8705
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "sub_media"
         Caption         =   "Sub Media"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "ISSUE_NO"
         Caption         =   "Issue #"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "TDATE"
         Caption         =   "Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "PAGE"
         Caption         =   "Page"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "comments"
         Caption         =   "Comments"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PRODUCT"
         Caption         =   "Product"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "AGENCY"
         Caption         =   "Agency"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "mat_code"
         Caption         =   "Material"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "TRA_AMOUNT"
         Caption         =   "Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
            ColumnWidth     =   1395.213
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   10335
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Preview without Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9000
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox CboSYear 
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
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox CboSMonth 
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
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox CboSmedia 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   4575
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Preview with Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9000
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox CboIssue 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   9360
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin VB.Label Label74 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Year"
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
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label75 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Month"
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
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label76 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Media"
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
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label LblIss 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Issue #"
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
         Height          =   255
         Left            =   6240
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblviewMedia 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2880
         TabIndex        =   7
         Top             =   2520
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label LblviewSubmedia 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   3240
         TabIndex        =   6
         Top             =   2520
         Visible         =   0   'False
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmPaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xx
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim sqlqry3 As String
Dim stat As Integer
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Public Function issueno()
 If xx <> "" Then
   Sqlqry1 = "SELECT sub_media,ISSUE_NO,TDATE,PAGE,comments,PRODUCT,AGENCY,mat_code,TRA_AMOUNT  from bo_tramag  where issue_no='" & xx & "'order by issue_no,sub_media,tdate"
   Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
 Else
     Sqlqry1 = "SELECT sub_media,ISSUE_NO,TDATE,PAGE,comments,PRODUCT,AGENCY,mat_code,TRA_AMOUNT  from bo_tramag order by issue_no,sub_media,tdate"
     Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
 End If
End Function
Private Sub CboIssue_LostFocus()
X = Trim(CboIssue.Text)
End Sub
Private Sub CboSmedia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboIssue.SetFocus
End Sub
Private Sub CboSmedia_LostFocus()
If CboSYear.Text = "" Then
 MsgBox " Invalid year"
 CboSYear.SetFocus
 Exit Sub
End If
If CboSMonth.Text = "" Then
 MsgBox " Invalid month"
 CboSMonth.SetFocus
 Exit Sub
End If
 
populateissuenos
End Sub
Private Sub CboSMonth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboSmedia.SetFocus
End Sub
Private Sub CboSMonth_LostFocus()
    If CboSYear.Text <> "" And CboSMonth.Text <> "" And CboSmedia.Text <> "" Then
    populateissuenos
    End If
End Sub

Private Sub CboSYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboSMonth.SetFocus
End Sub

Private Sub CboSYear_LostFocus()
If CboSYear.Text <> "" And CboSMonth.Text <> "" And CboSmedia.Text <> "" Then
populateissuenos
End If
End Sub

Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub CmdClose_Click()
DataGrid1.Visible = False
End Sub


Private Sub cmdUpdate_Click()

Dim l, o, p As String
Dim n, m As String
Dim Q
Dim pag
      
  If stat <> 1 Then Exit Sub
  
   n = Trim(lblviewMedia.Caption)
   m = Trim(LblviewSubmedia.Caption)
   
   Adodc1.Refresh
   
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry2 = "Delete * from Bo_TRAmag where sub_media='" & Trim(CboSmedia) & "' and year='" & Val(CboSYear) & "' and month='" & Trim(CboSMonth) & "' and issue_no='" & Trim(CboIssue) & "'"
   ws.BeginTrans
   db.Execute (Sqlqry2)
   ws.CommitTrans
          
     
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry1 = "Select * from dumbo_trapaging order by serial_no"
   Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
   If rs.RecordCount <> 0 Then
        rs.MoveFirst
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         
           Sqlqry2 = " Insert into bo_tramag values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(CboSMonth.ListIndex) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!NET_Amount) & ",'" _
                                     & Trim(rs!agcom) & "','" _
                                     & Trim(rs!adper) & "'," _
                                     & Trim(rs!addisc) & "," _
                                     & Trim(rs!surcharge) & ")"
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
         
         
          rs.MoveNext
         Loop
       End If
       
        
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry1 = "Select * from dumbo_trapagingext order by serial_no"
   Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
   If rs.RecordCount <> 0 Then
        rs.MoveFirst
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         
           Sqlqry2 = " Insert into bo_tramag values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(CboSMonth.ListIndex) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!NET_Amount) & ",'" _
                                     & Trim(rs!agcom) & "','" _
                                     & Trim(rs!adper) & "'," _
                                     & Trim(rs!addisc) & "," _
                                     & Trim(rs!surcharge) & ")"
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
         
         
          rs.MoveNext
         Loop
       End If
        
        
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry = "Delete * from dumBo_TRAPaging"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry = "Delete * from dumBo_TRAPagingext"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        
       Adodc1.Refresh
        
End Sub

Private Sub cmdView_Click()
Dim l, o, p As String
Dim n, m As String
Dim Q

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry = "Delete * from dumBo_TRAmag"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Delete * from dumBo_TRAPaging"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
        
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = "Delete * from dumBo_TRAPagingext"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
   Adodc1.Refresh

   n = Trim(lblviewMedia.Caption)
   m = Trim(LblviewSubmedia.Caption)
   stat = 1
   Sqlqry1 = "Select * from Bo_TRAmag where sub_media='" & Trim(CboSmedia) & "' and year='" & Val(CboSYear) & "' and month='" & Trim(CboSMonth) & "' and issue_no='" & Trim(CboIssue) & "'"
   Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount = 0 Then
         MsgBox " No Transactions"
         Exit Sub
      Else
         rs.MoveFirst
         
         Do Until rs.EOF
         
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           Sqlqry2 = " Insert into dumbo_tramag values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(CboSMonth.ListIndex) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!Amount) & ",'" _
                                     & Trim(rs!agcom) & "','" _
                                     & Trim(rs!adper) & "'," _
                                     & Trim(rs!addisc) & "," _
                                     & Trim(rs!surcharge) & ")"
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
            
            
           sqlqry3 = " Insert into dumbo_traPaging values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(CboSMonth.ListIndex) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!Amount) & ",'" _
                                     & Trim(rs!agcom) & "','" _
                                     & Trim(rs!adper) & "'," _
                                     & Trim(rs!addisc) & "," _
                                     & Trim(rs!surcharge) & ",'N')"
            ws.BeginTrans
            db.Execute (sqlqry3)
            ws.CommitTrans
           rs.MoveNext
         Loop
       End If
       
        Sqlqry = "Select * from bO_MAS WHERE CANCELL='Y'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
          rs.MoveFirst
             Do Until rs.EOF
                  Sqlqry1 = "DELETE * FROM DUMBO_TRAMAG WHERE SERIAL_NO='" & rs!serial_no & "' "
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
                   
                   Sqlqry2 = "update DUMBO_TRAPaging set Position='Y' WHERE SERIAL_NO='" & rs!serial_no & "' "
                   ws.BeginTrans
                   db.Execute (Sqlqry2)
                   ws.CommitTrans
               
              
               rs.MoveNext
              Loop
        End If
    
    
   Sqlqry1 = "Select * from dumbo_traPaging where Position='Y'"
   Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount <> 0 Then
         
      
         rs.MoveFirst
         
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           Sqlqry2 = " Insert into dumbo_traPagingext values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(CboSMonth.ListIndex) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!NET_Amount) & ",'" _
                                     & Trim(rs!agcom) & "','" _
                                     & Trim(rs!adper) & "'," _
                                     & Trim(rs!addisc) & "," _
                                     & Trim(rs!surcharge) & ",'" & rs!Position & "')"
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
            
         rs.MoveNext
       Loop
     End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Delete * from dumBo_trapaging where position='Y'"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
    Adodc1.Refresh
    Adodc1.RecordSource = "select sub_media,issue_no,tdate,page,Comments,Product,Agency,Mat_code,Tra_amount from dumBo_traPaging "
   
    X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
    If X = vbYes Then
     With CrystalReport1
      .DataFiles(0) = App.Path & "\misov.mdb"
      .ReportFileName = App.Path & "\fpmagpage.rpt"
      .Formulas(0) = "yyy='" & Trim(CboSmedia) & "'"
      .WindowState = crptMaximized
      .Action = 1
    End With
    End If
   
End Sub
Private Sub populateMedia()
    
    CboSmedia.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Media where media_type='Magazine' Order by Media_Type"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        rs.MoveFirst
            Do Until rs.EOF
              CboSmedia.AddItem Trim(rs!sub_Media)
              rs.MoveNext
            Loop
    End If
    
End Sub
Private Sub Command1_Click()
Dim l, o, p As String
Dim n, m As String

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry = "Delete * from dumBo_TRAmag"
        ws.BeginTrans
        db.Execute (Sqlqry)
        ws.CommitTrans
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = "Delete * from dumBo_TRAPaging"
     ws.BeginTrans
     db.Execute (Sqlqry1)
     ws.CommitTrans
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry1 = "Delete * from dumBo_TRAPagingext"
     ws.BeginTrans
     db.Execute (Sqlqry1)
     ws.CommitTrans
     
   n = Trim(lblviewMedia.Caption)
   m = Trim(LblviewSubmedia.Caption)
   stat = 1
   Sqlqry1 = "Select * from Bo_TRAmag where sub_media='" & Trim(CboSmedia) & "' and year='" & Val(CboSYear) & "' and month='" & Trim(CboSMonth) & "' and issue_no='" & Trim(CboIssue) & "'"
   Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount = 0 Then
         MsgBox " Transactions are not recorded"
         Exit Sub
      Else
         rs.MoveFirst
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           Sqlqry2 = " Insert into dumbo_tramag values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(CboSMonth.ListIndex) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!Amount) & ",'" _
                                     & Trim(rs!agcom) & "','" _
                                     & Trim(rs!adper) & "'," _
                                     & Trim(rs!addisc) & "," _
                                     & Trim(rs!surcharge) & ")"
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
            
            
            
           sqlqry3 = " Insert into dumbo_traPaging values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(CboSMonth.ListIndex) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!Amount) & ",'" _
                                     & Trim(rs!agcom) & "','" _
                                     & Trim(rs!adper) & "'," _
                                     & Trim(rs!addisc) & "," _
                                     & Trim(rs!surcharge) & ",'N')"
            ws.BeginTrans
            db.Execute (sqlqry3)
            ws.CommitTrans
            
          rs.MoveNext
         Loop
       End If
       
        Sqlqry = "Select * from BO_MAS WHERE CANCELL='Y'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
          rs.MoveFirst
             Do Until rs.EOF
                  Sqlqry1 = "DELETE * FROM DUMBO_TRAMAG WHERE SERIAL_NO='" & rs!serial_no & "' "
                   ws.BeginTrans
                   db.Execute (Sqlqry1)
                   ws.CommitTrans
                   
                   
                   Sqlqry2 = "Update DUMBO_TRAPaging set Position='Y' WHERE SERIAL_NO='" & rs!serial_no & "' "
                   ws.BeginTrans
                   db.Execute (Sqlqry2)
                   ws.CommitTrans
              
               rs.MoveNext
              Loop
        End If
    
    
   Sqlqry1 = "Select * from dumbo_traPaging where Position='Y'"
   Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
      If rs.RecordCount = 0 Then
         MsgBox " Transactions are not recorded"
         Exit Sub
      Else
         rs.MoveFirst
         
         Do Until rs.EOF
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
           Sqlqry2 = " Insert into dumbo_traPagingext values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "'," & Val(CboSMonth.ListIndex) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!issue_no) & "','" _
                                     & Trim(rs!tDate) & "','" _
                                     & Trim(rs!Page) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!Comments)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Space) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!Amount) & ",'" _
                                     & Trim(rs!agcom) & "','" _
                                     & Trim(rs!adper) & "'," _
                                     & Trim(rs!addisc) & "," _
                                     & Trim(rs!surcharge) & ",'" & rs!Position & "')"
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
            
         rs.MoveNext
       Loop
     End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Delete * from dumBo_trapaging where position='Y'"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
    Adodc1.Refresh
    Adodc1.RecordSource = "select sub_media,issue_no,tdate,page,Comments,Product,Agency,Mat_code,Tra_amount from dumBo_traPaging "
       
   
    X = MsgBox("Do You Want to Print.", vbInformation + vbYesNo, "Print Confirm")
    If X = vbYes Then
     With CrystalReport1
      .DataFiles(0) = App.Path & "\misov.mdb"
      .ReportFileName = App.Path & "\fpmagpagewoamt.rpt"
      .Formulas(0) = "yyy='" & Trim(CboSmedia) & "'"
      .WindowState = crptMaximized
      .Action = 1
     End With
    End If
End Sub

Private Sub CMDVis_Click()
 DataGrid1.Visible = True
End Sub
Private Sub Form_Load()
stat = 0
Dim X As String
CboSMonth.AddItem "January"
CboSMonth.AddItem "February"
CboSMonth.AddItem "March"
CboSMonth.AddItem "April"
CboSMonth.AddItem "May"
CboSMonth.AddItem "June"
CboSMonth.AddItem "July"
CboSMonth.AddItem "August"
CboSMonth.AddItem "September"
CboSMonth.AddItem "October"
CboSMonth.AddItem "November"
CboSMonth.AddItem "December"
xx = ""

i = 2000
DT = 28
For i = 2000 To 2100
 CboSYear.AddItem i
Next
X = 0

 CboSYear.Text = Year(Now())
 
 X = Month(Now())
 
 
If X = 1 Then
   
   CboSMonth.ListIndex = 0
   DT = 31
ElseIf X = 2 Then
   
   CboSMonth.ListIndex = 1
   DT = 28
ElseIf X = 3 Then
  
   CboSMonth.ListIndex = 2
   DT = 31
ElseIf X = 4 Then
  
   CboSMonth.ListIndex = 3
   DT = 30
ElseIf X = 5 Then
   
   CboSMonth.ListIndex = 4
   DT = 31
ElseIf X = 6 Then
 
   CboSMonth.ListIndex = 5
   DT = 30
ElseIf X = 7 Then
  
   CboSMonth.ListIndex = 6
   DT = 31
ElseIf X = 8 Then
 
   CboSMonth.ListIndex = 7
   DT = 31
ElseIf X = 9 Then
  
   CboSMonth.ListIndex = 8
   DT = 30
ElseIf X = 10 Then
 
   CboSMonth.ListIndex = 9
   DT = 31
ElseIf X = 11 Then
  
   CboSMonth.ListIndex = 10
   DT = 30
Else
   
   CboSMonth.ListIndex = 11
   DT = 31
End If

    
populateMedia

Adodc1.RecordSource = "select sub_media,issue_no,tdate,page,Comments,Product,Agency,Mat_code,Tra_amount from dumBo_traPaging "



End Sub
Private Sub populateissuenos()
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = "Select distinct(issue_no) from bo_tramag where sub_media='" & Trim(CboSmedia) & "' and year='" & Val(CboSYear) & "' and month='" & Trim(CboSMonth) & "' "
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            If rs.RecordCount = 0 Then
                 Exit Sub
            Else
                CboIssue.Clear
                rs.MoveFirst
                    Do Until rs.EOF
                      CboIssue.AddItem Trim(rs!issue_no)
                      rs.MoveNext
                    Loop
            End If

End Sub
    
