VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1845
   ClientTop       =   1530
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\udtest.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      DragMode        =   1  'Automatic
      Exclusive       =   0   'False
      Height          =   615
      Left            =   1080
      OLEDropMode     =   1  'Manual
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   480
      Width           =   2700
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5741
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim ws As Workspace
    Dim db As Database
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim Sqlqry As String
    Dim Sqlqry1 As String
    
 
 
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
      
   Sqlqry1 = " select * from acct_mas  "
   Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
   If rs.RecordCount <> 0 Then
      MsgBox " Account Code already exists"
      Exit Sub
   Else
   
      MsgBox "Record is inserted", vbDefaultButton3, "Status"
    Exit Sub
   End If
 
End Sub

Private Sub Flexitems()
    With DataGrid1
        .DataSource = " d:\mis\misov.mdb"
        .ClearFields
        .Columns = 6
        .Col = 0
        .Text = "Code"
        .C
        .ColWidth(0) = 1300
        .ColWidth(1) = 6300
        .ColWidth(2) = 700
        .ColWidth(3) = 1025
        .ColWidth(4) = 750
        .ColWidth(5) = 850
        .Col = 1
        .CellBackColor = RGB(180, 170, 160)
        .Text = "Description"
        .Col = 2
        .CellBackColor = RGB(180, 170, 160)
        .Text = "Unit"
        .Col = 3
        .CellBackColor = RGB(180, 170, 160)
        .Text = "Quantity"
        .Col = 4
        .CellBackColor = RGB(180, 170, 160)
        .Text = "Rate"
        .Col = 5
        .CellBackColor = RGB(180, 170, 160)
        .Text = "Amount"
        .Row = 0
        .Col = 1
      End With
End Sub
