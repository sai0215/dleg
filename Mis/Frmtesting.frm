VERSION 5.00
Begin VB.Form Frmtesting 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1410
   ClientTop       =   1665
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "Frmtesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Getdata()

  Dim dbs As DAO.Database
  Dim rs As Recordset
  Dim fldemployees As Field
  Dim intcount As Integer
  Dim eagency As String
  Dim esubmedia As String
  Dim ecurrency As String
  Dim elyactual As Currency
  Dim ecybudget As Currency
  Dim ecyactual As Currency
  
  
  intcount = 1
  
  Set dbs = DBEngine(0).OpenDatabase(App.Path & "/misov.mdb")
  Set rs = dbs.OpenRecordset("submediaanalysis", dbOpenTable)
  
  
  
  With Worksheets("sheet1").Rows(9)
    .Font.Bold = True
    .Cell(1, 5).Value = "Agency"
    .Cell(1, 6).Value = "Sub Media"
    .Cell(1, 7).Value = "Currency"
    .Cell(1, 8).Value = "Last year Actual"
    .Cell(1, 9).Value = "Current year Budget"
    .Cell(1, 10).Value = "Current year Actual"
  Endwith
  
  
  Do Until rs.EOF
     Set fldemployees = rs.Fields(1)
      eagency = fldemployees.Value
     Set fldemployees = rs.Fields(2)
      esubmedia = fldemployees.Value
    Set fldemployees = rs.Fields(3)
     ecurrency = fldemployees.Value
    Set fldemployees = rs.Fields(4)
     elyactual = fldemployees.Value
    Set fldemployees = rs.Fields(5)
     ecybudget = fldemployees.Value
    Set fldemployees = rs.Fields(6)
     ecyactual = fldemployees.Value
    
    
    intcount = intcount + 1
    
    Call addtosheet(intcount, eagency, esubmedia, ecurrency, elyactual, ecybudget, ecyactual)
    
    rs.MoveNext
   Loop
   
   With Worksheets("Sheet").Columns("E:G")
     .AutoFit
   End With
  
     
End Sub

Public Function addtosheet(intcount As Integer, eagency As String, ecurrency As String, elyactual As Currency, ecybudget As Currency, ecyactual As Currency)
  With Worksheets("Sheet").Rows(intcount)
    .Cell(10, 5).Value = eagency
    .Cell(10, 6).Value = esubmedia
    .Cell(10, 7).Value = ecurrency
    .Cell(10, 8).Value = elyactual
    .Cell(10, 9).Value = ecybudget
    .Cell(10, 10).Value = ecyactual
  End With
  
End Function

Private Sub Command1_Click()
    
    
 '   FileCopy SmediaTemplate, App.Path & "\Results\smedia.xls"
   
    
    ''in the smedia table
    'Dim oNWindConn As New ADODB.Connection, oProdRS As New ADODB.Recordset
    
    
    ''Open the ADO connection to the Excel workbook
    Dim oConn As ADODB.Connection
    Set oConn = New ADODB.Connection
    oConn.Open " & App.Path & " \ results \ smedia.xls
    oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & App.Path & "\Results\smedia.xls;" & _
               "Extended Properties=""Excel 9.0;HDR=NO;"""
    
    'Create a new table (or worksheet in the workbook)
    oConn.Execute "create table smedia (Agency char(255), Submedia char(255), Currency char(255), lastyearactual int, CurrentYearBudget int, CurrentYearActual int)"

    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    
    
    
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from submediaanalysis"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynamic)
    If rs.RecordCount <> 0 Then
    
    
    
    'oNWindConn.Open "provider=microsoft.jet.oledb.4.0; data source=" & App.Path & "/misov.mdb"
    'oProdRS.Open "SELECT * from smediaanalysis,onwindconn"

  '  rs.Open "Select * from smedia", oConn, adOpenKeyset, adLockOptimistic
    
    Do While Not (rs.EOF)
        oRS.AddNew
        oRS.Fields(0) = rs.Fields("agency").Value
        oRS.Fields(1) = rs.Fields("submedia").Value
        oRS.Fields(2) = rs.Fields("tcurrency").Value
        oRS.Fields(3) = rs.Fields("lyearactual").Value
        oRS.Fields(4) = rs.Fields("cyearbudget").Value
        oRS.Fields(5) = rs.Fields("cyearactual").Value
        oRS.Update
        rs.MoveNext
    Loop
    End If
    'Close the recordset and connection to Northwind
    'rs.Close
    'oNWindConn.Close
    
    'Close the recordset and connection to the Excel workbook
    oRS.Close
    oConn.Close
       
    'Open the workbook to examine the results
    DoEvents
    ShellExecute Me.hwnd, "Open", App.Path & "\Results\smedia.xls", "", "C:\", sw_shownormal
        
End Sub

