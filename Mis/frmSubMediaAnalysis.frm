VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmSubMediaAnalysis 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Last year Vs Current Year"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
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
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   10575
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Height          =   3495
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   9735
         Begin VB.ComboBox CboYear 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   6480
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1200
            Width           =   1815
         End
         Begin VB.ComboBox CboCurrency 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1200
            Width           =   1815
         End
         Begin VB.ComboBox CboSubMedia 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   480
            Width           =   7335
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
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   2280
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
            Height          =   855
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CommandButton cmdDisplay 
            BackColor       =   &H00FFFF80&
            Caption         =   "P&review"
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
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   2280
            Width           =   1095
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   6000
            Top             =   2760
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   262150
         End
         Begin VB.Label lblYear 
            BackColor       =   &H00FFFFC0&
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
            Height          =   375
            Left            =   5640
            TabIndex        =   10
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Currency"
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
            Left            =   480
            TabIndex        =   7
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Sub Media"
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
            Left            =   480
            TabIndex        =   6
            Top             =   600
            Width           =   1575
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   9720
            Y1              =   2040
            Y2              =   2040
         End
      End
   End
End
Attribute VB_Name = "frmSubMediaAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database

Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim Sqlqry3 As String

Dim curyearact As Currency
Dim lastyearact As Currency
Dim curyearbud  As Currency

Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                   (ByVal hwnd As Long, ByVal lpszOp As String, _
                    ByVal lpszFile As String, ByVal lpszParams As String, _
                    ByVal LpszDir As String, ByVal FsShowCmd As Long) _
                    As Long
Private Const SW_NORMAL = 1
Private Const sw_shownormal = 1

Private sNwind As String            'Path to sample Access database
Private sOrdersTemplate As String   'Path to Orders Workbook "Template"
Private sEmpDataTemplate As String  'Path to Employee Data Workbook "Template"
Private SmediaTemplate As String 'Path to Products Workbook "Template"
Private sChartTemplate As String    'Path to Workbook "Template" containing a chart
Private sSourceData As String


Private Sub populateMedia()

 cbosubmedia.Clear
 
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Media Order by Sub_Media"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        cbosubmedia.AddItem "Cinema"
      
       rs.MoveFirst
           Do Until rs.EOF
              If rs!Media_Type = "Cinema" Then
               cbosubmedia.AddItem "Cinema" & " : " & Trim(rs!sub_media)
                Else
               cbosubmedia.AddItem Trim(rs!sub_media)
              End If
               rs.MoveNext
           Loop
    End If
End Sub
  
Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdDisplay.SetFocus
End Sub

Private Sub cmdBack_Click()
 Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub

Private Sub cmdClear_Click()
 textclear
End Sub
 Private Sub PopulateDhs()
      If cbosubmedia = "Magazine" Then
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = "Select * from agndtls order by agentname"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            If rs.RecordCount = 0 Then
               MsgBox " No Records Found In The Agency Register"
               Exit Sub
            Else
              rs.MoveFirst
                Do Until rs.EOF
                      curyearact = 0
                      lastyearact = 0
                      curyearbud = 0
                    
                      '  Current Year Actual
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearact = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearact = curyearact + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Last Year Actual
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then lastyearact = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then lastyearact = lastyearact + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Current Year Budget
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = curyearbud + Val(rs1.Fields(0) * convertion)
                         
                        Sqlqry1 = " Insert into submediaanalysis values('" & findfirstfixup(Trim(rs!agentname)) & "','" & Trim(cbosubmedia) & "','" & _
                                  Trim(cboCurrency) & "'," & _
                                  lastyearact & "," & _
                                  curyearbud & "," & _
                                  curyearact & ")"
                      
                        ws.BeginTrans
                        db.Execute (Sqlqry1)
                        ws.CommitTrans
                    
                  
                  
                        lastyearact = 0
                        curyearbud = 0
                        curyearact = 0
                  rs.MoveNext
                  Loop
              End If
              
         ElseIf cbosubmedia = "Television" Then
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = "Select * from agndtls order by agentname"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            If rs.RecordCount = 0 Then
               MsgBox " No Records Found In The Agency Register"
               Exit Sub
            Else
              rs.MoveFirst
                Do Until rs.EOF
                      curyearact = 0
                      lastyearact = 0
                      curyearbud = 0
                    
                      '  Current Year Actual
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearact = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearact = curyearact + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Last Year Actual
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then lastyearact = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then lastyearact = lastyearact + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Current Year Budget
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = curyearbud + Val(rs1.Fields(0) * convertion)
                         
                        Sqlqry1 = " Insert into submediaanalysis values('" & findfirstfixup(Trim(rs!agentname)) & "','" & Trim(cbosubmedia) & "','" & _
                                  Trim(cboCurrency) & "'," & _
                                  lastyearact & "," & _
                                  curyearbud & "," & _
                                  curyearact & ")"
                      
                        ws.BeginTrans
                        db.Execute (Sqlqry1)
                        ws.CommitTrans
                    
                  
                  
                        lastyearact = 0
                        curyearbud = 0
                        curyearact = 0
                  rs.MoveNext
                  Loop
              End If
              
         ' online
         ElseIf cbosubmedia = "Online" Then
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = "Select * from agndtls order by agentname"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            If rs.RecordCount = 0 Then
               MsgBox " No Records Found In The Agency Register"
               Exit Sub
            Else
              rs.MoveFirst
                Do Until rs.EOF
                      curyearact = 0
                      lastyearact = 0
                      curyearbud = 0
                    
                      '  Current Year Actual
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearact = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearact = curyearact + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Last Year Actual
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then lastyearact = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then lastyearact = lastyearact + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Current Year Budget
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = curyearbud + Val(rs1.Fields(0) * convertion)
                         
                        Sqlqry1 = " Insert into submediaanalysis values('" & findfirstfixup(Trim(rs!agentname)) & "','" & Trim(cbosubmedia) & "','" & _
                                  Trim(cboCurrency) & "'," & _
                                  lastyearact & "," & _
                                  curyearbud & "," & _
                                  curyearact & ")"
                      
                        ws.BeginTrans
                        db.Execute (Sqlqry1)
                        ws.CommitTrans
                    
                  
                  
                        lastyearact = 0
                        curyearbud = 0
                        curyearact = 0
                  rs.MoveNext
                  Loop
              End If
          ' Cinema
          
          ElseIf cbosubmedia = "Cinema" Then
                  Set ws = DBEngine.Workspaces(0)
                  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                  Sqlqry = "Select * from agndtls order by agentname"
                  Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                   If rs.RecordCount = 0 Then
                     MsgBox " No Records Found In The Agency Register"
                     Exit Sub
                   Else
                    rs.MoveFirst
                     Do Until rs.EOF
                      curyearact = 0
                      lastyearact = 0
                      curyearbud = 0
                      
                 
                     
                      '  Current Year Actual
                         Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) & "' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearact = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) & "' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearact = curyearact + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Last Year Actual
                         Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) - 1 & "' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then lastyearact = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) - 1 & "' and media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then lastyearact = lastyearact + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Current Year Budget
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = curyearbud + Val(rs1.Fields(0) * convertion)
                         
                          Sqlqry1 = " Insert into submediaanalysis values('" & findfirstfixup(Trim(rs!agentname)) & "','" & Trim(cbosubmedia) & "','" & _
                                    Trim(cboCurrency) & "'," & _
                                    lastyearact & "," & _
                                    curyearbud & "," & _
                                    curyearact & ")"
                        
                            ws.BeginTrans
                            db.Execute (Sqlqry1)
                            ws.CommitTrans
                        
                        
                        
                          lastyearact = 0
                          curyearbud = 0
                          curyearact = 0
                      rs.MoveNext
                      Loop
                     End If
                     
          ElseIf Mid(cbosubmedia, 1, 3) = "Cin" Then
                  Set ws = DBEngine.Workspaces(0)
                  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                  Sqlqry = "Select * from agndtls order by agentname"
                  Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                   If rs.RecordCount = 0 Then
                     MsgBox " No Records Found In The Agency Register"
                     Exit Sub
                   Else
                    rs.MoveFirst
                     Do Until rs.EOF
                      curyearact = 0
                      lastyearact = 0
                      curyearbud = 0
                      
                 
                     
                      '  Current Year Actual
                         Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) & "' and Sub_media='" & Trim(Mid(cbosubmedia.Text, 10, 50)) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearact = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) & "' and Sub_media='" & Trim(Mid(cbosubmedia.Text, 10, 50)) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearact = curyearact + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Last Year Actual
                         Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) - 1 & "' and Sub_media='" & Trim(Mid(cbosubmedia.Text, 10, 50)) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then lastyearact = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) - 1 & "' and Sub_media='" & Trim(Mid(cbosubmedia.Text, 10, 50)) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then lastyearact = lastyearact + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Current Year Budget
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = curyearbud + Val(rs1.Fields(0) * convertion)
                         
                        Sqlqry1 = " Insert into submediaanalysis values('" & findfirstfixup(Trim(rs!agentname)) & "','" & Trim(cbosubmedia) & "','" & _
                                  Trim(cboCurrency) & "'," & _
                                  lastyearact & "," & _
                                  curyearbud & "," & _
                                  curyearact & ")"
                      
                            ws.BeginTrans
                            db.Execute (Sqlqry1)
                            ws.CommitTrans
                        
                        
                        
                          lastyearact = 0
                          curyearbud = 0
                          curyearact = 0
                      rs.MoveNext
                      Loop
                     End If
           
          Else
          
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
            Sqlqry = "Select * from agndtls order by agentname"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            If rs.RecordCount = 0 Then
               MsgBox " No Records Found In The Agency Register"
               Exit Sub
            Else
              rs.MoveFirst
                Do Until rs.EOF
                      curyearact = 0
                      lastyearact = 0
                      curyearbud = 0
                    
                      '  Current Year Actual
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and sub_media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearact = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and sub_media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearact = curyearact + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Last Year Actual
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and sub_media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then lastyearact = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and sub_media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then lastyearact = lastyearact + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Current Year Budget
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = curyearbud + Val(rs1.Fields(0) * convertion)
                         
                        Sqlqry1 = " Insert into submediaanalysis values('" & findfirstfixup(Trim(rs!agentname)) & "','" & Trim(cbosubmedia) & "','" & _
                                  Trim(cboCurrency) & "'," & _
                                  lastyearact & "," & _
                                  curyearbud & "," & _
                                  curyearact & ")"
                      
                        ws.BeginTrans
                        db.Execute (Sqlqry1)
                        ws.CommitTrans
                    
                  
                  
                        lastyearact = 0
                        curyearbud = 0
                        curyearact = 0
                  rs.MoveNext
                  Loop
              End If
                
        End If
 End Sub
 
 Private Sub PopulateUsd()
 
     If cbosubmedia = "Magazine" Then
     
              Set ws = DBEngine.Workspaces(0)
              Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
              Sqlqry = "Select * from agndtls order by agentname"
              Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                   
                   If rs.RecordCount <> 0 Then
                      rs.MoveFirst
                          Do Until rs.EOF
                             curyearact = 0
                             lastyearact = 0
                             curyearbud = 0
                              
                            '  Current Year Actual
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearact = rs1.Fields(0)
                               
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearact = curyearact + Val(rs1.Fields(0) / convertion)
                               
                               
                            '  Last Year Actual
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then lastyearact = rs1.Fields(0)
                               
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then lastyearact = lastyearact + Val(rs1.Fields(0) / convertion)
                               
                               
                            '  Current Year Budget
                               Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='USD'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearbud = rs1.Fields(0)
                               
                               Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='DHS'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearbud = curyearbud + Val(rs1.Fields(0) / convertion)
                               
                                   Sqlqry1 = " Insert into submediaanalysis values('" & findfirstfixup(Trim(rs!agentname)) & "','" & Trim(cbosubmedia) & "','" & _
                                            Trim(cboCurrency) & "'," & _
                                            lastyearact & "," & _
                                            curyearbud & "," & _
                                            curyearact & ")"
                                  
                                    ws.BeginTrans
                                    db.Execute (Sqlqry1)
                                    ws.CommitTrans
                          
                        
                        
                              lastyearact = 0
                              curyearbud = 0
                              curyearact = 0
                       rs.MoveNext
                      Loop
                  End If
                  
         
             ElseIf cbosubmedia = "Television" Then
     
                Set ws = DBEngine.Workspaces(0)
                Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                Sqlqry = "Select * from agndtls order by agentname"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     
                   If rs.RecordCount <> 0 Then
                      rs.MoveFirst
                          Do Until rs.EOF
                             curyearact = 0
                             lastyearact = 0
                             curyearbud = 0
                              
                            '  Current Year Actual
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearact = rs1.Fields(0)
                               
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearact = curyearact + Val(rs1.Fields(0) / convertion)
                               
                               
                            '  Last Year Actual
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then lastyearact = rs1.Fields(0)
                               
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then lastyearact = lastyearact + Val(rs1.Fields(0) / convertion)
                               
                               
                            '  Current Year Budget
                               Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='USD'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearbud = rs1.Fields(0)
                               
                               Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='DHS'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearbud = curyearbud + Val(rs1.Fields(0) / convertion)
                               
                                   Sqlqry1 = " Insert into submediaanalysis values('" & findfirstfixup(Trim(rs!agentname)) & "','" & Trim(cbosubmedia) & "','" & _
                                            Trim(cboCurrency) & "'," & _
                                            lastyearact & "," & _
                                            curyearbud & "," & _
                                            curyearact & ")"
                                  
                                    ws.BeginTrans
                                    db.Execute (Sqlqry1)
                                    ws.CommitTrans
                          
                        
                        
                              lastyearact = 0
                              curyearbud = 0
                              curyearact = 0
                       rs.MoveNext
                      Loop
                  End If
         
         
         
             ElseIf cbosubmedia = "Online" Then
     
              Set ws = DBEngine.Workspaces(0)
              Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
              Sqlqry = "Select * from agndtls order by agentname"
              Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                   
                   If rs.RecordCount <> 0 Then
                      rs.MoveFirst
                          Do Until rs.EOF
                             curyearact = 0
                             lastyearact = 0
                             curyearbud = 0
                              
                            '  Current Year Actual
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearact = rs1.Fields(0)
                               
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearact = curyearact + Val(rs1.Fields(0) / convertion)
                               
                               
                            '  Last Year Actual
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then lastyearact = rs1.Fields(0)
                               
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then lastyearact = lastyearact + Val(rs1.Fields(0) / convertion)
                               
                               
                            '  Current Year Budget
                               Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "'  and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='USD'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearbud = rs1.Fields(0)
                               
                               Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='DHS'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearbud = curyearbud + Val(rs1.Fields(0) / convertion)
                               
                                   Sqlqry1 = " Insert into submediaanalysis values('" & findfirstfixup(Trim(rs!agentname)) & "','" & Trim(cbosubmedia) & "','" & _
                                            Trim(cboCurrency) & "'," & _
                                            lastyearact & "," & _
                                            curyearbud & "," & _
                                            curyearact & ")"
                                  
                                    ws.BeginTrans
                                    db.Execute (Sqlqry1)
                                    ws.CommitTrans
                          
                        
                        
                              lastyearact = 0
                              curyearbud = 0
                              curyearact = 0
                       rs.MoveNext
                      Loop
                  End If
            
         ElseIf cbosubmedia = "Cinema" Then
         
                    Set ws = DBEngine.Workspaces(0)
                    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                    Sqlqry = "Select * from agndtls order by agentname"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     If rs.RecordCount = 0 Then
                       MsgBox " No Records Found In The Agency Register"
                       Exit Sub
                     Else
                      rs.MoveFirst
                       Do Until rs.EOF
                        curyearact = 0
                        lastyearact = 0
                        curyearbud = 0
                        
                        '  Current Year Actual
                           Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) & "' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                           Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                           If IsNull(rs1.Fields(0)) = False Then curyearact = rs1.Fields(0)
                           
                           Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) & "' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                           Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                           If IsNull(rs1.Fields(0)) = False Then curyearact = curyearact + Val(rs1.Fields(0) / convertion)
                           
                        '  Last Year Actual
                           Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) - 1 & "' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                           Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                           If IsNull(rs1.Fields(0)) = False Then lastyearact = rs1.Fields(0)
                           
                           Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) - 1 & "' and Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                           Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                           If IsNull(rs1.Fields(0)) = False Then lastyearact = lastyearact + Val(rs1.Fields(0) / convertion)
                           
                           
                        '  Current Year Budget
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = curyearbud + Val(rs1.Fields(0) / convertion)
                      
                          Sqlqry1 = " Insert into submediaanalysis values('" & findfirstfixup(Trim(rs!agentname)) & "','" & Trim(cbosubmedia) & "','" & _
                                  Trim(cboCurrency) & "'," & _
                                  lastyearact & "," & _
                                  curyearbud & "," & _
                                  curyearact & ")"
                      
                            ws.BeginTrans
                            db.Execute (Sqlqry1)
                            ws.CommitTrans
                                              
                        lastyearact = 0
                        curyearbud = 0
                        curyearact = 0
                            
                      rs.MoveNext
                      Loop
                     End If
            
            
            
         
         ElseIf Mid(cbosubmedia, 1, 3) = "Cin" Then
         
                    Set ws = DBEngine.Workspaces(0)
                    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                    Sqlqry = "Select * from agndtls order by agentname"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     If rs.RecordCount = 0 Then
                       MsgBox " No Records Found In The Agency Register"
                       Exit Sub
                     Else
                      rs.MoveFirst
                       Do Until rs.EOF
                        curyearact = 0
                        lastyearact = 0
                        curyearbud = 0
                        
                        '  Current Year Actual
                           Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) & "' and sub_Media='" & Trim(Mid(cbosubmedia.Text, 10, 50)) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                           Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                           If IsNull(rs1.Fields(0)) = False Then curyearact = rs1.Fields(0)
                           
                           Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) & "' and Sub_Media='" & Trim(Mid(cbosubmedia.Text, 10, 50)) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                           Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                           If IsNull(rs1.Fields(0)) = False Then curyearact = curyearact + Val(rs1.Fields(0) / convertion)
                           
                        '  Last Year Actual
                           Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) - 1 & "' and sub_Media='" & Trim(Mid(cbosubmedia.Text, 10, 50)) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                           Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                           If IsNull(rs1.Fields(0)) = False Then lastyearact = rs1.Fields(0)
                           
                           Sqlqry1 = " select sum(Tra_namount) from bo_tracin where year ='" & Val(cboyear.Text) - 1 & "' and sub_Media='" & Trim(Mid(cbosubmedia.Text, 10, 50)) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                           Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                           If IsNull(rs1.Fields(0)) = False Then lastyearact = lastyearact + Val(rs1.Fields(0) / convertion)
                           
                           
                        '  Current Year Budget
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(Mid(cbosubmedia.Text, 10, 50)) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(Mid(cbosubmedia.Text, 10, 50)) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then curyearbud = curyearbud + Val(rs1.Fields(0) / convertion)
                      
                          Sqlqry1 = " Insert into submediaanalysis values('" & findfirstfixup(Trim(rs!agentname)) & "','" & Trim(Mid(cbosubmedia, 10, 50)) & "','" & _
                                  Trim(cboCurrency) & "'," & _
                                  lastyearact & "," & _
                                  curyearbud & "," & _
                                  curyearact & ")"
                      
                            ws.BeginTrans
                            db.Execute (Sqlqry1)
                            ws.CommitTrans
                                              
                        lastyearact = 0
                        curyearbud = 0
                        curyearact = 0
                            
                      rs.MoveNext
                      Loop
                     End If
                     
            
         
         
      Else
         
              Set ws = DBEngine.Workspaces(0)
              Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
              Sqlqry = "Select * from agndtls order by agentname"
              Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                   
                   If rs.RecordCount <> 0 Then
                      rs.MoveFirst
                          Do Until rs.EOF
                             curyearact = 0
                             lastyearact = 0
                             curyearbud = 0
                              
                            '  Current Year Actual
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and sub_Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearact = rs1.Fields(0)
                               
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) & "' and status='N' and sub_Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearact = curyearact + Val(rs1.Fields(0) / convertion)
                               
                               
                            '  Last Year Actual
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and sub_Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='USD'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then lastyearact = rs1.Fields(0)
                               
                               Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(cboyear.Text) - 1 & "' and status='N' and sub_Media='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency='DHS'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then lastyearact = lastyearact + Val(rs1.Fields(0) / convertion)
                               
                               
                            '  Current Year Budget
                               Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='USD'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearbud = rs1.Fields(0)
                               
                               Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & cboyear.Text & "' and Submedia='" & Trim(cbosubmedia.Text) & "' and Agency ='" & findfirstfixup(Trim(rs!agentname)) & "' and Tcurrency ='DHS'"
                               Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                               If IsNull(rs1.Fields(0)) = False Then curyearbud = curyearbud + Val(rs1.Fields(0) / convertion)
                               
                                   Sqlqry1 = " Insert into submediaanalysis values('" & findfirstfixup(Trim(rs!agentname)) & "','" & Trim(cbosubmedia) & "','" & _
                                            Trim(cboCurrency) & "'," & _
                                            lastyearact & "," & _
                                            curyearbud & "," & _
                                            curyearact & ")"
                                  
                                    ws.BeginTrans
                                    db.Execute (Sqlqry1)
                                    ws.CommitTrans
                          
                        
                        
                              lastyearact = 0
                              curyearbud = 0
                              curyearact = 0
                       rs.MoveNext
                      Loop
                  End If
                     
            
           End If
         
                     
           
 End Sub
 Private Sub cmdDisplay_Click()
    
   ' first if
   If ValidateData = True Then
   
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Delete * from submediaanalysis"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
     
    
    If cboCurrency.Text = "USD" Then
      PopulateUsd
    Else
      PopulateDhs
    End If
    
  With CrystalReport1
       .DataFiles(0) = App.Path & "\misov.mdb"
       .ReportFileName = App.Path & "\SubMediaAna.rpt"
       .Formulas(0) = "xxx3='" & Val(Val(cboyear.Text) - 1) & "'"
       .Formulas(1) = "xxx2='" & Val(cboyear.Text) & "'"
       .WindowMaxButton = True
       .WindowState = crptMaximized
       .Action = 1
  End With
    
   
   End If
  
End Sub

Private Sub Form_Load()
 Dim i
   populateMedia
   cboCurrency.AddItem "DHS"
   cboCurrency.AddItem "USD"
  
  i = 2000

For i = 2000 To 2100
 cboyear.AddItem i
Next

End Sub


Private Sub cbosubmedia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cboCurrency.SetFocus
End Sub

Private Function ValidateData()
 ValidateData = False

If cbosubmedia.Text = "" Then
   MsgBox "Invalid Sub Media", vbInformation, "Invalid Entry"
   cbosubmedia.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf cboyear.Text = "" Then
   MsgBox "Invalid year", vbInformation, "Invalid Entry"
   cboyear.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf cboCurrency.Text = "" Then
   MsgBox "Invalid Currency", vbInformation, "Invalid Entry"
   cboCurrency.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
   
   
Else
   ValidateData = True
End If

End Function

Private Sub textclear()
 cbosubmedia.ListIndex = -1
 cboyear.ListIndex = -1
 cboCurrency.ListIndex = -1
End Sub
