VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmTOTANALYSIS 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Last year Vs Current Year"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11805
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
         Begin VB.CommandButton cmdviewwoz 
            BackColor       =   &H00FFFF80&
            Caption         =   "P&review Only with Balances"
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
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2160
            Width           =   1335
         End
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
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   840
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
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   840
            Width           =   1815
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
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   2160
            Width           =   1215
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
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdDisplay 
            BackColor       =   &H00FFFF80&
            Caption         =   "P&review All"
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
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   2160
            Width           =   1455
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   6000
            Top             =   2520
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
            Left            =   5160
            TabIndex        =   8
            Top             =   840
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
            Left            =   1560
            TabIndex        =   5
            Top             =   840
            Width           =   1215
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   9720
            Y1              =   1800
            Y2              =   1800
         End
      End
   End
End
Attribute VB_Name = "frmTOTANALYSIS"
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

Dim dlyaalam As Currency
Dim dcybalam As Currency
Dim dcyaalam As Currency
Dim dlyazeina As Currency
Dim dcybzeina As Currency
Dim dcyazeina As Currency
Dim dlyaaawm As Currency
Dim dcybaawm As Currency
Dim dcyaaawm As Currency
Dim dlyaalburaq As Currency
Dim dcybalburaq As Currency
Dim dcyaalburaq As Currency
Dim dlyasoufra As Currency
Dim dcybsoufra As Currency
Dim dcyasoufra As Currency
Dim dlyacnn As Currency
Dim dcybcnn As Currency
Dim dcyacnn As Currency
Dim dlyacinema As Currency
Dim dcybcinema As Currency
Dim dcyacinema As Currency
Dim dlyamadina As Currency
Dim dcybmadina As Currency
Dim dcyamadina As Currency
Dim dlyaeim As Currency
Dim dcybeim As Currency
Dim dcyaeim As Currency
Dim dlyatot As Currency
Dim dcybtot As Currency
Dim dcyatot As Currency
Dim totper As Currency

Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
  
Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdDisplay.SetFocus
End Sub

Private Sub cmdBack_Click()
 Unload Me
End Sub

Private Sub cmdviewwoz_Click()
   If ValidateData = True Then
   
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Delete * from Totanalysis"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
     
    
    If CboCurrency.Text = "USD" Then
      PopulateUsd
    Else
      PopulateDhs
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Delete * from Totanalysis where lyatot=0 and cybtot=0 and cyatot=0"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
    
    
    
  With CrystalReport1
       .DataFiles(0) = App.Path & "\misov.mdb"
       .ReportFileName = App.Path & "\TotAnalysiswoz.rpt"
       .Formulas(0) = "xxx3='" & Val(Val(Cboyear.Text) - 1) & "'"
       .Formulas(1) = "xxx2='" & Val(Cboyear.Text) & "'"
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

Private Sub PopulateDhs()
      
            
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
                
                If rs!AGENTNAME = "STARCOM DB" Then
                 MsgBox "wAIT A MINUTE"
                End If
                
 dlyaalam = 0
 dcybalam = 0
 dcyaalam = 0
 dlyazeina = 0
 dcybzeina = 0
 dcyazeina = 0
 dlyaaawm = 0
 dcybaawm = 0
 dcyaaawm = 0
 dlyaalburaq = 0
 dcybalburaq = 0
 dcyaalburaq = 0
 dlyasoufra = 0
 dcybsoufra = 0
 dcyasoufra = 0
 dlyacnn = 0
 dcybcnn = 0
 dcyacnn = 0
 dlyacinema = 0
 dcybcinema = 0
 dcyacinema = 0
 dlyamadina = 0
 dcybmadina = 0
 dcyamadina = 0
 dlyaeim = 0
 dcybeim = 0
 dcyaeim = 0
 dlyatot = 0
 dcybtot = 0
 dcyatot = 0
 totper = 0
                    
                      
                      
                      
                      '  Sub Media  Alam Assayarat
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and sub_media='ALAM ASSAYARRAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS' and status='N'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaalam = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and sub_media='ALAM ASSAYARRAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD' and status='N'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaalam = dcyaalam + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Sub Media  ZEINA
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='ZEINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyazeina = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='ZEINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyazeina = dcyazeina + Val(rs1.Fields(0) * convertion)
                         
                       '  Sub Media  ALAM ASSAAT
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='ALAM ASSAAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaaawm = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='ALAM ASSAAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaaawm = dcyaaawm + Val(rs1.Fields(0) * convertion)
                         
                        '  Sub Media  AL Buraq
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='AL BURAQ' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaalburaq = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='AL BURAQ' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaalburaq = dcyaalburaq + Val(rs1.Fields(0) * convertion)
                         
                        '  Sub Media  SOUFRA DAIMEH
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='SOUFRA DAIMEH' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyasoufra = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='SOUFRA DAIMEH' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyasoufra = dcyasoufra + Val(rs1.Fields(0) * convertion)
                         
                                               
                        '  Sub Media  MADINA
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='MADINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyamadina = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='MADINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyamadina = dcyamadina + Val(rs1.Fields(0) * convertion)
                         
                       '  Sub Media  CNN
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='CNN' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyacnn = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='CNN' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyacnn = dcyacnn + Val(rs1.Fields(0) * convertion)
                        
                      '  Media  Online
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and media='Online' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaeim = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and media='Online' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaeim = dcyaeim + Val(rs1.Fields(0) * convertion)
                                              
                                                                                            
                      '  Media  Cinema
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and media='Cinema' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyacinema = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and media='Cinema' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyacinema = dcyacinema + Val(rs1.Fields(0) * convertion)
                         
                         dcyatot = dcyaalam + dcyazeina + dcyaaawm + dcyaalburaq + dcyasoufra + dcyacnn + dcyacinema + dcyamadina + dcyaeim
                      
                      ' Last year Actual
                      
                        '  Last year Sub Media  Alam Assayarat
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='ALAM ASSAYARRAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaalam = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='ALAM ASSAYARRAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaalam = dlyaalam + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Sub Media  ZEINA
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='ZEINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyazeina = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='ZEINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyazeina = dlyazeina + Val(rs1.Fields(0) * convertion)
                         
                       '  Sub Media  ALAM ASSAAT
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='ALAM ASSAAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaaawm = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='ALAM ASSAAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaaawm = dlyaaawm + Val(rs1.Fields(0) * convertion)
                         
                        '  Sub Media  AL Buraq
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='AL BURAQ' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaalburaq = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='AL BURAQ' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaalburaq = dlyaalburaq + Val(rs1.Fields(0) * convertion)
                         
                         '  Sub Media  SOUFRA DAIMEH
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='SOUFRA DAIMEH' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyasoufra = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='SOUFRA DAIMEH' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyasoufra = dlyasoufra + Val(rs1.Fields(0) * convertion)
                         
                                               
                        '  Sub Media  MADINA
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='MADINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyamadina = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='MADINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyamadina = dlyamadina + Val(rs1.Fields(0) * convertion)
                         
                       '  Sub Media  CNN
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='CNN' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyacnn = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='CNN' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyacnn = dlyacnn + Val(rs1.Fields(0) * convertion)
                        
                      '  Media  Online
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and media='Online' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaeim = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and media='Online' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaeim = dlyaeim + Val(rs1.Fields(0) * convertion)
                                              
                                                                                            
                      '  Media  Cinema
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and media='Cinema' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyacinema = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and and status='N' media='Cinema' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyacinema = dlyacinema + Val(rs1.Fields(0) * convertion)
                      
                                              
                                              
                        dlyatot = dlyaalam + dlyazeina + dlyaaawm + dlyaalburaq + dlyasoufra + dlyacnn + dlyacinema + dlyamadina + dlyaeim
                                              
                      ' Budget transaction
                      
                       '  Last year Sub Media  Alam Assayarat
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='ALAM ASSAYARRAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybalam = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='ALAM ASSAYARRAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybalam = dcybalam + Val(rs1.Fields(0) * convertion)
                         
                         
                      '  Sub Media  ZEINA
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='ZEINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybzeina = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='ZEINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybzeina = dcybzeina + Val(rs1.Fields(0) * convertion)
                         
                       '  Sub Media  ALAM ASSAAT
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='ALAM ASSAAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybaawm = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='ALAM ASSAAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybaawm = dcybaawm + Val(rs1.Fields(0) * convertion)
                         
                        '  Sub Media  AL Buraq
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='AL BURAQ' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybalburaq = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='AL BURAQ' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybalburaq = dcybalburaq + Val(rs1.Fields(0) * convertion)
                         
                         '  Sub Media  SOUFRA DAIMEH
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='SOUFRA DAIMEH' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybsoufra = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='SOUFRA DAIMEH' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybsoufra = dcybsoufra + Val(rs1.Fields(0) * convertion)
                         
                                               
                        '  Sub Media  MADINA
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='MADINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybmadina = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='MADINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybmadina = dcybmadina + Val(rs1.Fields(0) * convertion)
                         
                       '  Sub Media  CNN
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='CNN' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybcnn = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='CNN' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybcnn = dcybcnn + Val(rs1.Fields(0) * convertion)
                        
                      '  Media  Online
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='Online' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybeim = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='Online' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybeim = dcybeim + Val(rs1.Fields(0) * convertion)
                                              
                                                                                            
                      '  Media  CINEMA
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and mid(submedia,1,6)='Cinema' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybcinema = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and mid(submedia,1,6)='Cinema' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybcinema = dcybcinema + Val(rs1.Fields(0) * convertion)
                      
                         dcybtot = dcybalam + dcybzeina + dcybaawm + dcybalburaq + dcybsoufra + dcybcnn + dcybcinema + dcybmadina + dcybeim
                      
                       If dcybtot = 0 Then
                          totper = 0
                       Else
                         totper = Val(dcyatot * 100 / dcybtot)
                       End If
                                                                 
                      
                        Sqlqry1 = " Insert into totanalysis values('" & findfirstfixup(Trim(rs!AGENTNAME)) & "','" & Trim(CboCurrency) & "'," & _
                                  dlyaalam & "," & dcybalam & "," & dcyaalam & "," & _
                                  dlyazeina & "," & dcybzeina & "," & dcyazeina & "," & _
                                  dlyaaawm & "," & dcybaawm & "," & dcyaaawm & "," & _
                                  dlyaalburaq & "," & dcybalburaq & "," & dcyaalburaq & "," & _
                                  dlyasoufra & "," & dcybsoufra & "," & dcyasoufra & "," & _
                                  dlyacnn & "," & dcybcnn & "," & dcyacnn & "," & _
                                  dlyacinema & "," & dcybcinema & "," & dcyacinema & "," & _
                                  dlyamadina & "," & dcybmadina & "," & dcyamadina & "," & _
                                  dlyaeim & "," & dcybeim & "," & dcyaeim & "," & _
                                  dlyatot & "," & dcybtot & "," & dcyatot & "," & totper & ")"
                        ws.BeginTrans
                        db.Execute (Sqlqry1)
                        ws.CommitTrans
                    
                  
                  
                  
                      
                  rs.MoveNext
                  Loop
              End If
              
End Sub
 
 Private Sub PopulateUsd()
            
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
 dlyaalam = 0
 dcybalam = 0
 dcyaalam = 0
 dlyazeina = 0
 dcybzeina = 0
 dcyazeina = 0
 dlyaaawm = 0
 dcybaawm = 0
 dcyaaawm = 0
 dlyaalburaq = 0
 dcybalburaq = 0
 dcyaalburaq = 0
 dlyasoufra = 0
 dcybsoufra = 0
 dcyasoufra = 0
 dlyacnn = 0
 dcybcnn = 0
 dcyacnn = 0
 dlyacinema = 0
 dcybcinema = 0
 dcyacinema = 0
 dlyamadina = 0
 dcybmadina = 0
 dcyamadina = 0
 dlyaeim = 0
 dcybeim = 0
 dcyaeim = 0
 dlyatot = 0
 dcybtot = 0
 dcyatot = 0
 totper = 0
                    
                      
                      
                      
                      '  Sub Media  Alam Assayarat
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='ALAM ASSAYARRAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaalam = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='ALAM ASSAYARRAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaalam = dcyaalam + Val(rs1.Fields(0) / convertion)
                         
                         
                      '  Sub Media  ZEINA
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='ZEINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyazeina = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='ZEINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyazeina = dcyazeina + Val(rs1.Fields(0) / convertion)
                         
                       '  Sub Media  ALAM ASSAAT
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='ALAM ASSAAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaaawm = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='ALAM ASSAAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaaawm = dcyaaawm + Val(rs1.Fields(0) / convertion)
                         
                        '  Sub Media  AL Buraq
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='AL BURAQ' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaalburaq = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='AL BURAQ' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaalburaq = dcyaalburaq + Val(rs1.Fields(0) / convertion)
                         
                         '  Sub Media  SOUFRA DAIMEH
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='SOUFRA DAIMEH' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyasoufra = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='SOUFRA DAIMEH' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyasoufra = dcyasoufra + Val(rs1.Fields(0) / convertion)
                         
                                               
                        '  Sub Media  MADINA
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='MADINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyamadina = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and sub_media='MADINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyamadina = dcyamadina + Val(rs1.Fields(0) / convertion)
                         
                       '  Sub Media  CNN
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='CNN' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyacnn = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and sub_media='CNN' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyacnn = dcyacnn + Val(rs1.Fields(0) / convertion)
                        
                      '  Media  Online
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and media='Online' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaeim = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and media='Online' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyaeim = dcyaeim + Val(rs1.Fields(0) / convertion)
                                              
                                                                                            
                      '  Media  cINEMA
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and media='Cinema' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyacinema = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) & "' and status='N' and media='Cinema' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcyacinema = dcyacinema + Val(rs1.Fields(0) / convertion)
                         
                         dcyatot = dcyaalam + dcyazeina + dcyaaawm + dcyaalburaq + dcyasoufra + dcyacnn + dcyacinema + dcyamadina + dcyaeim
                      
                      ' Last year Actual
                      
                        '  Last year Sub Media  Alam Assayarat
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='ALAM ASSAYARRAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaalam = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='ALAM ASSAYARRAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaalam = dlyaalam + Val(rs1.Fields(0) / convertion)
                         
                         
                      '  Sub Media  ZEINA
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='ZEINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyazeina = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='ZEINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyazeina = dlyazeina + Val(rs1.Fields(0) / convertion)
                         
                       '  Sub Media  ALAM ASSAAT
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='ALAM ASSAAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaaawm = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='ALAM ASSAAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaaawm = dlyaaawm + Val(rs1.Fields(0) / convertion)
                         
                        '  Sub Media  AL Buraq
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='AL BURAQ' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaalburaq = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='AL BURAQ' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaalburaq = dlyaalburaq + Val(rs1.Fields(0) / convertion)
                         
                         '  Sub Media  SOUFRA DAIMEH
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='SOUFRA DAIMEH' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyasoufra = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='SOUFRA DAIMEH' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyasoufra = dlyasoufra + Val(rs1.Fields(0) / convertion)
                         
                                               
                        '  Sub Media  MADINA
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='MADINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyamadina = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='MADINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyamadina = dlyamadina + Val(rs1.Fields(0) / convertion)
                         
                       '  Sub Media  CNN
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='CNN' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyacnn = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and sub_media='CNN' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyacnn = dlyacnn + Val(rs1.Fields(0) / convertion)
                        
                      '  Media  Online
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and media='Online' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaeim = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and media='Online' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyaeim = dlyaeim + Val(rs1.Fields(0) / convertion)
                                              
                                                                                            
                      '  Media  cINEMA
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and media='Cinema' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyacinema = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Tra_namount) from bo_mas where year ='" & Val(Cboyear.Text) - 1 & "' and status='N' and media='Cinema' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dlyacinema = dlyacinema + Val(rs1.Fields(0) / convertion)
                      
                                              
                                              
                        dlyatot = dlyaalam + dlyazeina + dlyaaawm + dlyaalburaq + dlyasoufra + dlyacnn + dlyacinema + dlyamadina + dlyaeim
                                              
                      ' Budget transaction
                      
                       '  Last year Sub Media  Alam Assayarat
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='ALAM ASSAYARRAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybalam = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='ALAM ASSAYARRAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybalam = dcybalam + Val(rs1.Fields(0) / convertion)
                         
                         
                      '  Sub Media  ZEINA
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='ZEINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybzeina = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='ZEINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybzeina = dcybzeina + Val(rs1.Fields(0) / convertion)
                         
                       '  Sub Media  ALAM ASSAAT
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='ALAM ASSAAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybaawm = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='ALAM ASSAAT' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybaawm = dcybaawm + Val(rs1.Fields(0) / convertion)
                         
                        '  Sub Media  AL Buraq
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='AL BURAQ' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybalburaq = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='AL BURAQ' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybalburaq = dcybalburaq + Val(rs1.Fields(0) / convertion)
                         
                         '  Sub Media  SOUFRA DAIMEH
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='SOUFRA DAIMEH' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybsoufra = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='SOUFRA DAIMEH' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybsoufra = dcybsoufra + Val(rs1.Fields(0) / convertion)
                         
                                               
                        '  Sub Media  MADINA
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='MADINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybmadina = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='MADINA' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybmadina = dcybmadina + Val(rs1.Fields(0) / convertion)
                         
                       '  Sub Media  CNN
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='CNN' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybcnn = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='CNN' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybcnn = dcybcnn + Val(rs1.Fields(0) / convertion)
                        
                      '  Media  Online
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='Online' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybeim = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and submedia='Online' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybeim = dcybeim + Val(rs1.Fields(0) / convertion)
                                              
                                                                                            
                      '  Media  cINEMA
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and mid(submedia,1,6)='Cinema' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='USD'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybcinema = rs1.Fields(0)
                         
                         Sqlqry1 = " select sum(Budget) from agmediabudget where tyear ='" & Val(Cboyear.Text) & "' and mid(submedia,1,6)='Cinema' and Agency ='" & findfirstfixup(Trim(rs!AGENTNAME)) & "' and Tcurrency='DHS'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then dcybcinema = dcybcinema + Val(rs1.Fields(0) / convertion)
                      
                         dcybtot = dcybalam + dcybzeina + dcybaawm + dcybalburaq + dcybsoufra + dcybcnn + dcybcinema + dcybmadina + dcybeim
                      
                       If dcybtot = 0 Then
                          totper = 0
                       Else
                         totper = Val(dcyatot * 100 / dcybtot)
                       End If
                                                                 
                      
                        Sqlqry1 = " Insert into totanalysis values('" & findfirstfixup(Trim(rs!AGENTNAME)) & "','" & Trim(CboCurrency) & "'," & _
                                  dlyaalam & "," & dcybalam & "," & dcyaalam & "," & _
                                  dlyazeina & "," & dcybzeina & "," & dcyazeina & "," & _
                                  dlyaaawm & "," & dcybaawm & "," & dcyaaawm & "," & _
                                  dlyaalburaq & "," & dcybalburaq & "," & dcyaalburaq & "," & _
                                  dlyasoufra & "," & dcybsoufra & "," & dcyasoufra & "," & _
                                  dlyacnn & "," & dcybcnn & "," & dcyacnn & "," & _
                                  dlyacinema & "," & dcybcinema & "," & dcyacinema & "," & _
                                  dlyamadina & "," & dcybmadina & "," & dcyamadina & "," & _
                                  dlyaeim & "," & dcybeim & "," & dcyaeim & "," & _
                                  dlyatot & "," & dcybtot & "," & dcyatot & "," & totper & ")"
                        ws.BeginTrans
                        db.Execute (Sqlqry1)
                        ws.CommitTrans
                    
                  
                  
                      
                  rs.MoveNext
                  Loop
              End If
 End Sub
 
 Private Sub cmdDisplay_Click()
    
   ' first if
   
   If ValidateData = True Then
   
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Delete * from Totanalysis"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
     
    
    If CboCurrency.Text = "USD" Then
      PopulateUsd
    Else
      PopulateDhs
    End If
    
  With CrystalReport1
       .DataFiles(0) = App.Path & "\misov.mdb"
       .ReportFileName = App.Path & "\TotAnalysis.rpt"
       .Formulas(0) = "xxx3='" & Val(Val(Cboyear.Text) - 1) & "'"
       .Formulas(1) = "xxx2='" & Val(Cboyear.Text) & "'"
       .WindowMaxButton = True
       .WindowState = crptMaximized
       .Action = 1
  End With
    
   
   End If
  
End Sub

Private Sub Form_Load()
 Dim i
  
   CboCurrency.AddItem "DHS"
   CboCurrency.AddItem "USD"
  
  i = 2000

For i = 2000 To 2100
 Cboyear.AddItem i
Next

End Sub

Private Function ValidateData()
 ValidateData = False

If Cboyear.Text = "" Then
   MsgBox "Invalid year", vbInformation, "Invalid Entry"
   Cboyear.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
ElseIf CboCurrency.Text = "" Then
   MsgBox "Invalid Currency", vbInformation, "Invalid Entry"
   CboCurrency.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
   ValidateData = True
End If

End Function

Private Sub textclear()
 Cboyear.ListIndex = -1
 CboCurrency.ListIndex = -1
End Sub
