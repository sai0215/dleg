VERSION 5.00
Begin VB.Form frmdatarefresh 
   Caption         =   "Data Refresh"
   ClientHeight    =   5715
   ClientLeft      =   1665
   ClientTop       =   1395
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdmag 
      Caption         =   "Botramag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      TabIndex        =   1
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton cmdDataRefresh 
      Caption         =   "Data Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
End
Attribute VB_Name = "frmdatarefresh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws As Workspace
Dim rs As Recordset
Dim rs1 As Recordset
Dim sqlqry As String
Dim sqlqry1 As String
Dim db As Database

Private Sub cmdback_Click()
 Unload Me
End Sub

Private Sub cmdDataRefresh_Click()
Dim X

 X = MsgBox("Do you want to delete all the transactions in the misovdum.mdb database", vbCritical + vbYesNo, "Confirmation")
  If X = vbYes Then
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from bo_mas"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from bo_tracin"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from bo_tramag"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from bo_traol"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from bo_tratv"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from bpmt_mas"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from bpmt_tra"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from brpt_mas"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from brpt_tra"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from capr_mas"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from capr_tra"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from cinema_rates"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from cpmt_mas"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from cpmt_tra"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from crdt_mas"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from crpr_mas"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from crpr_tra"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from crpt_mas"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from crpt_tra"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
   Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from debt_mas"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from jrnl_tra"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from material"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from media"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from ppmt_mas"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from ppmt_tra"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from products"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from prpt_mas1"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Delete * from prpt_tra"
    ws.BeginTrans
    db.Execute (sqlqry)
    ws.CommitTrans
    
    
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misovdum.mdb")
  sqlqry = "Update docu_mas set doc_no=" & 100000 & " where doc_no>=" & Val(100000) & ""
    ws.BeginTrans
    db.Execute sqlqry
    ws.CommitTrans
    
End If
End Sub

Private Sub cmdmag_Click()
Dim X

  X = MsgBox("Do you want to Modify all the transactions in the misov.mdb database", vbCritical + vbYesNo, "Confirmation")
  If X = vbYes Then
  
   Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
  sqlqry = "Select * from bo_mas order by serial_no "
 Set rs = db.OpenRecordset(sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                         sqlqry1 = "Update Bo_tramag set agcom='" & Trim(rs!disc_rate) & "'," & _
                                            " adper ='" & Trim(rs!disc_percentage) & "'," & _
                                            " addisc =" & Val(rs!add_discount) & "," & _
                                            " surcharge=" & Val(rs!surcharge) & " Where serial_NO = '" & Val(rs!serial_no) & "'"
                                            
                 
                        
                         ws.BeginTrans
                         db.Execute (sqlqry1)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
    
  
End If
End Sub
