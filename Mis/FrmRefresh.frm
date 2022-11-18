VERSION 5.00
Begin VB.Form FrmRefresh 
   Caption         =   "Refresh"
   ClientHeight    =   5715
   ClientLeft      =   1095
   ClientTop       =   1755
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.CommandButton CmdFresh 
      Caption         =   "Refresh"
      Height          =   1215
      Left            =   2520
      TabIndex        =   0
      Top             =   3480
      Width           =   1695
   End
End
Attribute VB_Name = "FrmRefresh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim agnnm As String
Dim ws As Workspace
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Private Sub CmdFresh_Click()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Select * from bo_mas order by serial_no where Media='Magazine'"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
          rs.MoveFirst
          Do Until rs.EOF
            Sqlqry1 = " Select * from bo_tramag where serial_no='" & Trim(rs!serial_no) & "'"
            Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
            
            
          
           MsgBox " Agency Already existing"
          Exit Sub
        Else
          Sqlqry1 = " Insert into a"

End Sub
