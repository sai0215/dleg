VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmDealReport 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   8595
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmDealReport.frx":0000
      Left            =   960
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   8415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11655
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Height          =   3615
         Left            =   6840
         TabIndex        =   13
         Top             =   3360
         Width           =   4455
         Begin VB.TextBox txtselclients 
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   120
            TabIndex        =   15
            Text            =   "Selected clients to include in the Deal"
            Top             =   120
            Width           =   4215
         End
         Begin VB.ListBox lstClientsselected 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2940
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   14
            Top             =   480
            Width           =   4215
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Height          =   3615
         Left            =   360
         TabIndex        =   10
         Top             =   3360
         Width           =   4695
         Begin VB.TextBox txtname1 
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   120
            TabIndex        =   12
            Text            =   "Total Clients for the selected Agency"
            Top             =   120
            Width           =   4455
         End
         Begin VB.ListBox lstclientsttl 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2940
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   11
            Top             =   480
            Width           =   4455
         End
      End
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Print Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         Picture         =   "frmDealReport.frx":001D
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton cmdfrom 
         BackColor       =   &H80000016&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton cmdToall 
         BackColor       =   &H80000016&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdFromAll 
         BackColor       =   &H80000016&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton cmdto 
         BackColor       =   &H80000016&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4200
         Width           =   1095
      End
      Begin VB.ListBox lstDeals 
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
         Height          =   2220
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   10815
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
         Left            =   4920
         Picture         =   "frmDealReport.frx":011F
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   7440
         Width           =   1335
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
         Left            =   6240
         Picture         =   "frmDealReport.frx":0561
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7440
         Width           =   1455
      End
      Begin VB.Label lblRemarks 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   480
         TabIndex        =   16
         Top             =   2760
         Width           =   10695
      End
      Begin VB.Label lblSubMediaName 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   6480
         TabIndex        =   3
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   11880
         Y1              =   7320
         Y2              =   7320
      End
   End
End
Attribute VB_Name = "frmDealReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim i As Integer
Dim X, Y, Z, f As Integer
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Public zzz, zzzmed, zzzcin, zzzmag, zzzol, zzztv As Currency

'Private Sub ChkAll_Click()
'If f = 1 Then

' ChkCinema.Enabled = True
' ChkMagazine.Enabled = True
' ChkTelevision.Enabled = True
' ChkOnline.Enabled = True
' ChkAll.Value = 0
' f = 0
'Else
' ChkCinema.Enabled = False
' ChkMagazine.Enabled = False
' ChkTelevision.Enabled = False
' ChkOnline.Enabled = False
' ChkCinema.Value = 0
' ChkMagazine.Value = 0
' ChkTelevision.Value = 0
' ChkOnline.Value = 0
' ChkAll.Value = 1
' f = 1
'End If
'End Sub
Private Sub cmdBack_Click()
  Unload Me
End Sub
Private Sub cmdClear_Click()
  textclear
End Sub
Private Sub textclear()
    lstDeals.ListIndex = -1
    lblRemarks.Caption = ""
    lblRemarks.Visible = False
    lstclientsttl.Clear
    lstClientsselected.Clear
End Sub
Private Sub cmdDisplay_Click()
 Dim i
 Dim a, B, C
 Dim dealcur
 Dim stdate As Date
 Dim EdDate As Date
 Dim DEALAGENCY
 Dim totpgmedia, totpgcin, totpgmag, totpgol, totpgtv As Currency
 Dim totpnmedia, totpncin, totpnmag, totpnol, totpntv As Currency
 Dim totfmedia, totfcin, totfmag, totfol, totftv As Currency
 Dim totbmedia, totbcin, totbmag, totbol, totbtv As Currency
 Dim dvol1, dvol2, dvol3, dvol4 As Currency
 Dim dvol1disc, dvol2disc, dvol3disc, dvol4disc As Currency
 Dim ReqDisc As Currency
 
 
 If lstDeals.Text = "" Then
    MsgBox "Invalid Deals", vbInformation, "Invalid Entry"
    lstDeals.SetFocus
    SendKeys " {Home} + {End} "
    Exit Sub
 End If
   
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
  Sqlqry = "Delete * from dumbo_mas"
  ws.BeginTrans
  db.Execute (Sqlqry)
  ws.CommitTrans
   
   
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
  Sqlqry = "Delete * from dealmonitor"
  ws.BeginTrans
  db.Execute (Sqlqry)
  ws.CommitTrans
   
   
   If lstClientsselected.ListCount = 0 Then
    MsgBox "Select Clients ", vbInformation, "Invalid Entry"
    lstclientsttl.SetFocus
    SendKeys " {Home} + {End} "
    Exit Sub
   End If
   
   stdate = Now()
   EdDate = Now()
   dealcur = ""
   DEALAGENCY = ""
   
   dvol1 = 0
   dvol2 = 0
   dvol3 = 0
   dvol4 = 0
   
   dvol1disc = 0
   dvol2disc = 0
   dvol3disc = 0
   dvol4disc = 0
   
   ReqDisc = 0
   
   
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = "Select * from deal where name ='" & lstDeals.Text & "'"
   Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
   If rs.RecordCount <> 0 Then
     stdate = Format(rs!DATEFROM, "DD/MM/YYYY")
     EdDate = Format(rs!Dateto, "DD/MM/YYYY")
     dealcur = Trim(rs!tcurrency)
     DEALAGENCY = Trim(rs!Agency)
     dvol1 = Trim(rs!vol1)
     dvol2 = Trim(rs!vol2)
     dvol3 = Trim(rs!vol3)
     dvol4 = Trim(rs!vol4)
     
     dvol1disc = Trim(rs!vol1disc)
     dvol2disc = Trim(rs!vol2disc)
     dvol3disc = Trim(rs!vol3disc)
     dvol4disc = Trim(rs!vol4disc)
     
   End If
   
     If dealcur = "DHS" Then
                    f = lstClientsselected.ListIndex
                    Set ws = DBEngine.Workspaces(0)
                    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                    For f = 0 To lstClientsselected.ListCount - 1
                     'Invoice_Date Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "# order by media"
                         Sqlqry = "Select * from Bo_mas where tdate >=#" & DateValue(Format(stdate, "dd/mm/yyyy")) & "#  and  tdate<=#" & DateValue(Format(EdDate, "dd/mm/yyyy")) & "# and Client='" & Trim(lstClientsselected.List(f)) & "' AND AGENCY='" & findfirstfixup(DEALAGENCY) & "' and cancell='N'"
                         Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                         If rs.RecordCount <> 0 Then
                            rs.MoveFirst
                            Do Until rs.EOF
                             If rs!tcurrency = "USD" Then
                                    Set ws = DBEngine.Workspaces(0)
                                    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                                    Sqlqry1 = " Insert into DumBo_Mas values('" & rs!serial_no & "','" & findfirstfixup(Trim(lstDeals)) & "','" & Trim(rs!tDate) & "','DHS'," & rs!tconvertion & "," & Val(rs!tra_gamount) * convertion & "," & Round(Val(rs!tra_namount * convertion), 2) & ",'" & rs!Year & "','" _
                                                        & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                                        & findfirstfixup(Trim(rs!region)) & "','" & findfirstfixup(Trim(rs!boremarks)) & "','" _
                                                        & findfirstfixup(rs!Product) & "','" _
                                                        & findfirstfixup(rs!client) & "','" _
                                                        & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                        & Trim(rs!sub_Media) & "','" _
                                                        & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                                        & rs!gross_amount & "," _
                                                        & Round(Val(rs!Tot_free) * convertion, 2) & "," _
                                                        & Round(Val(rs!Tot_barter) * convertion, 2) & ",'" _
                                                        & Val(rs!disc_percentage) & "','" _
                                                        & Val(rs!disc_rate) & "'," _
                                                        & Round(Val(rs!add_discount) * convertion, 2) & "," & Round(Val(rs!surcharge) * convertion, 2) & "," _
                                                        & rs!NET_Amount & ",'" & Trim(rs!invoice_date) & "','" & Trim(rs!acct_code) & "','" & rs!Status & "','" & rs!cancell & "')"
                                       ws.BeginTrans
                                       db.Execute (Sqlqry1)
                                       ws.CommitTrans
                                Else
                                    Set ws = DBEngine.Workspaces(0)
                                    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                                    Sqlqry1 = " Insert into DumBo_Mas values('" & rs!serial_no & "','" & findfirstfixup(Trim(lstDeals.Text)) & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & Val(rs!tra_gamount) & "," & Val(rs!tra_namount) & ",'" & rs!Year & "','" _
                                                        & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                                        & findfirstfixup(Trim(rs!region)) & "','" & findfirstfixup(Trim(rs!boremarks)) & "','" _
                                                        & findfirstfixup(rs!Product) & "','" _
                                                        & findfirstfixup(rs!client) & "','" _
                                                        & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                        & Trim(rs!sub_Media) & "','" _
                                                        & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                                        & Val(rs!gross_amount) & "," _
                                                        & Val(rs!Tot_free) & "," _
                                                        & Val(rs!Tot_barter) & ",'" _
                                                        & Val(rs!disc_percentage) & "','" _
                                                        & Val(rs!disc_rate) & "'," _
                                                        & Val(rs!add_discount) & "," & Val(rs!surcharge) & "," _
                                                        & Val(rs!NET_Amount) & ",'" & Trim(rs!invoice_date) & "','" & Trim(rs!acct_code) & "','" & rs!Status & "','" & rs!cancell & "')"
                                       ws.BeginTrans
                                       db.Execute (Sqlqry1)
                                       ws.CommitTrans
                               End If
                               rs.MoveNext
                               Loop
                          End If
                    Next
      Else
       ' if dealcur = USD
        f = lstClientsselected.ListIndex
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        For f = 0 To lstClientsselected.ListCount - 1
        'Invoice_Date Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "# order by media"
             Sqlqry = "Select * from Bo_mas where tdate >=#" & DateValue(Format(stdate, "dd/mm/yyyy")) & "#  and  tdate<=#" & DateValue(Format(EdDate, "dd/mm/yyyy")) & "# and Client='" & Trim(lstClientsselected.List(f)) & "' AND AGENCY='" & findfirstfixup(DEALAGENCY) & "' and cancell='N'"
             Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
             If rs.RecordCount <> 0 Then
                rs.MoveFirst
                Do Until rs.EOF
                 If rs!tcurrency = "DHS" Then
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                        Sqlqry1 = " Insert into DumBo_Mas values('" & rs!serial_no & "','" & findfirstfixup(Trim(lstDeals)) & "','" & Trim(rs!tDate) & "','USD'," & rs!tconvertion & "," & Round(Val(rs!tra_gamount) / convertion, 2) & "," & Round(Val(rs!tra_namount) / convertion, 2) & ",'" & rs!Year & "','" _
                                            & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                            & findfirstfixup(Trim(rs!region)) & "','" & findfirstfixup(Trim(rs!boremarks)) & "','" _
                                            & findfirstfixup(rs!Product) & "','" _
                                            & findfirstfixup(rs!client) & "','" _
                                            & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                            & Trim(rs!sub_Media) & "','" _
                                            & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                            & Round(Val(rs!gross_amount) / convertion, 2) & "," _
                                            & Round(Val(rs!Tot_free) / convertion, 2) & "," _
                                            & Round(Val(rs!Tot_barter) / convertion, 2) & ",'" _
                                            & Val(rs!disc_percentage) & "','" _
                                            & Val(rs!disc_rate) & "'," _
                                            & Round(Val(rs!add_discount) / convertion, 2) & "," & Round(Val(rs!surcharge) / convertion, 2) & "," _
                                            & Round(Val(rs!NET_Amount) / convertion, 2) & ",'" & Trim(rs!invoice_date) & "','" & Trim(rs!acct_code) & "','" & rs!Status & "','" & rs!cancell & "')"
                           ws.BeginTrans
                           db.Execute (Sqlqry1)
                           ws.CommitTrans
                    Else
                        Set ws = DBEngine.Workspaces(0)
                        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                        Sqlqry1 = " Insert into DumBo_Mas values('" & rs!serial_no & "','" & findfirstfixup(Trim(lstDeals.Text)) & "','" & Trim(rs!tDate) & "','" & rs!tcurrency & "'," & rs!tconvertion & "," & Val(rs!tra_gamount) & "," & Val(rs!tra_namount) & ",'" & rs!Year & "','" _
                                            & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                            & findfirstfixup(Trim(rs!region)) & "','" & findfirstfixup(Trim(rs!boremarks)) & "','" _
                                            & findfirstfixup(rs!Product) & "','" _
                                            & findfirstfixup(rs!client) & "','" _
                                            & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                            & Trim(rs!sub_Media) & "','" _
                                            & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                            & Round(Val(rs!gross_amount) / convertion, 2) & "," _
                                            & Val(rs!Tot_free) & "," _
                                            & Val(rs!Tot_barter) & ",'" _
                                            & Val(rs!disc_percentage) & "','" _
                                            & Val(rs!disc_rate) & "'," _
                                            & Val(rs!add_discount) & "," & Val(rs!surcharge) & "," _
                                            & Round(Val(rs!NET_Amount) / convertion, 2) & ",'" & Trim(rs!invoice_date) & "','" & Trim(rs!acct_code) & "','" & rs!Status & "','" & rs!cancell & "')"
                           ws.BeginTrans
                           db.Execute (Sqlqry1)
                           ws.CommitTrans
                   End If
                   rs.MoveNext
                   Loop
              End If
        Next
    End If
    
 totpgmedia = 0
 totpgcin = 0
 totpgmag = 0
 totpgtv = 0
 totpgol = 0
 
 totpnmedia = 0
 totpncin = 0
 totpnmag = 0
 totpntv = 0
 totpnol = 0
 
 totfmedia = 0
 totfcin = 0
 totfmag = 0
 totftv = 0
 totfol = 0
 
 totbmedia = 0
 totbcin = 0
 totbmag = 0
 totbtv = 0
 totbol = 0
 
        'Gross totals
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(tra_gamount) from dumbo_mas"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totpgmedia = rs1.Fields(0)
        
        
        If totpgmedia <= dvol1 Then
         ReqDisc = dvol1disc
        ElseIf totpgmedia > dvol1 And totpgmedia <= dvol2 Then
         ReqDisc = dvol2disc
        ElseIf totpgmedia > dvol2 And totpgmedia <= dvol3 Then
         ReqDisc = dvol3disc
        ElseIf totpgmedia > dvol3 Then
         ReqDisc = dvol4disc
        End If
        
        
        
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(tra_gamount) from dumbo_mas where Media='Cinema'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totpgcin = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(tra_gamount) from dumbo_mas where Media='Magazine'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totpgmag = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(tra_gamount) from dumbo_mas where Media='Online'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totpgol = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(tra_gamount) from dumbo_mas where Media='Television'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totpgtv = rs1.Fields(0)
        
        
        'Net totals
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(tra_namount) from dumbo_mas"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totpnmedia = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(tra_namount) from dumbo_mas where Media='Cinema'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totpncin = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(tra_namount) from dumbo_mas where Media='Magazine'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totpnmag = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(tra_namount) from dumbo_mas where Media='Online'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totpnol = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(tra_namount) from dumbo_mas where Media='Television'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totpntv = rs1.Fields(0)
        
        
        'free
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(Tot_free) from dumbo_mas"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totfmedia = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(Tot_free) from dumbo_mas where Media='Cinema'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totfcin = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(Tot_free) from dumbo_mas where Media='Magazine'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totfmag = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(Tot_free) from dumbo_mas where Media='Online'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totfol = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(Tot_free) from dumbo_mas where Media='Television'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totftv = rs1.Fields(0)
        
        

        'Barter
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(Tot_Barter) from dumbo_mas"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totbmedia = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(Tot_Barter) from dumbo_mas where Media='Cinema'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totbcin = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(Tot_Barter) from dumbo_mas where Media='Magazine'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totbmag = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(Tot_Barter) from dumbo_mas where Media='Online'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totbol = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = " select sum(Tot_Barter) from dumbo_mas where Media='Television'"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If IsNull(rs1.Fields(0)) = False Then totbtv = rs1.Fields(0)
        
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = "Select * from deal where name='" & Trim(lstDeals) & "'"
        Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs.RecordCount <> 0 Then
        rs.MoveFirst
            Sqlqry2 = " Insert into dealMonitor values('" & rs!serial_no & "','" & rs!Name & "','" & rs!tcurrency & "'," _
                                                 & totpgmedia & " ," & totpgcin & "," & totpgmag & "," & totpgol & "," & totpgtv & "," _
                                                 & totfmedia & " ," & totfcin & "," & totfmag & "," & totfol & "," & totftv & "," _
                                                 & totbmedia & " ," & totbcin & "," & totbmag & "," & totbol & "," & totbtv & "," _
                                                 & totpnmedia & " ," & totpncin & "," & totpnmag & "," & totpnol & "," & totpntv & ")"
                                       
            ws.BeginTrans
            db.Execute (Sqlqry2)
            ws.CommitTrans
             
        End If
    ' Dim x As Currency
        
    With CrystalReport1
      .DataFiles(0) = App.Path & "\misov.mdb"
      .ReportFileName = App.Path & "\dealfinal.rpt"
      .Formulas(0) = "zzz=" & Round(ReqDisc, 0) & ""
      .Formulas(1) = "zzzmed=" & Round(totpgmedia * ReqDisc / 100, 0) & ""
      .Formulas(2) = "zzzcin=" & Round(totpgcin * ReqDisc / 100, 0) & ""
      .Formulas(3) = "zzzmag=" & Round(totpgmag * ReqDisc / 100, 0) & ""
      .Formulas(4) = "zzzol=" & Round(totpgol * ReqDisc / 100, 0) & ""
      .Formulas(5) = "zzztv=" & Round(totpgtv * ReqDisc / 100, 0) & ""
      .WindowState = crptMaximized
      .Action = 1
   End With
    
End Sub

Private Sub cmdfrom_Click()
 For f = lstClientsselected.ListCount - 1 To 0 Step -1
    
    If lstClientsselected.Selected(f) Then
       lstclientsttl.AddItem lstClientsselected.Text
       lstClientsselected.RemoveItem (f)
    End If
 Next
End Sub

Private Sub cmdFromAll_Click()
For i = lstClientsselected.ListCount - 1 To 0 Step -1
         lstclientsttl.AddItem lstClientsselected.List(i)
         lstClientsselected.RemoveItem (i)
 Next i
End Sub

Private Sub cmdto_Click()
 For i = lstclientsttl.ListCount - 1 To 0 Step -1
    If lstclientsttl.Selected(i) Then
       lstClientsselected.AddItem lstclientsttl.List(i)
       lstclientsttl.RemoveItem (i)
    End If
Next
End Sub

Private Sub cmdToall_Click()
 For i = lstclientsttl.ListCount - 1 To 0 Step -1
         lstClientsselected.AddItem lstclientsttl.List(i)
         lstclientsttl.RemoveItem (i)
 Next i
End Sub

Private Sub Form_Load()
   Populatedeals
   lblRemarks.Visible = False
   
End Sub

Private Sub Populatedeals()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
Sqlqry = "Select * from deal order by name"
Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)

lstDeals.Clear

If rs.RecordCount = 0 Then
    MsgBox "No Records found in the Deal Register"
Else
   rs.MoveFirst
   Do Until rs.EOF
      lstDeals.AddItem rs!Name
      rs.MoveNext
   Loop
End If
End Sub

Private Sub lstDeals_Click()
Dim i
Dim tempBln As String

    If lstDeals.ListIndex = -1 Then
        tempBln = False
    Else
        tempBln = True
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    i = Trim(lstDeals)
        Sqlqry = "Select * from deal Where name= '" & i & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
         If rs.RecordCount <> 0 Then
            lblRemarks.Caption = ""
            lblRemarks.Visible = True
            rs.MoveFirst
            lblRemarks.Caption = rs!remarks
                Sqlqry1 = " select distinct(client) from bo_mas where agency='" & rs!Agency & "'"
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                lstclientsttl.Clear
                lstClientsselected.Clear
                If rs1.RecordCount <> 0 Then
                   rs1.MoveFirst
                    Do Until rs1.EOF
                     lstclientsttl.AddItem rs1!client
                     rs1.MoveNext
                    Loop
                End If
          End If
            
End Sub


