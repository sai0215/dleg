VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "PVMASK.OCX"
Begin VB.Form frmCinRepScMonRevOriginal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Screen  Monitoring "
   ClientHeight    =   8595
   ClientLeft      =   -75
   ClientTop       =   225
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   4920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame FraMain 
      BackColor       =   &H00FFFFFF&
      Height          =   7935
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   10935
      Begin VB.Frame Fradatesel 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   2280
         TabIndex        =   6
         Top             =   6600
         Width           =   6375
         Begin VB.CommandButton cmddtprint 
            BackColor       =   &H00C0E0FF&
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
            Left            =   360
            Picture         =   "frmCinRepScMonRevOriginal.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton CmdDtback 
            BackColor       =   &H00C0E0FF&
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
            Height          =   735
            Left            =   3960
            Picture         =   "frmCinRepScMonRevOriginal.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton cmddtClear 
            BackColor       =   &H00C0E0FF&
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
            Left            =   2160
            Picture         =   "frmCinRepScMonRevOriginal.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Fradate 
         BackColor       =   &H00FFFFFF&
         Height          =   6255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   10695
         Begin VB.CommandButton cmddtgg 
            BackColor       =   &H00C0C0C0&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4320
            Picture         =   "frmCinRepScMonRevOriginal.frx":0CC6
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   3240
            Width           =   1095
         End
         Begin VB.ListBox lstSubMedia 
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
            ForeColor       =   &H00404040&
            Height          =   3660
            Left            =   600
            MultiSelect     =   1  'Simple
            TabIndex        =   14
            Top             =   2040
            Width           =   2895
         End
         Begin VB.ListBox lstSubMediasel 
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
            ForeColor       =   &H00404040&
            Height          =   3660
            Left            =   6240
            MultiSelect     =   1  'Simple
            TabIndex        =   13
            Top             =   2040
            Width           =   3015
         End
         Begin VB.CommandButton cmddtg 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Selected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4320
            Picture         =   "frmCinRepScMonRevOriginal.frx":1108
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton CmdDtll 
            BackColor       =   &H00C0C0C0&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4320
            Picture         =   "frmCinRepScMonRevOriginal.frx":154A
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   4920
            Width           =   1095
         End
         Begin VB.CommandButton cmddtl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Selected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4320
            Picture         =   "frmCinRepScMonRevOriginal.frx":198C
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   4080
            Width           =   1095
         End
         Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
            Height          =   255
            Left            =   2880
            TabIndex        =   0
            Top             =   600
            Width           =   1335
            _Version        =   65541
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
         Begin PVMaskEditLib.PVMaskEdit txtdateto 
            Height          =   255
            Left            =   6840
            TabIndex        =   1
            Top             =   600
            Width           =   1335
            _Version        =   65541
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
         Begin VB.Line Line2 
            BorderColor     =   &H008080FF&
            BorderWidth     =   2
            X1              =   0
            X2              =   10680
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Date From"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   255
            Left            =   1440
            TabIndex        =   5
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Date To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   255
            Left            =   5640
            TabIndex        =   4
            Top             =   600
            Width           =   1095
         End
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1320
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
End
Attribute VB_Name = "frmCinRepScMonRevOriginal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim i
Dim f
Dim Z
Dim g
Dim FD As Date
Dim GD As Date
Dim TODT As Date
Dim FMDT As Date
Dim DT
Dim sum As Integer
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Private Sub CmdDtBack_Click()
 Unload Me
End Sub
Private Sub CmdDtClear_Click()
    txtdatefrom.TextWithMask = Now()
    txtdateto.TextWithMask = Now()
    lstSubMedia.Text = ""
    lstSubMediasel.Text = ""
    Populatesubmedia
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub
Private Sub Form_Load()
Dim X
txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
txtdateto.TextWithMask = Format(Now, "dd/mm/yyyy")
lstSubMedia.Clear
lstSubMediasel.Clear
Populatesubmedia
End Sub
Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtdateto.SetFocus
End Sub
Private Sub txtdatefrom_LostFocus()
If IsDate(txtdatefrom.TextWithMask) = False Then
   MsgBox "Invalid Date from", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub
Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then lstSubMedia.SetFocus
End Sub
Private Sub CmdDtprint_Click()

    Dim i
    Dim a, B, C
    
    TODT = Now()
    FMDT = Now()
 
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = "Delete * from dumbo_tracin"
   ws.BeginTrans
   db.Execute (Sqlqry)
   ws.CommitTrans
                     
   
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = "Delete * from Dumbo_traScreen"
   ws.BeginTrans
   db.Execute (Sqlqry)
   ws.CommitTrans
                    
        f = lstSubMediasel.ListIndex
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        For f = 0 To lstSubMediasel.ListCount - 1
        '  Invoice_Date Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "# order by media"
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry = "Select * from Bo_tracin where sub_media='" & Trim(lstSubMediasel.List(f)) & "'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
             If rs.RecordCount <> 0 Then
              rs.MoveFirst
              Do Until rs.EOF
              
                 Sqlqry1 = " Insert into dumbo_tracin values('" & rs!serial_no & "','" & rs!Year & "','" _
                                     & Trim(rs!Month) & "','" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                     & Trim(rs!SUB_MEDIA) & "','" _
                                     & Trim(rs!DATEFROM) & "','" _
                                     & Trim(rs!Dateto) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                     & Trim(rs!Day) & "','" _
                                     & Trim(rs!Length) & "','" _
                                     & findfirstfixup(Trim(rs!Description)) & "','" _
                                     & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                     & Trim(rs!Type) & "','" _
                                     & Trim(rs!tcurrency) & "'," _
                                     & Trim(rs!tconvertion) & "," _
                                     & Trim(rs!tra_amount) & "," _
                                     & Trim(rs!Amount) & ")"
           
                 
                     ws.BeginTrans
                     db.Execute (Sqlqry1)
                     ws.CommitTrans
                
                  rs.MoveNext
                Loop
              End If
          
        Next
        
        PROCMON
        
        
         
                          
End Sub



Private Sub Cmddtg_Click()
 For i = lstSubMedia.ListCount - 1 To 0 Step -1
    If lstSubMedia.Selected(i) Then
       lstSubMediasel.AddItem lstSubMedia.List(i)
       lstSubMedia.RemoveItem (i)
    End If
 Next
End Sub
Private Sub Cmddtgg_Click()
  For i = lstSubMedia.ListCount - 1 To 0 Step -1
         lstSubMediasel.AddItem lstSubMedia.List(i)
         lstSubMedia.RemoveItem (i)
  Next i

End Sub
Private Sub Cmddtl_Click()
 For f = lstSubMediasel.ListCount - 1 To 0 Step -1
    
    If lstSubMediasel.Selected(f) Then
       lstSubMedia.AddItem lstSubMediasel.Text
       lstSubMediasel.RemoveItem (f)
    End If
 Next
End Sub
Private Sub Cmddtll_Click()
 For i = lstSubMediasel.ListCount - 1 To 0 Step -1
         lstSubMedia.AddItem lstSubMediasel.List(i)
         lstSubMediasel.RemoveItem (i)
 Next i
End Sub
Private Sub LstSubmedia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lstSubMediasel.SetFocus
End Sub
Private Sub LstSubmediasel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmddtprint.SetFocus
End Sub
Private Sub PROCMON()
Dim FD As Date
Dim Ed As Date
Dim Hd As Date
Dim GD As Date
Dim DT As Date
Dim TODT As Date
Dim FMDT As Date
Dim m, n
 
 n = 0
 m = 0
   
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = "Delete * from dumbo_trascreen"
   ws.BeginTrans
   db.Execute (Sqlqry)
   ws.CommitTrans
    
    
    n = DateDiff("D", FD, GD)
     
         FD = DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy"))
         Hd = DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy"))
         GD = DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy"))
           
           
               'Sqlqry = "select * from dumbo_tracin where fd Between #" & Format(Datefrom, "mm-dd-yyyy") & "# and #" & Format(Dateto, "mm-dd-yyyy") & "# "
               Set ws = DBEngine.Workspaces(0)
               Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
               Sqlqry2 = "select * from dumbo_tracin "
               Set rs = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                   
                     If rs.RecordCount <> 0 Then
                      rs.MoveFirst
                          
                      Do Until rs.EOF
                       ' rs.MovePrevious
                             m = 0
                             n = 0
                             
                            Ed = Trim(rs!DATEFROM)
                            FMDT = Trim(rs!DATEFROM)
                            'FMDT = DateValue(rs!Datefrom)
                            TODT = Trim(rs!Dateto)
                           ' TODT = DateValue(rs!Dateto)
                            
                            If TODT < Hd Or FMDT > GD Then
                             m = -1
                            ElseIf Hd >= FMDT And GD <= TODT Then
                              m = DateDiff("d", Hd, GD)
                              Ed = Hd
                            ElseIf Hd >= FMDT And GD >= TODT Then
                              m = DateDiff("d", Hd, TODT)
                              Ed = Hd
                            ElseIf Hd <= FMDT And GD >= TODT Then
                              m = DateDiff("d", FMDT, TODT)
                              Ed = FMDT
                             ElseIf Hd <= FMDT And GD <= TODT Then
                              m = DateDiff("d", FMDT, GD)
                              Ed = FMDT
                              
                           ' elseif hd between fmdt and todt then
                              
                            End If
                            
                            n = 0
                            
                          For n = 0 To m
                                      Set ws = DBEngine.Workspaces(0)
                                      Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                                      Sqlqry1 = " Insert into dumbo_trascreen values('" & rs!serial_no & "','" & Trim(Ed) & "','" & rs!Year & "','" _
                                                          & Trim(rs!Month) & "','" _
                                                          & findfirstfixup(rs!Product) & "','" _
                                                          & findfirstfixup(rs!client) & "','" _
                                                          & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                                          & Trim(rs!SUB_MEDIA) & "','" _
                                                          & Trim(rs!DATEFROM) & "','" _
                                                          & Trim(rs!Dateto) & "','" _
                                                          & findfirstfixup(Trim(rs!bo_ref)) & "','" _
                                                          & Trim(rs!Day) & "'," _
                                                          & Val(rs!Length) & ",'" _
                                                          & findfirstfixup(Trim(rs!Description)) & "','" _
                                                          & findfirstfixup(Trim(rs!mat_code)) & "','" _
                                                          & Trim(rs!Type) & "','" _
                                                          & Trim(rs!tcurrency) & "'," _
                                                          & Trim(rs!tconvertion) & "," _
                                                          & Trim(rs!tra_amount) & "," _
                                                          & Trim(rs!NET_Amount) & ") "
                                
                                      
                                          ws.BeginTrans
                                          db.Execute (Sqlqry1)
                                          ws.CommitTrans
                                    
                                   '  rs.MoveNext
                                   '  Loop
                            '   End If
                            Ed = Ed + 1
                            Next
                            
                          rs.MoveNext
                          DT = FD
                          Ed = FD
                         Loop
                       End If
Populatesubmedia
With CrystalReport1
  .DataFiles(0) = App.Path & "\misov.mdb"
  .ReportFileName = App.Path & "\cinrepscreenworegion.rpt"
  .Formulas(0) = "zzz='" & " From: " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
  .WindowState = crptMaximized
  .Action = 1
End With
End Sub

Private Sub Populatesubmedia()
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = "Select * from cinema_rates"
      Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
      If rs.RecordCount <> 0 Then
      lstSubMedia.Clear
      lstSubMediasel.Clear
      rs.MoveFirst
      Do Until rs.EOF
         lstSubMedia.AddItem rs!SUB_MEDIA
         rs.MoveNext
       Loop
     End If
 
        
End Sub

Private Sub txtdateto_LostFocus()
If IsDate(txtdateto.TextWithMask) = False Then
   MsgBox "Invalid Date to", vbInformation, "Invalid Entry"
   txtdateto.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub
