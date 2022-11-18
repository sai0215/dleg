VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#1.5#0"; "pvmask.ocx"
Begin VB.Form frmCinRepscreen 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Cinema"
   ClientHeight    =   8775
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame FraMain 
      BackColor       =   &H00FFFFC0&
      Height          =   5175
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   10935
      Begin VB.Frame Fradatesel 
         BackColor       =   &H00FFFFC0&
         Height          =   1095
         Left            =   2160
         TabIndex        =   6
         Top             =   3480
         Width           =   6735
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
            Left            =   1080
            Picture         =   "frmCinRepscreen.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1575
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
            Left            =   4200
            Picture         =   "frmCinRepscreen.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1695
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
            Left            =   2640
            Picture         =   "frmCinRepscreen.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Fradate 
         BackColor       =   &H00FFFFC0&
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   10695
         Begin VB.ComboBox CboRegion 
            BackColor       =   &H80000014&
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
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1320
            Width           =   5415
         End
         Begin PVMaskEditLib.PVMaskEdit txtdatefrom 
            Height          =   255
            Left            =   4200
            TabIndex        =   0
            Top             =   360
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
            Left            =   6720
            TabIndex        =   1
            Top             =   360
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   " Region"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   1320
            TabIndex        =   11
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Line Line2 
            BorderColor     =   &H008080FF&
            X1              =   0
            X2              =   10680
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2880
            TabIndex        =   5
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   5640
            TabIndex        =   4
            Top             =   360
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
Attribute VB_Name = "frmCinRepscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ws As Workspace
Dim db As Database
Dim i
Dim f As Date
Dim g As Date
Dim Z
Dim sum As Integer
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String

Private Sub CboRegion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmddtprint.SetFocus
End Sub

Private Sub CmdDtBack_Click()
 Unload Me
End Sub

Private Sub CmdDtClear_Click()
    txtdatefrom.TextWithMask = Now()
    txtdateto.TextWithMask = Now()
    CboRegion.Text = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub
Private Sub Form_Load()
    Dim X
    txtdatefrom.TextWithMask = Format(Now, "dd/mm/yyyy")
    txtdateto.TextWithMask = Format(Now, "dd/mm/yyyy")
    Populateregion
    
End Sub
Private Sub txtdatefrom_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then txtdateto.SetFocus
End Sub

Private Sub txtdatefrom_LostFocus()
If IsDate(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy")) = False Then
   MsgBox "Invalid Date from", vbInformation, "Invalid Entry"
   txtdatefrom.SetFocus
   SendKeys " {Home} + {End} "
End If
End Sub

Private Sub txtdateto_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboRegion.SetFocus
End Sub

' New Entry

Private Sub CmdDtprint_Click()
    Dim i
    Dim a, B, C
    Dim dt

   If CboRegion.Text = "" Then
    MsgBox "Invalid Region", vbInformation, "Invalid Entry"
    CboRegion.SetFocus
    SendKeys " {Home} + {End} "
    Exit Sub
   End If
   
   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = "Delete * from dumbo_tracin"
   ws.BeginTrans
   db.Execute (Sqlqry)
   ws.CommitTrans
                
   
     
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
        Sqlqry1 = "Select * from bo_mas where region='" & Trim(CboRegion) & "' and Media='Cinema"
        Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        If rs1.RecordCount <> 0 Then
        rs1.MoveFirst
        Do Until rs1.EOF
              
               f = DateValue(Format(txtdatefrom.TextWithMask, "dd/mm/yyyy"))
               g = DateValue(Format(txtdateto.TextWithMask, "dd/mm/yyyy"))
               For dt = f To g
                Sqlqry = "Select * from Bo_tracin where datefrom >=#" & f & "#  and  dateto<=#" & f & "# and serial_no='" & Trim(rs1!serial_no) & "'"
                     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     If rs.RecordCount <> 0 Then
                      rs.MoveFirst
                      Do Until rs.EOF
                         Set ws = DBEngine.Workspaces(0)
                         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         Sqlqry1 = " Insert into dumbo_trascreen values('" & rs!serial_no & "',#" & f & "#,'" & rs!Year & "','" _
                                             & Trim(rs!Month) & "','" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & Trim(rs!media) & "','" _
                                             & Trim(rs!sub_media) & "','" _
                                             & Trim(rs!datefrom) & "','" _
                                             & Trim(rs!dateto) & "','" _
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
                    f = f + 1
                Next
        
             rs1.MoveNext
            Loop
           End If
        
 With CrystalReport1
  .DataFiles(0) = App.Path & "\misov.mdb"
  .ReportFileName = App.Path & "\cinrepscreen.rpt"
  .Formulas(0) = "zzz='" & " Date From: " & Trim(txtdatefrom.TextWithMask) & " To " & Trim(txtdateto.TextWithMask) & "'"
  .Formulas(1) = "yyy='" & Trim(CboRegion) & "'"
  .WindowState = crptMaximized
  .Action = 1
 End With
                          
End Sub

Private Sub Populateregion()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select distinct(region) from cinema_rates Order by region "
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
      rs.MoveFirst
        CboRegion.Clear
        CboRegion.AddItem "All"
        Do Until rs.EOF
            CboRegion.AddItem rs!region
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
