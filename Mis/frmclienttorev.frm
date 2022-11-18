VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmClientTorev 
   BackColor       =   &H80000005&
   Caption         =   "ClientTO"
   ClientHeight    =   8550
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmclienttorev.frx":0000
      Left            =   960
      Top             =   7800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "                                       Turnover / Client                                    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   8415
      Left            =   360
      TabIndex        =   10
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
         Caption         =   "Sort"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   22
         Top             =   6480
         Width           =   10335
         Begin VB.OptionButton OptSubMedia 
            BackColor       =   &H80000009&
            Caption         =   "Sub Media"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   840
            TabIndex        =   27
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton OptProduct 
            BackColor       =   &H80000009&
            Caption         =   "Product"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   4560
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptClient 
            BackColor       =   &H80000009&
            Caption         =   "Client"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   2760
            TabIndex        =   25
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptRegion 
            BackColor       =   &H80000009&
            Caption         =   "Region"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   8280
            TabIndex        =   24
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptMonth 
            BackColor       =   &H80000009&
            Caption         =   "Month"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   6360
            TabIndex        =   23
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.ComboBox cboregion 
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
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   4965
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3600
         Width           =   2295
      End
      Begin VB.ComboBox cboCurrency 
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
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   4965
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2880
         Width           =   2295
      End
      Begin VB.ComboBox cboProduct 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   390
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   5280
         Width           =   5655
      End
      Begin VB.ComboBox cboMediaType 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   390
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   4440
         Width           =   5655
      End
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00C0C0C0&
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
         Left            =   3840
         Picture         =   "frmclienttorev.frx":001D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7560
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00C0C0C0&
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
         Left            =   5040
         Picture         =   "frmclienttorev.frx":045F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7560
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00C0C0C0&
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
         Height          =   735
         Left            =   6240
         Picture         =   "frmclienttorev.frx":08A1
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7560
         Width           =   1335
      End
      Begin VB.ComboBox CboClient 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   390
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   6000
         Width           =   5655
      End
      Begin VB.ComboBox cbomonthTo 
         BackColor       =   &H80000018&
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
         Height          =   420
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cbomonthfrom 
         BackColor       =   &H80000018&
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
         Height          =   420
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ComboBox cboyear 
         BackColor       =   &H80000018&
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
         Height          =   420
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Region"
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
         Left            =   3480
         TabIndex        =   21
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblSubMediaName 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   6480
         TabIndex        =   18
         Top             =   4440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblMedianame 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   4440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Product"
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
         Left            =   2280
         TabIndex        =   16
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Media Type"
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
         Left            =   1680
         TabIndex        =   15
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   11280
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Client"
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
         Left            =   2400
         TabIndex        =   14
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Month To"
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
         Left            =   3600
         TabIndex        =   13
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Month From"
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
         Left            =   3240
         TabIndex        =   12
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "  Year"
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
         Left            =   3840
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmClientTorev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ws As Workspace
Dim db As Database
Dim i As Integer
Dim d, e, f, g As Integer
Dim C, X, Y, Z As Integer
Dim adddisc As Currency
Dim scharge As Currency
Dim ntra As Currency
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim crdtamt As Currency
Dim Addiscamt As Currency
Dim totaddiscamt As Currency
Public n, m
Private Sub CboCurrency_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboregion.SetFocus
End Sub
Private Sub cboMediaType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboProduct.SetFocus
End Sub
Private Sub cboMediatype_LostFocus()
 LblMediaName.Caption = ""
 lblSubMediaName.Caption = ""
 
If Mid(cboMediaType.Text, 1, 3) = "Cin" Then
   LblMediaName.Caption = "Cinema"
   lblSubMediaName.Caption = Trim(Mid(cboMediaType, 8, 30))
ElseIf Mid(cboMediaType.Text, 1, 3) = "Mag" Then
   LblMediaName.Caption = "Magazine"
   lblSubMediaName.Caption = Trim(Mid(cboMediaType, 10, 30))
ElseIf Mid(cboMediaType.Text, 1, 3) = "Onl" Then
   LblMediaName.Caption = "Online"
   lblSubMediaName.Caption = Trim(Mid(cboMediaType, 8, 30))
ElseIf Mid(cboMediaType.Text, 1, 3) = "Tel" Then
   LblMediaName.Caption = "Television"
   lblSubMediaName.Caption = Trim(Mid(cboMediaType, 12, 30))
End If
CboProduct.SetFocus
End Sub

Private Sub cbomonthfrom_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbomonthTo.SetFocus
End Sub

Private Sub cbomonthTo_GotFocus()
cbomonthTo.Clear
If cbomonthfrom.ListIndex = 0 Then
    cbomonthTo.AddItem "January"
    cbomonthTo.AddItem "February"
    cbomonthTo.AddItem "March"
    cbomonthTo.AddItem "April"
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 1 Then
    cbomonthTo.AddItem "February"
    cbomonthTo.AddItem "March"
    cbomonthTo.AddItem "April"
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 2 Then
    cbomonthTo.AddItem "March"
    cbomonthTo.AddItem "April"
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 3 Then
    cbomonthTo.AddItem "April"
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 4 Then
    cbomonthTo.AddItem "May"
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 5 Then
    cbomonthTo.AddItem "June"
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 6 Then
    cbomonthTo.AddItem "July"
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 7 Then
    cbomonthTo.AddItem "August"
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 8 Then
    cbomonthTo.AddItem "September"
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 9 Then
    cbomonthTo.AddItem "October"
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 10 Then
    cbomonthTo.AddItem "November"
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
ElseIf cbomonthfrom.ListIndex = 11 Then
    cbomonthTo.AddItem "December"
    cbomonthTo.SetFocus
Else
    cbomonthTo.SetFocus
End If
End Sub
Private Sub cbomonthTo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CboCurrency.SetFocus
End Sub
Private Sub cboProduct_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboClient.SetFocus
End Sub


Private Sub CboRegion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboMediaType.SetFocus
End Sub

Private Sub cboyear_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cbomonthfrom.SetFocus
End Sub
Private Sub cmdBack_Click()
 Unload Me
End Sub
Private Sub cmdClear_Click()
 textclear
End Sub
Private Sub cmdDisplay_Click()
    Dim l, o, p As String
    Dim uname As String
    Dim compname As String
    Dim objnet
    Dim fmname
    Dim fmid
    Dim temp


  
   If ValidateData = True Then
              
  On Error GoTo xyz
  

fmname = Me.Caption
fmid = Me.Name

temp = False
On Error Resume Next

Set objnet = CreateObject("WScript.Network")

If Err.number <> 0 Then
  MsgBox "Error in Getting computer name." & vbCrLf & _
   """No""If your browser warns you."
End If

uname = ""
compname = ""
uname = objnet.UserName
compname = objnet.computername

Set objnet = Nothing

   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = " Select * from formcontrol1 where form_caption='" & Trim(fmname) & "'"
   Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
  ' MsgBox Sqlqry
   If rs.RecordCount = 0 Then
     MsgBox "Form Caption is not matching"
     Exit Sub
   Else
    rs.MoveFirst
'     fmid = ""
     fmid = rs!form_name
     
    If rs!lock_status = "Y" Then
            uname = rs!u_name
            MsgBox "Table has been locked exclusively by the user." & uname
            Exit Sub
        
            cmdDisplay.SetFocus
    
    
    Else
       Set ws = DBEngine.Workspaces(0)
       Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
       Sqlqry1 = "Update formcontrol1 set " _
                 & " U_Name='" & uname & "'," _
                 & " Comp_Name='" & compname & "'," _
                 & " Lock_status='Y' where form_caption='" & fmname & "'"
       ' MsgBox Sqlqry1
        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
    
    
                'cmdModify.Enabled = Fal
     End If
   End If


   
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = " Delete * from To_Client"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
         n = ""
         m = ""
         o = ""
         p = ""
         l = ""
         crdtamt = 0
         
         
       n = LblMediaName.Caption
       m = lblSubMediaName.Caption
       
       If CboClient.Text <> "All" Then o = CboClient.Text
       If CboProduct.Text <> "All" Then p = CboProduct.Text
       If cboregion.Text <> "All" Then l = cboregion.Text
             
     If cboregion.Text = "All" Then
       prev1
    Else
     If CboClient.Text = "All" And CboProduct.Text = "All" And cboMediaType.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & ""
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
                
        ElseIf CboClient.Text = "All" And CboProduct.Text = "All" And cboMediaType.Text = "Cinema" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' and region='" & Trim(cboregion.Text) & "' and CANCELL='N' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
         
         ElseIf CboClient.Text = "All" And CboProduct.Text = "All" And cboMediaType.Text = "Magazine" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and media = '" & Trim(cboMediaType.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                        crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
            ElseIf CboClient.Text = "All" And CboProduct.Text = "All" And cboMediaType.Text = "Online" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  
                  Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
            ElseIf CboClient.Text = "All" And CboProduct.Text = "All" And cboMediaType.Text = "Television" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
               ElseIf CboClient.Text = "All" And CboProduct.Text = "All" And cboMediaType.Text = n & " " & m Then
                
                If Mid(n, 1, 3) <> "Cin" Then
                    Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(LblMediaName.Caption) & "' and sub_media='" & Trim(lblSubMediaName.Caption) & " ' "
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                    If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                      Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                         Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                 & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                 & rs!sub_media & "'," _
                                                 & rs!monthind & "," _
                                                 & Val(rs!tra_gamount) & "," _
                                                 & Val(rs!Tot_free) & "," _
                                                 & Val(rs!Tot_barter) & "," _
                                                 & Val(rs!disc_percentage) & "," _
                                                 & Val(rs!surcharge) & ",'" _
                                                 & Trim(rs!disc_rate) & "'," _
                                                 & Val(rs!add_discount) + crdtamt & "," _
                                                 & Val(rs!tra_namount) - crdtamt & ")"
                             ws.BeginTrans
                             db.Execute (Sqlqry)
                             ws.CommitTrans
                         rs.MoveNext
                
                      Loop
                     End If
                 Else
                  prevcin
                 End If
           
                ElseIf CboClient.Text = "All" And CboProduct.Text = p And cboMediaType.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' AND Product='" & Trim(CboProduct.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & ""
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                ElseIf CboClient.Text = "All" And CboProduct.Text = p And cboMediaType.Text = "Cinema" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and product='" & Trim(CboProduct.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Media='" & Trim(cboMediaType.Text) & "' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = "All" And CboProduct.Text = p And cboMediaType.Text = "Magazine" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and product='" & Trim(CboProduct.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Media='" & Trim(cboMediaType.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
            
                ElseIf CboClient.Text = "All" And CboProduct.Text = p And cboMediaType.Text = "Television" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and product='" & Trim(CboProduct.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Media='" & Trim(cboMediaType.Text) & "' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = "All" And CboProduct.Text = p And cboMediaType.Text = "Online" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and product='" & Trim(CboProduct.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Media='" & Trim(cboMediaType.Text) & "' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                
                ElseIf CboClient.Text = "All" And CboProduct.Text = p And cboMediaType.Text = n & " " & m Then
                
                If Mid(n, 1, 3) <> "Cin" Then
                    Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & n & "' and sub_media='" & m & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Product ='" & Trim(CboProduct.Text) & "'"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                    If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                      Do Until rs.EOF
                        crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                         Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                 & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                 & rs!sub_media & "'," _
                                                 & rs!monthind & "," _
                                                 & Val(rs!tra_gamount) & "," _
                                                 & Val(rs!Tot_free) & "," _
                                                 & Val(rs!Tot_barter) & "," _
                                                 & Val(rs!disc_percentage) & "," _
                                                 & Val(rs!surcharge) & ",'" _
                                                 & Trim(rs!disc_rate) & "'," _
                                                 & Val(rs!add_discount) + crdtamt & "," _
                                                 & Val(rs!tra_namount) - crdtamt & ")"
                             ws.BeginTrans
                             db.Execute (Sqlqry)
                             ws.CommitTrans
                         rs.MoveNext
                
                      Loop
                     End If
                 Else
                   prevcin
                 End If
    ' modify  1
                ElseIf CboClient.Text = o And CboProduct.Text = p And cboMediaType.Text = n & " " & m Then
                If Mid(n, 1, 3) <> "Cin" Then
                 Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & n & "' and sub_media='" & m & "' and Client='" & Trim(CboClient.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Product ='" & Trim(CboProduct.Text) & "'"
                 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                    If rs.RecordCount <> 0 Then
                      rs.MoveFirst
                       Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                      
                         Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                 & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                 & rs!sub_media & "'," _
                                                 & rs!monthind & "," _
                                                 & Val(rs!tra_gamount) & "," _
                                                 & Val(rs!Tot_free) & "," _
                                                 & Val(rs!Tot_barter) & "," _
                                                 & Val(rs!disc_percentage) & "," _
                                                 & Val(rs!surcharge) & ",'" _
                                                 & Trim(rs!disc_rate) & "'," _
                                                 & Val(rs!add_discount) + crdtamt & "," _
                                                 & Val(rs!tra_namount) - crdtamt & ")"
                             ws.BeginTrans
                             db.Execute (Sqlqry)
                             ws.CommitTrans
                         rs.MoveNext
                         Loop
                     End If
                 Else
                 
                  prevcin
                 End If
 
               ElseIf CboClient.Text = o And CboProduct.Text = "All" And cboMediaType.Text = n & " " & m Then
                If Mid(n, 1, 3) <> "Cin" Then
                    Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(LblMediaName.Caption) & "' and sub_media='" & Trim(lblSubMediaName.Caption) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Client='" & Trim(CboClient.Text) & "'"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                    If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                      Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                         Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                 & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                 & rs!sub_media & "'," _
                                                 & rs!monthind & "," _
                                                 & Val(rs!tra_gamount) & "," _
                                                 & Val(rs!Tot_free) & "," _
                                                 & Val(rs!Tot_barter) & "," _
                                                 & Val(rs!disc_percentage) & "," _
                                                 & Val(rs!surcharge) & ",'" _
                                                 & Trim(rs!disc_rate) & "'," _
                                                 & Val(rs!add_discount) + crdtamt & "," _
                                                 & Val(rs!tra_namount) - crdtamt & ")"
                             ws.BeginTrans
                             db.Execute (Sqlqry)
                             ws.CommitTrans
                         rs.MoveNext
                
                      Loop
                     End If
                  Else
                   prevcin
                End If
                ElseIf CboClient.Text = o And CboProduct.Text = "All" And cboMediaType.Text = "Cinema" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Client='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                     Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                     Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                     If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                ElseIf CboClient.Text = o And CboProduct.Text = "All" And cboMediaType.Text = "Magazine" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Client='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                     Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = "All" And cboMediaType.Text = "Online" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Client='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                   crdtamt = 0
                  Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = "All" And cboMediaType.Text = "Television" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Client='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                  crdtamt = 0
                  Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = "All" And cboMediaType.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Client='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                   crdtamt = 0
                  Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = p And cboMediaType.Text = "Cinema" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' and Client='" & Trim(CboClient.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Product ='" & Trim(CboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                    crdtamt = 0
                  Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = p And cboMediaType.Text = "Magazine" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' and Client='" & Trim(CboClient.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Product ='" & Trim(CboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                    crdtamt = 0
                  Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
        End If
      End If
            
            continuation2
            Cinadjustments
            curadjustments
            continuation1
          
          checkout
      
    Exit Sub
            
   Else
        MsgBox "Improper Dates", vbDefaultButton1, "Invalid entry"
        Exit Sub
  End If
  
xyz:
  MsgBox " Table Has been Locked by Other User, Wait few Seconds and process your request"
  
 End Sub
Private Sub checkout()

Dim uname As String
Dim compname As String
Dim objnet
Dim fmname
Dim fmid

On Error Resume Next


fmname = Me.Caption
fmid = Me.Name

Set objnet = CreateObject("WScript.Network")

If Err.number <> 0 Then
  MsgBox "Error in Getting computer name." & vbCrLf & _
   "Do not Press""No""If your browser warns you."
End If

uname = ""
compname = ""
uname = objnet.UserName
compname = objnet.computername

Set objnet = Nothing

   Set ws = DBEngine.Workspaces(0)
   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
   Sqlqry = " Select * from formcontrol1 where form_caption='" & fmname & "' and u_name='" & uname & "'"
   Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
  ' MsgBox Sqlqry
   If rs.RecordCount <> 0 Then
     rs.MoveFirst
 
     fmid = rs!form_name
        Sqlqry1 = "Update formcontrol1 set " _
                 & " U_Name=''," _
                 & " Comp_Name=''," _
                 & " Lock_status='N' where form_caption='" & fmname & "'"
        ws.BeginTrans
        db.Execute (Sqlqry1)
        ws.CommitTrans
   End If
 
End Sub

Private Sub continuation2()
Dim l, o, p As String
  
         n = ""
         m = ""
         o = ""
         p = ""
         l = ""
         crdtamt = 0
         
         
       n = LblMediaName.Caption
       m = lblSubMediaName.Caption
       
       If CboClient.Text <> "All" Then o = CboClient.Text
       If CboProduct.Text <> "All" Then p = CboProduct.Text
       If cboregion.Text <> "All" Then l = cboregion.Text
             

 If CboClient.Text = o And CboProduct.Text = p And cboMediaType.Text = "Online" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' and Client='" & Trim(CboClient.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Product ='" & Trim(CboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                   crdtamt = 0
                  Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = p And cboMediaType.Text = "Television" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Media='" & Trim(cboMediaType.Text) & "' and Client='" & Trim(CboClient.Text) & "' and Product='" & Trim(CboProduct.Text) & "' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                    crdtamt = 0
                   Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = p And cboMediaType.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Client='" & Trim(CboClient.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Product ='" & Trim(CboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
      End If
End Sub

Private Sub continuation1()
    
    If Mid(cboMediaType.Text, 1, 3) = "Mag" Then
         If OptSubMedia.Value = True Then
            With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Toclientmagsm.rpt"
                .Formulas(0) = "yyy='" & Val(Cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "mty='" & Trim(cboMediaType.Text) & "'"
                .Formulas(3) = "prd='" & Trim(CboProduct.Text) & "'"
                .Formulas(4) = "agn='" & Trim(CboClient.Text) & "'"
                .Formulas(5) = "Cur='" & Trim(CboCurrency.Text) & "'"
                .Formulas(6) = "reg='" & Trim(cboregion.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
          ElseIf OptClient.Value = True Then
            With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Toclientmagcl.rpt"
                .Formulas(0) = "yyy='" & Val(Cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "mty='" & Trim(cboMediaType.Text) & "'"
                .Formulas(3) = "prd='" & Trim(CboProduct.Text) & "'"
                .Formulas(4) = "agn='" & Trim(CboClient.Text) & "'"
                .Formulas(5) = "Cur='" & Trim(CboCurrency.Text) & "'"
                .Formulas(6) = "reg='" & Trim(cboregion.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
          ElseIf OptProduct.Value = True Then
            With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Toclientmagpr.rpt"
                .Formulas(0) = "yyy='" & Val(Cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "mty='" & Trim(cboMediaType.Text) & "'"
                .Formulas(3) = "prd='" & Trim(CboProduct.Text) & "'"
                .Formulas(4) = "agn='" & Trim(CboClient.Text) & "'"
                .Formulas(5) = "Cur='" & Trim(CboCurrency.Text) & "'"
                .Formulas(6) = "reg='" & Trim(cboregion.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
            
          ElseIf OptMonth.Value = True Then
            With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Toclientmagmo.rpt"
                .Formulas(0) = "yyy='" & Val(Cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "mty='" & Trim(cboMediaType.Text) & "'"
                .Formulas(3) = "prd='" & Trim(CboProduct.Text) & "'"
                .Formulas(4) = "agn='" & Trim(CboClient.Text) & "'"
                .Formulas(5) = "Cur='" & Trim(CboCurrency.Text) & "'"
                .Formulas(6) = "reg='" & Trim(cboregion.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
           End With
           
          ElseIf OptRegion.Value = True Then
            With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Toclientmagre.rpt"
                .Formulas(0) = "yyy='" & Val(Cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "mty='" & Trim(cboMediaType.Text) & "'"
                .Formulas(3) = "prd='" & Trim(CboProduct.Text) & "'"
                .Formulas(4) = "agn='" & Trim(CboClient.Text) & "'"
                .Formulas(5) = "Cur='" & Trim(CboCurrency.Text) & "'"
                .Formulas(6) = "reg='" & Trim(cboregion.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
           End With
           
         End If
         
       Else
       
        If OptSubMedia.Value = True Then
              With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Toclientsm.rpt"
                .Formulas(0) = "yyy='" & Val(Cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "mty='" & Trim(cboMediaType.Text) & "'"
                .Formulas(3) = "prd='" & Trim(CboProduct.Text) & "'"
                .Formulas(4) = "agn='" & Trim(CboClient.Text) & "'"
                .Formulas(5) = "Cur='" & Trim(CboCurrency.Text) & "'"
                .Formulas(6) = "reg='" & Trim(cboregion.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
         ElseIf OptClient.Value = True Then
              With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Toclientcl.rpt"
                .Formulas(0) = "yyy='" & Val(Cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "mty='" & Trim(cboMediaType.Text) & "'"
                .Formulas(3) = "prd='" & Trim(CboProduct.Text) & "'"
                .Formulas(4) = "agn='" & Trim(CboClient.Text) & "'"
                .Formulas(5) = "Cur='" & Trim(CboCurrency.Text) & "'"
                .Formulas(6) = "reg='" & Trim(cboregion.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
          ElseIf OptProduct.Value = True Then
              With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Toclientpr.rpt"
                .Formulas(0) = "yyy='" & Val(Cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "mty='" & Trim(cboMediaType.Text) & "'"
                .Formulas(3) = "prd='" & Trim(CboProduct.Text) & "'"
                .Formulas(4) = "agn='" & Trim(CboClient.Text) & "'"
                .Formulas(5) = "Cur='" & Trim(CboCurrency.Text) & "'"
                .Formulas(6) = "reg='" & Trim(cboregion.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
          ElseIf OptMonth.Value = True Then
              With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Toclientmo.rpt"
                .Formulas(0) = "yyy='" & Val(Cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "mty='" & Trim(cboMediaType.Text) & "'"
                .Formulas(3) = "prd='" & Trim(CboProduct.Text) & "'"
                .Formulas(4) = "agn='" & Trim(CboClient.Text) & "'"
                .Formulas(5) = "Cur='" & Trim(CboCurrency.Text) & "'"
                .Formulas(6) = "reg='" & Trim(cboregion.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
          ElseIf OptRegion.Value = True Then
              With CrystalReport1
                .DataFiles(0) = App.Path & "\misov.mdb"
                .ReportFileName = App.Path & "\Toclientre.rpt"
                .Formulas(0) = "yyy='" & Val(Cboyear.Text) & "'"
                .Formulas(1) = "mmm='" & " From " & Trim(cbomonthfrom.Text) & " To " & Trim(cbomonthTo.Text) & "'"
                .Formulas(2) = "mty='" & Trim(cboMediaType.Text) & "'"
                .Formulas(3) = "prd='" & Trim(CboProduct.Text) & "'"
                .Formulas(4) = "agn='" & Trim(CboClient.Text) & "'"
                .Formulas(5) = "Cur='" & Trim(CboCurrency.Text) & "'"
                .Formulas(6) = "reg='" & Trim(cboregion.Text) & "'"
                .WindowState = crptMaximized
                .Action = 1
            End With
        End If
            
      End If
End Sub
Private Sub prev1()

Dim l, o, p As String

        
         o = ""
         p = ""
         l = ""
         
         
       n = LblMediaName.Caption
       m = lblSubMediaName.Caption
       
       If CboClient.Text <> "All" Then o = CboClient.Text
       If CboProduct.Text <> "All" Then p = CboProduct.Text
       If cboregion.Text <> "All" Then l = cboregion.Text
             

       If CboClient.Text = "All" And CboProduct.Text = "All" And cboMediaType.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " AND CANCELL='N'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                   crdtamt = 0
                     Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
                
        ElseIf CboClient.Text = "All" And CboProduct.Text = "All" And cboMediaType.Text = "Cinema" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " AND CANCELL='N' and Media='" & Trim(cboMediaType.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
         ElseIf CboClient.Text = "All" And CboProduct.Text = "All" And cboMediaType.Text = "Magazine" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' and media = '" & Trim(cboMediaType.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " AND CANCELL='N' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                    crdtamt = 0
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
           ElseIf CboClient.Text = "All" And CboProduct.Text = "All" And cboMediaType.Text = "Online" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " AND CANCELL='N' and Media='" & Trim(cboMediaType.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                     Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                End If
                
            ElseIf CboClient.Text = "All" And CboProduct.Text = "All" And cboMediaType.Text = "Television" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " AND CANCELL='N' and Media='" & Trim(cboMediaType.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
               ElseIf CboClient.Text = "All" And CboProduct.Text = "All" And cboMediaType.Text = n & " " & m Then
                
                 If Mid(n, 1, 3) <> "Cin" Then
                    Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " AND CANCELL='N' and Media='" & Trim(LblMediaName.Caption) & "' and sub_media='" & Trim(lblSubMediaName.Caption) & " ' "
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                    If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                      Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                         Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                 & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                 & rs!sub_media & "'," _
                                                 & rs!monthind & "," _
                                                 & Val(rs!tra_gamount) & "," _
                                                 & Val(rs!Tot_free) & "," _
                                                 & Val(rs!Tot_barter) & "," _
                                                 & Val(rs!disc_percentage) & "," _
                                                 & Val(rs!surcharge) & ",'" _
                                                 & Trim(rs!disc_rate) & "'," _
                                                 & Val(rs!add_discount) + crdtamt & "," _
                                                 & Val(rs!tra_namount) - crdtamt & ")"
                             ws.BeginTrans
                             db.Execute (Sqlqry)
                             ws.CommitTrans
                         rs.MoveNext
                
                      Loop
                     End If
                  Else
                   prevcin
                 End If
                ElseIf CboClient.Text = "All" And CboProduct.Text = p And cboMediaType.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND Product='" & Trim(CboProduct.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " AND CANCELL='N'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                    crdtamt = 0
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount - crdtamt) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                ElseIf CboClient.Text = "All" And CboProduct.Text = p And cboMediaType.Text = "Cinema" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and product='" & Trim(CboProduct.Text) & "' AND CANCELL='N' and Media='" & Trim(cboMediaType.Text) & "' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = "All" And CboProduct.Text = p And cboMediaType.Text = "Magazine" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and product='" & Trim(CboProduct.Text) & "' AND CANCELL='N' and Media='" & Trim(cboMediaType.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                      crdtamt = 0
                        Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
            
                ElseIf CboClient.Text = "All" And CboProduct.Text = p And cboMediaType.Text = "Television" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and product='" & Trim(CboProduct.Text) & "' AND CANCELL='N' and Media='" & Trim(cboMediaType.Text) & "' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                rs.MoveFirst
                  Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = "All" And CboProduct.Text = p And cboMediaType.Text = "Online" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and product='" & Trim(CboProduct.Text) & "' AND CANCELL='N' and Media='" & Trim(cboMediaType.Text) & "' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                    crdtamt = 0
                     Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                
                ElseIf CboClient.Text = "All" And CboProduct.Text = p And cboMediaType.Text = n & " " & m Then
                 If Mid(n, 1, 3) <> "Cin" Then
                    Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & n & " ' and sub_media='" & m & "' AND CANCELL='N' and Product ='" & Trim(CboProduct.Text) & "'"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                    If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                      Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                         Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                 & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                 & rs!sub_media & "'," _
                                                 & rs!monthind & "," _
                                                 & Val(rs!tra_gamount) & "," _
                                                 & Val(rs!Tot_free) & "," _
                                                 & Val(rs!Tot_barter) & "," _
                                                 & Val(rs!disc_percentage) & "," _
                                                 & Val(rs!surcharge) & ",'" _
                                                 & Trim(rs!disc_rate) & "'," _
                                                 & Val(rs!add_discount) + crdtamt & "," _
                                                 & Val(rs!tra_namount) - crdtamt & ")"
                             ws.BeginTrans
                             db.Execute (Sqlqry)
                             ws.CommitTrans
                         rs.MoveNext
                
                      Loop
                     End If
                   Else
                    prevcin
                   End If
                     
                ElseIf CboClient.Text = o And CboProduct.Text = p And cboMediaType.Text = n & " " & m Then
                 If Mid(n, 1, 3) <> "Cin" Then
                    Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & n & "' and sub_media='" & m & "' AND CANCELL='N' and Client='" & Trim(CboClient.Text) & "' and Product ='" & Trim(CboProduct.Text) & "'"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                    If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                      Do Until rs.EOF
                          crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                         Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                 & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                 & rs!sub_media & "'," _
                                                 & rs!monthind & "," _
                                                 & Val(rs!tra_gamount) & "," _
                                                 & Val(rs!Tot_free) & "," _
                                                 & Val(rs!Tot_barter) & "," _
                                                 & Val(rs!disc_percentage) & "," _
                                                 & Val(rs!surcharge) & ",'" _
                                                 & Trim(rs!disc_rate) & "'," _
                                                 & Val(rs!add_discount) + crdtamt & "," _
                                                 & Val(rs!tra_namount) - crdtamt & ")"
                             ws.BeginTrans
                             db.Execute (Sqlqry)
                             ws.CommitTrans
                         rs.MoveNext
                
                      Loop
                     End If
                  Else
                    prevcin
                  End If
                     
                ElseIf CboClient.Text = o And CboProduct.Text = "All" And cboMediaType.Text = n & " " & m Then
                 If Mid(n, 1, 3) <> "Cin" Then
                    Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(LblMediaName.Caption) & "' AND CANCELL='N' and sub_media='" & Trim(lblSubMediaName.Caption) & "' and Client='" & Trim(CboClient.Text) & "'"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                    If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                      Do Until rs.EOF
                         crdtamt = 0
                        Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                         Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                 & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                 & rs!sub_media & "'," _
                                                 & rs!monthind & "," _
                                                 & Val(rs!tra_gamount) & "," _
                                                 & Val(rs!Tot_free) & "," _
                                                 & Val(rs!Tot_barter) & "," _
                                                 & Val(rs!disc_percentage) & "," _
                                                 & Val(rs!surcharge) & ",'" _
                                                 & Trim(rs!disc_rate) & "'," _
                                                 & Val(rs!add_discount) + crdtamt & "," _
                                                 & Val(rs!tra_namount) - crdtamt & ")"
                             ws.BeginTrans
                             db.Execute (Sqlqry)
                             ws.CommitTrans
                         rs.MoveNext
                
                      Loop
                     End If
                   Else
                     prevcin
                   End If
                     
                ElseIf CboClient.Text = o And CboProduct.Text = "All" And cboMediaType.Text = "Cinema" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' AND CANCELL='N' and Client='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                ElseIf CboClient.Text = o And CboProduct.Text = "All" And cboMediaType.Text = "Magazine" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' AND CANCELL='N' and Client='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                      crdtamt = 0
                     Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = "All" And cboMediaType.Text = "Online" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' AND CANCELL='N' and Client='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = "All" And cboMediaType.Text = "Television" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' AND CANCELL='N' and Client='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = "All" And cboMediaType.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " AND CANCELL='N' and Client='" & Trim(CboClient.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                    crdtamt = 0
                   Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = p And cboMediaType.Text = "Cinema" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' AND CANCELL='N' and Client='" & Trim(CboClient.Text) & "' and Product ='" & Trim(CboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = p And cboMediaType.Text = "Magazine" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' AND CANCELL='N' and Client='" & Trim(CboClient.Text) & "' and Product ='" & Trim(CboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                     Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = p And cboMediaType.Text = "Online" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' AND CANCELL='N' and Client='" & Trim(CboClient.Text) & "' and Product ='" & Trim(CboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                   Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = p And cboMediaType.Text = "Television" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & Trim(cboMediaType.Text) & "' AND CANCELL='N' and Client='" & Trim(CboClient.Text) & "' and Product='" & Trim(CboProduct.Text) & "' "
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                      crdtamt = 0
                    Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
                ElseIf CboClient.Text = o And CboProduct.Text = p And cboMediaType.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Client='" & Trim(CboClient.Text) & "' AND CANCELL='N' and Product ='" & Trim(CboProduct.Text) & "'"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                  Do Until rs.EOF
                     crdtamt = 0
                     Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                     
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!monthind & "," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) + crdtamt & "," _
                                             & Val(rs!tra_namount) - crdtamt & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                  Loop
                 End If
                 
              End If

End Sub
Private Sub Cinadjustments()
Dim dumregion As String

Sqlqry = " Delete * from To_Client1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
     
     If Len(cboMediaType) > 10 Then
              
        Sqlqry = "Select * from To_Client where Media='Cinema'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into To_Client1 values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!bo_ref & "," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!NET_Amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
           
     Sqlqry = " Delete * from To_Client where media='Cinema'"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
              
        adddisc = 0
        scharge = 0
        ntra = 0
                
                       
              
      Sqlqry = "Select * from To_Client1"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
              
                    
                Do Until rs.EOF
                 
                adddisc = 0
                scharge = 0
                ntra = 0
                
                 Sqlqry1 = "Select * from bo_tracin where serial_no='" & Trim(rs!serial_no) & "' AND TYPE ='Paid' and sub_media='" & m & "'"
                 Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                    If rs1.RecordCount = 0 Then
                        adddisc = rs!add_discount
                        scharge = rs!surcharge
                    Else
                        adddisc = rs!add_discount / rs1.RecordCount
                        scharge = rs!surcharge / rs1.RecordCount
                    End If
                
                
                
                
                Sqlqry1 = "Select * from bo_tracin where serial_no='" & Trim(rs!serial_no) & "' and sub_media='" & m & "'"
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                 If rs1.RecordCount <> 0 Then
                  rs1.MoveFirst
                   Do Until rs1.EOF
                     
                     Sqlqry2 = "Select region from cinema_rates where sub_media='" & rs1!sub_media & "' "
                     Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                     If IsNull(rs2.Fields(0)) = False Then dumregion = rs2.Fields(0)
                     
                     
                     If rs1!Type = "Paid" Then
                       
                            If rs!add_discount = 0 Then
                              ntra = Val(rs1!tra_amount) - (Val(rs1!tra_amount) * rs!disc_rate / 100) - (((rs1!tra_amount) - (rs1!tra_amount * rs!disc_rate / 100)) * rs!disc_percentage / 100)
                            Else
                              ntra = Val(rs1!tra_amount) - (Val(rs1!tra_amount) * rs!disc_rate / 100) - (((rs1!tra_amount) - (rs1!tra_amount * rs!disc_rate / 100)) * rs!disc_percentage / 100) - Val(adddisc)
                            End If
        ' check this carefully suspected error
                       '   If rs1!tcurrency = " USD" Then
                       '    ntra = ntra * Val(rs1!tconvertion)
                       '   End If
        ' check this carefully suspected error
        
                          Sqlqry2 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                 & Trim(rs!Month) & " '," & rs!monthind & ",'" & dumregion & "','" & rs1!tcurrency & "'," & rs1!tra_amount & "," & ntra & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                 & rs1!sub_media & "'," _
                                                 & rs!bo_ref & "," _
                                                 & Val(rs1!tra_amount) & "," & 0 & "," & 0 & "," & rs!disc_percentage & "," & scharge & ", '" _
                                                 & Trim(rs!disc_rate) & "'," _
                                                 & adddisc & "," _
                                                 & ntra & ")"
                         
                         
                                             
                      ElseIf rs1!Type = "Free" Then
                          Sqlqry2 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & dumregion & "','" & rs1!tcurrency & "'," & 0 & "," & 0 & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs1!sub_media & "'," _
                                             & rs!bo_ref & "," _
                                             & 0 & "," & Val(rs1!tra_amount) & "," & 0 & "," & 0 & "," & 0 & ", '" _
                                             & 0 & "'," _
                                             & 0 & "," & 0 & ")"
                       Else
                          Sqlqry2 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & dumregion & "','" & rs1!tcurrency & "'," & 0 & "," & 0 & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs1!sub_media & "'," _
                                             & rs!bo_ref & "," _
                                             & 0 & "," & 0 & "," & Val(rs1!tra_amount) & "," & 0 & "," & 0 & ", '" _
                                             & 0 & "'," _
                                             & 0 & "," & 0 & ")"
                       End If
                     
                                ws.BeginTrans
                                db.Execute (Sqlqry2)
                                ws.CommitTrans
                  rs1.MoveNext
                  Loop
                 End If
               rs.MoveNext
               Loop
           End If
    Else
      Sqlqry = "Select * from To_Client where Media='Cinema'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into To_Client1 values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!bo_ref & "," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!NET_Amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
           
     Sqlqry = " Delete * from To_Client where media='Cinema'"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
              
        adddisc = 0
        scharge = 0
        ntra = 0
                
                       
              
        Sqlqry = "Select * from To_Client1"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
              
                    
                Do Until rs.EOF
                 
                adddisc = 0
                scharge = 0
                ntra = 0
                
                 Sqlqry1 = "Select * from bo_tracin where serial_no='" & Trim(rs!serial_no) & "' AND TYPE ='Paid'"
                 Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                    If rs1.RecordCount = 0 Then
                        adddisc = rs!add_discount
                        scharge = rs!surcharge
                    Else
                        adddisc = rs!add_discount / rs1.RecordCount
                        scharge = rs!surcharge / rs1.RecordCount
                    End If
                
                
                
                
                Sqlqry1 = "Select * from bo_tracin where serial_no='" & Trim(rs!serial_no) & "' "
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                 If rs1.RecordCount <> 0 Then
                  rs1.MoveFirst
                   Do Until rs1.EOF
                   
                     Sqlqry2 = "Select region from cinema_rates where sub_media='" & rs1!sub_media & "' "
                     Set rs2 = db.OpenRecordset(Sqlqry2, dbOpenDynaset)
                     If IsNull(rs2.Fields(0)) = False Then dumregion = rs2.Fields(0)
                     
                     
                     If rs1!Type = "Paid" Then
                       
                            If rs!add_discount = 0 Then
                              ntra = Val(rs1!tra_amount) - (Val(rs1!tra_amount) * rs!disc_rate / 100) - (((rs1!tra_amount) - (rs1!tra_amount * rs!disc_rate / 100)) * rs!disc_percentage / 100)
                            Else
                              ntra = Val(rs1!tra_amount) - (Val(rs1!tra_amount) * rs!disc_rate / 100) - (((rs1!tra_amount) - (rs1!tra_amount * rs!disc_rate / 100)) * rs!disc_percentage / 100) - Val(adddisc)
                            End If
        ' check this carefully suspected error
                       '   If rs1!tcurrency = " USD" Then
                       '    ntra = ntra * Val(rs1!tconvertion)
                       '   End If
        ' check this carefully suspected error
        
                          Sqlqry2 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                 & Trim(rs!Month) & " '," & rs!monthind & ",'" & dumregion & "','" & rs1!tcurrency & "'," & rs1!tra_amount & "," & ntra & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                 & rs1!sub_media & "'," _
                                                 & rs!bo_ref & "," _
                                                 & Val(rs1!tra_amount) & "," & 0 & "," & 0 & "," & rs!disc_percentage & "," & scharge & ", '" _
                                                 & Trim(rs!disc_rate) & "'," _
                                                 & adddisc & "," _
                                                 & ntra & ")"
                         
                         
                                             
                      ElseIf rs1!Type = "Free" Then
                          Sqlqry2 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & dumregion & "','" & rs1!tcurrency & "'," & 0 & "," & 0 & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs1!sub_media & "'," _
                                             & rs!bo_ref & "," _
                                             & 0 & "," & Val(rs1!tra_amount) & "," & 0 & "," & 0 & "," & 0 & ", '" _
                                             & 0 & "'," _
                                             & 0 & "," & 0 & ")"
                       Else
                          Sqlqry2 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & dumregion & "','" & rs1!tcurrency & "'," & 0 & "," & 0 & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs1!sub_media & "'," _
                                             & rs!bo_ref & "," _
                                             & 0 & "," & 0 & "," & Val(rs1!tra_amount) & "," & 0 & "," & 0 & ", '" _
                                             & 0 & "'," _
                                             & 0 & "," & 0 & ")"
                       End If
                     
                                ws.BeginTrans
                                db.Execute (Sqlqry2)
                                ws.CommitTrans
                  rs1.MoveNext
                  Loop
                 End If
               rs.MoveNext
               Loop
           End If
    End If
    
   If cboregion.Text <> "All" Then
        Sqlqry2 = "Delete * from toclient where region<>'" & cboregion.Text & "'"
        ws.BeginTrans
        db.Execute (Sqlqry2)
        ws.CommitTrans
   End If
     
End Sub
Private Sub Form_Load()
cbomonthfrom.AddItem "January"
cbomonthfrom.AddItem "February"
cbomonthfrom.AddItem "March"
cbomonthfrom.AddItem "April"
cbomonthfrom.AddItem "May"
cbomonthfrom.AddItem "June"
cbomonthfrom.AddItem "July"
cbomonthfrom.AddItem "August"
cbomonthfrom.AddItem "September"
cbomonthfrom.AddItem "October"
cbomonthfrom.AddItem "November"
cbomonthfrom.AddItem "December"

LblMediaName.Caption = ""
lblSubMediaName.Caption = ""


CboCurrency.AddItem "DHS"
CboCurrency.AddItem "USD"
 
Populateregion

i = 2000

For i = 2000 To 2100
 Cboyear.AddItem i
Next
X = 0

 Cboyear.Text = Year(Now())
 
 X = Month(Now())
  
If X = 1 Then
   cbomonthfrom.ListIndex = 0
ElseIf X = 2 Then
   cbomonthfrom.ListIndex = 1
ElseIf X = 3 Then
   cbomonthfrom.ListIndex = 2
ElseIf X = 4 Then
   cbomonthfrom.ListIndex = 3
ElseIf X = 5 Then
   cbomonthfrom.ListIndex = 4
ElseIf X = 6 Then
   cbomonthfrom.ListIndex = 5
ElseIf X = 7 Then
   cbomonthfrom.ListIndex = 6
ElseIf X = 8 Then
   cbomonthfrom.ListIndex = 7
ElseIf X = 9 Then
   cbomonthfrom.ListIndex = 8
ElseIf X = 10 Then
   cbomonthfrom.ListIndex = 9
ElseIf X = 11 Then
   cbomonthfrom.ListIndex = 10
Else
   cbomonthfrom.ListIndex = 11
End If

PopulateAgencycodes
populateMedia
populateproducts

End Sub
Private Sub Populateregion()
    cboregion.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select distinct(region) from bo_mas Order by region"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        cboregion.AddItem "All"
        rs.MoveFirst
       Do Until rs.EOF
        If IsEmpty(rs!region) = True Then
         rs.MoveNext
        Else
         cboregion.AddItem rs!region
         rs.MoveNext
        End If
       Loop
    End If
 End Sub
Private Sub populateproducts()
    CboProduct.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from products Order by Product_Name"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
           CboProduct.AddItem "All"
        rs.MoveFirst
            Do Until rs.EOF
              CboProduct.AddItem rs!product_name
            rs.MoveNext
       Loop
    End If
 End Sub

Private Sub populateMedia()
    cboMediaType.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from Media Order by Media_Type"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
         Exit Sub
    Else
        cboMediaType.AddItem "All"
        cboMediaType.AddItem "Cinema"
        cboMediaType.AddItem "Magazine"
        cboMediaType.AddItem "Online"
        cboMediaType.AddItem "Television"
        rs.MoveFirst
            Do Until rs.EOF
              cboMediaType.AddItem rs!Media_Type & " " & Trim(rs!sub_media)
            rs.MoveNext
       Loop
    End If
 End Sub
 
Private Sub curadjustments()

     Sqlqry = " Delete * from To_Client1"
     ws.BeginTrans
     db.Execute (Sqlqry)
     ws.CommitTrans
              
     If CboCurrency.Text = "USD" Then
       
        Sqlqry = "Select * from To_Client where Tcurrency='DHS'"
        'Sqlqry = "Select * from To_Client "
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into To_Client1 values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & Round(rs!tra_gamount / convertion, 2) & "," & Round(rs!tra_namount / convertion, 2) & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!bo_ref & "," _
                                             & Round(Val(rs!gross_amount) / convertion, 2) & "," _
                                             & Round(Val(rs!Tot_free) / convertion, 2) & "," _
                                             & Round(Val(rs!Tot_barter) / convertion, 2) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Round(Val(rs!surcharge) / convertion, 2) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Round(Val(rs!add_discount) / convertion, 2) & "," _
                                             & Round(Val(rs!NET_Amount) / convertion, 2) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
           
            Sqlqry = " Delete * from To_Client where Tcurrency='DHS'"
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
                     
            Sqlqry = "Select * from To_Client1 "
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & Val(rs!tra_gamount) & "," & Val(rs!tra_namount) & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!bo_ref & "," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!NET_Amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
           
     Else
        Sqlqry = "Select * from To_Client where Tcurrency='USD'"
        Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
           If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into To_Client1 values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & Round(rs!tra_gamount * convertion, 2) & "," & Round(rs!tra_namount * convertion, 2) & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!bo_ref & "," _
                                             & Round(Val(rs!gross_amount) * convertion, 2) & "," _
                                             & Round(Val(rs!Tot_free) * convertion, 2) & "," _
                                             & Round(Val(rs!Tot_barter) * convertion, 2) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Round(Val(rs!surcharge) * convertion, 2) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Round(Val(rs!add_discount) * convertion, 2) & "," _
                                             & Round(Val(rs!NET_Amount) * convertion, 2) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
           
            Sqlqry = " Delete * from To_Client where Tcurrency='USD'"
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
                     
            Sqlqry = "Select * from To_Client1 "
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
            If rs.RecordCount <> 0 Then
              rs.MoveFirst
                  Do Until rs.EOF
                     Sqlqry = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                             & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & Val(rs!tra_gamount) & "," & Val(rs!tra_namount) & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                             & rs!sub_media & "'," _
                                             & rs!bo_ref & "," _
                                             & Val(rs!gross_amount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & "," _
                                             & Val(rs!disc_percentage) & "," _
                                             & Val(rs!surcharge) & ",'" _
                                             & Trim(rs!disc_rate) & "'," _
                                             & Val(rs!add_discount) & "," _
                                             & Val(rs!NET_Amount) & ")"
                         ws.BeginTrans
                         db.Execute (Sqlqry)
                         ws.CommitTrans
                     rs.MoveNext
            
                 Loop
           End If
        
       End If
   
End Sub
Private Sub PopulateAgencycodes()
    CboClient.Clear
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select * from clientdtls Order by clientName"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
      rs.MoveFirst
        CboClient.Clear
         CboClient.AddItem "All"
        Do Until rs.EOF
            CboClient.AddItem rs!clientname
            rs.MoveNext
        Loop
    End If
        
End Sub

Private Sub prevcin()

C = 0
                    d = 0
                    e = 0
                    f = 0
                    g = 0
                    
          If CboClient.Text <> "All" And cboregion.Text <> "All" And CboProduct.Text <> "All" Then
                    Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & n & "' and Client='" & Trim(CboClient.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Product ='" & Trim(CboProduct.Text) & "'"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                     Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                        
                            Sqlqry1 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                  & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                  & findfirstfixup(rs!Product) & "','" _
                                                  & findfirstfixup(rs!client) & "','" _
                                                  & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                  & rs!sub_media & "'," _
                                                  & rs!monthind & "," _
                                                  & Val(rs!tra_gamount) & "," _
                                                  & Val(rs!Tot_free) & "," _
                                                  & Val(rs!Tot_barter) & "," _
                                                  & Val(rs!disc_percentage) & "," _
                                                  & Val(rs!surcharge) & ",'" _
                                                  & Trim(rs!disc_rate) & "'," _
                                                  & Val(rs!add_discount) + crdtamt & "," _
                                                  & Val(rs!tra_namount) - crdtamt & ")"
                              ws.BeginTrans
                              db.Execute (Sqlqry1)
                              ws.CommitTrans
                                                
                                  
                        rs.MoveNext
                        Loop
                     End If
              ElseIf CboClient.Text = "All" And cboregion.Text <> "All" And CboProduct.Text <> "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & n & "'  and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' and Product ='" & Trim(CboProduct.Text) & "'"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                     Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                        
                            Sqlqry1 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                  & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                  & findfirstfixup(rs!Product) & "','" _
                                                  & findfirstfixup(rs!client) & "','" _
                                                  & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                  & rs!sub_media & "'," _
                                                  & rs!monthind & "," _
                                                  & Val(rs!tra_gamount) & "," _
                                                  & Val(rs!Tot_free) & "," _
                                                  & Val(rs!Tot_barter) & "," _
                                                  & Val(rs!disc_percentage) & "," _
                                                  & Val(rs!surcharge) & ",'" _
                                                  & Trim(rs!disc_rate) & "'," _
                                                  & Val(rs!add_discount) + crdtamt & "," _
                                                  & Val(rs!tra_namount) - crdtamt & ")"
                              ws.BeginTrans
                              db.Execute (Sqlqry1)
                              ws.CommitTrans
                                    
                        rs.MoveNext
                        Loop
                     End If
            ElseIf CboClient.Text = "All" And cboregion.Text = "All" And CboProduct.Text <> "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " AND CANCELL='N' and Media='" & n & "'  and Product ='" & Trim(CboProduct.Text) & "'"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                     Do Until rs.EOF
                          crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                        
                            Sqlqry1 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                  & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                  & findfirstfixup(rs!Product) & "','" _
                                                  & findfirstfixup(rs!client) & "','" _
                                                  & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                  & rs!sub_media & "'," _
                                                  & rs!monthind & "," _
                                                  & Val(rs!tra_gamount) & "," _
                                                  & Val(rs!Tot_free) & "," _
                                                  & Val(rs!Tot_barter) & "," _
                                                  & Val(rs!disc_percentage) & "," _
                                                  & Val(rs!surcharge) & ",'" _
                                                  & Trim(rs!disc_rate) & "'," _
                                                  & Val(rs!add_discount) + crdtamt & "," _
                                                  & Val(rs!tra_namount) - crdtamt & ")"
                              ws.BeginTrans
                              db.Execute (Sqlqry1)
                              ws.CommitTrans
                                    
                        rs.MoveNext
                        Loop
                     End If
                End If
               prevcin2
                     
                                
End Sub
Private Sub prevcin2()

            
            If CboClient.Text = "All" And cboregion.Text = "All" And CboProduct.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " AND CANCELL='N' and Media='" & n & "' "
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                     Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                        
                            Sqlqry1 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                  & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                  & findfirstfixup(rs!Product) & "','" _
                                                  & findfirstfixup(rs!client) & "','" _
                                                  & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                  & rs!sub_media & "'," _
                                                  & rs!monthind & "," _
                                                  & Val(rs!tra_gamount) & "," _
                                                  & Val(rs!Tot_free) & "," _
                                                  & Val(rs!Tot_barter) & "," _
                                                  & Val(rs!disc_percentage) & "," _
                                                  & Val(rs!surcharge) & ",'" _
                                                  & Trim(rs!disc_rate) & "'," _
                                                  & Val(rs!add_discount) + crdtamt & "," _
                                                  & Val(rs!tra_namount) - crdtamt & ")"
                              ws.BeginTrans
                              db.Execute (Sqlqry1)
                              ws.CommitTrans
                                    
                        rs.MoveNext
                        Loop
                     End If
            
            ElseIf CboClient.Text <> "All" And cboregion.Text <> "All" And CboProduct.Text = "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & n & "'  and Client='" & Trim(CboClient.Text) & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N' "
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                     Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                        
                            Sqlqry1 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                  & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                  & findfirstfixup(rs!Product) & "','" _
                                                  & findfirstfixup(rs!client) & "','" _
                                                  & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                  & rs!sub_media & "'," _
                                                  & rs!monthind & "," _
                                                  & Val(rs!tra_gamount) & "," _
                                                  & Val(rs!Tot_free) & "," _
                                                  & Val(rs!Tot_barter) & "," _
                                                  & Val(rs!disc_percentage) & "," _
                                                  & Val(rs!surcharge) & ",'" _
                                                  & Trim(rs!disc_rate) & "'," _
                                                  & Val(rs!add_discount) + crdtamt & "," _
                                                  & Val(rs!tra_namount) - crdtamt & ")"
                              ws.BeginTrans
                              db.Execute (Sqlqry1)
                              ws.CommitTrans
                                    
                        rs.MoveNext
                        Loop
                     End If
            ElseIf CboClient.Text <> "All" And cboregion.Text = "All" And CboProduct.Text <> "All" Then
                Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & n & "' and Client='" & Trim(CboClient.Text) & "' AND CANCELL='N' and Product ='" & Trim(CboProduct.Text) & "'"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                     Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                        
                            Sqlqry1 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                  & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                  & findfirstfixup(rs!Product) & "','" _
                                                  & findfirstfixup(rs!client) & "','" _
                                                  & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                  & rs!sub_media & "'," _
                                                  & rs!monthind & "," _
                                                  & Val(rs!tra_gamount) & "," _
                                                  & Val(rs!Tot_free) & "," _
                                                  & Val(rs!Tot_barter) & "," _
                                                  & Val(rs!disc_percentage) & "," _
                                                  & Val(rs!surcharge) & ",'" _
                                                  & Trim(rs!disc_rate) & "'," _
                                                  & Val(rs!add_discount) + crdtamt & "," _
                                                  & Val(rs!tra_namount) - crdtamt & ")"
                              ws.BeginTrans
                              db.Execute (Sqlqry1)
                              ws.CommitTrans
                                    
                        rs.MoveNext
                        Loop
                     End If
            ElseIf CboClient.Text = "All" And cboregion.Text <> "All" And CboProduct.Text = "All" Then
                    Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & n & "' and region='" & Trim(cboregion.Text) & "'  and CANCELL='N'"
                   ' MsgBox Sqlqry
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                     Do Until rs.EOF
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                        
                            Sqlqry1 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                  & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                  & findfirstfixup(rs!Product) & "','" _
                                                  & findfirstfixup(rs!client) & "','" _
                                                  & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                  & rs!sub_media & "'," _
                                                  & rs!monthind & "," _
                                                  & Val(rs!tra_gamount) & "," _
                                                  & Val(rs!Tot_free) & "," _
                                                  & Val(rs!Tot_barter) & "," _
                                                  & Val(rs!disc_percentage) & "," _
                                                  & Val(rs!surcharge) & ",'" _
                                                  & Trim(rs!disc_rate) & "'," _
                                                  & Val(rs!add_discount) + crdtamt & "," _
                                                  & Val(rs!tra_namount) - crdtamt & ")"
                              ws.BeginTrans
                              db.Execute (Sqlqry1)
                              ws.CommitTrans
                                    
                        rs.MoveNext
                        Loop
                     End If
                   
            ElseIf CboClient.Text = "All" And cboregion.Text = "All" And CboProduct.Text <> "All" Then
                    Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " and Media='" & n & "' AND CANCELL='N' and product='" & Trim(CboProduct.Text) & "'"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                     Do Until rs.EOF
                          crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                        
                            Sqlqry1 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                  & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                  & findfirstfixup(rs!Product) & "','" _
                                                  & findfirstfixup(rs!client) & "','" _
                                                  & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                  & rs!sub_media & "'," _
                                                  & rs!monthind & "," _
                                                  & Val(rs!tra_gamount) & "," _
                                                  & Val(rs!Tot_free) & "," _
                                                  & Val(rs!Tot_barter) & "," _
                                                  & Val(rs!disc_percentage) & "," _
                                                  & Val(rs!surcharge) & ",'" _
                                                  & Trim(rs!disc_rate) & "'," _
                                                  & Val(rs!add_discount) + crdtamt & "," _
                                                  & Val(rs!tra_namount) - crdtamt & ")"
                              ws.BeginTrans
                              db.Execute (Sqlqry1)
                              ws.CommitTrans
                                    
                        rs.MoveNext
                        Loop
                     End If
                     
            ElseIf CboClient.Text <> "All" And cboregion.Text = "All" And CboProduct.Text = "All" Then
                    Sqlqry = "Select * from bo_mas where year='" & Val(Cboyear.Text) & "' AND monthind >=" & Val(cbomonthfrom.ListIndex) & " AND monthind<= " & Val(cbomonthTo.ListIndex) + Val(cbomonthfrom.ListIndex) & " AND CANCELL='N' and Media='" & n & "'  and Client='" & Trim(CboClient.Text) & "'"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                     If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                     Do Until rs.EOF
                        
                         crdtamt = 0
                         Sqlqry1 = "select sum(tra_amount) from crdt_mas where val(mid(ref_no,1,7))='" & rs!serial_no & "'"
                         Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                         If IsNull(rs1.Fields(0)) = False Then crdtamt = rs1.Fields(0)
                        
                            Sqlqry1 = " Insert into To_Client values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "','" _
                                                  & Trim(rs!Month) & " '," & rs!monthind & ",'" & findfirstfixup(Trim(rs!region)) & "','" & rs!tcurrency & "'," & rs!tra_gamount & "," & rs!tra_namount & ",'" _
                                                  & findfirstfixup(rs!Product) & "','" _
                                                  & findfirstfixup(rs!client) & "','" _
                                                  & findfirstfixup(rs!Agency) & "','" & rs!Media & "','" _
                                                  & rs!sub_media & "'," _
                                                  & rs!monthind & "," _
                                                  & Val(rs!tra_gamount) & "," _
                                                  & Val(rs!Tot_free) & "," _
                                                  & Val(rs!Tot_barter) & "," _
                                                  & Val(rs!disc_percentage) & "," _
                                                  & Val(rs!surcharge) & ",'" _
                                                  & Trim(rs!disc_rate) & "'," _
                                                  & Val(rs!add_discount) + crdtamt & "," _
                                                  & Val(rs!tra_namount) - crdtamt & ")"
                              ws.BeginTrans
                              db.Execute (Sqlqry1)
                              ws.CommitTrans
                                    
                        rs.MoveNext
                        Loop
                     End If
                     
                     
                 End If
            
End Sub
Private Function ValidateData()

ValidateData = False
If Cboyear.Text = "" Then
   MsgBox "Invalid year", vbInformation, "Invalid Entry"
   Cboyear.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
 ElseIf cbomonthfrom.Text = "" Then
   MsgBox "Select Month From", vbInformation, "Invalid Entry"
   cbomonthfrom.SetFocus
   SendKeys " {Home} + {end} "
   Exit Function
 ElseIf cbomonthTo.Text = "" Then
   MsgBox "Select Month To", vbInformation, "Invalid Entry"
   cbomonthTo.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
 ElseIf CboClient.Text = "" Then
   MsgBox "Select Agency", vbInformation, "Invalid Entry"
   CboClient.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
 ElseIf cboMediaType.Text = "" Then
   MsgBox "Select Media Type", vbInformation, "Invalid Entry"
   cboMediaType.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
 ElseIf CboProduct.Text = "" Then
   MsgBox "Select Product", vbInformation, "Invalid Entry"
   CboProduct.SetFocus
   SendKeys " {Home} + {End} "
   Exit Function
Else
  ValidateData = True
End If
End Function

Private Sub textclear()
 CboClient.ListIndex = -1
 CboProduct.ListIndex = -1
 cboMediaType.ListIndex = -1
 Cboyear.ListIndex = -1
 cbomonthfrom.ListIndex = -1
 cbomonthTo.ListIndex = -1
 


End Sub



