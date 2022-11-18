VERSION 2.00
Begin Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Com Routines"
   ClientHeight    =   1650
   ClientLeft      =   2205
   ClientTop       =   1110
   ClientWidth     =   4785
   FontTransparent =   0   'False
   Height          =   2055
   Left            =   2145
   LinkMode        =   1  'Source
   LinkTopic       =   "System"
   ScaleHeight     =   1650
   ScaleWidth      =   4785
   Top             =   765
   Width           =   4905
   Begin MSComm prtcomm 
      CommPort        =   2
      Interval        =   55
      Left            =   3000
      RTSEnable       =   -1  'True
      Settings        =   "9600,e,7,1"
      Top             =   720
   End
   Begin TextBox txtdata 
      Height          =   285
      Left            =   1320
      LinkTimeout     =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3255
   End
   Begin MSComm comm 
      Interval        =   1000
      Left            =   2040
      NullDiscard     =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      Settings        =   "4800,e,7,1"
      SThreshold      =   1
      Top             =   720
   End
   Begin Label Label1 
      Caption         =   "Command"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Dim comnbr    As Integer
Dim i         As Integer
Dim j         As Integer
Dim bytecom   As String
Dim buffer    As String
Dim datap     As String
Dim comstr    As String
Dim combidon  As String
Dim comflag As String
Dim cntlrcur As Integer
Dim cntlrtab As String
Dim publoop As Integer
Dim pubcom As Integer
Dim comtimer As String

Sub comm_OnComm ()
If Comm.CommEvent = 2 Then
    While Comm.InBufferCount > 0
	  comstr = Comm.Input
	  For i = 1 To Len(comstr)
	      bytecom = Mid$(comstr, i, 1)
	      If bytecom > Chr$(15) Then
		 buffer = buffer + bytecom
	      End If
	      If bytecom = Chr$(13) Then
		 frmmain.txtdata.LinkExecute buffer
		 buffer = ""
	      End If
	     Next i
	   Wend
      End If
End Sub

Sub Command1_Click ()
    prtcomm.Output = txtdata.Text
End Sub

Sub Form_Load ()
    Comm.SThreshold = 1
    Comm.RThreshold = 1
    Comm.CommPort = 2
    Comm.Settings = "9600,E,7,1"
    Comm.InputLen = 0
    Comm.Interval = 500
    Comm.InBufferSize = 20480
    Comm.OutBufferSize = 512
    combidon = Chr$(0)
    comstr = ""
    Comm.PortOpen = True

    prtcomm.SThreshold = 1
    prtcomm.RThreshold = 1
    prtcomm.CommPort = 1
    prtcomm.Settings = "9600,N,8,1"
    prtcomm.InputLen = 0
    prtcomm.InBufferSize = 5000
    prtcomm.OutBufferSize = 5120
    prtcomm.ParityReplace = ""
    prtcomm.PortOpen = True
Rem DDE Server
    appname = "myservice"
    topic = "mycomtopic"
    item = "mycomitem"
    txtdata.LinkMode = NONE
    txtdata.LinkTopic = "myservice|mycomtopic"
    txtdata.LinkItem = "mycomitem"
    txtdata.LinkMode = 1
    CreateLink = Err
    txtdata.Text = ""
Rem Variables
    comtimer = ""
    publoop = 2
    bytecom = Space(10)
    pubcom = True
    cntlrcur = 0
    comflag = "O"
    'myrange = 15950
    myrange = 900
    Open "ALL.TXT" For Append Access Write As #1
End Sub

Sub FORM_UNLOAD (Cancel As Integer)
    Comm.PortOpen = False
    prtcomm.PortOpen = Fals
    Close #1
End Sub

Sub txtdata_Change ()
    If Mid(txtdata.Text, 1, 1) = "P" Then
	prtcomm.Output = Mid(txtdata.Text, 2)
	txtdata.Text = " "
    End If
    If txtdata.Text = "U" Then
       Unload frmmain
    End If
End Sub

Private Sub wait (twait As Integer)
For ttwait = 1 To twait
    For tttwait = 1 To 1000
    Next tttwait
Next ttwait
End Sub

