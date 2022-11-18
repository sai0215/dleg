VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmInvoicing 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Invoice "
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   315
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7935
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   10095
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   600
         TabIndex        =   19
         Top             =   600
         Width           =   9015
         Begin VB.OptionButton OptParticularno 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Serial #"
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
            Height          =   375
            Left            =   7680
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OptAgency 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Agency"
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
            Height          =   375
            Left            =   2880
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptProduct 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Product"
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
            Height          =   375
            Left            =   4200
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptSerialNo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Serial &Range"
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
            Height          =   375
            Left            =   5400
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton Optmedia 
            BackColor       =   &H00C0C0C0&
            Caption         =   "M&edia"
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
            Height          =   375
            Left            =   1560
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptMonth 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Month"
            DisabledPicture =   "FrmInvoicing.frx":0000
            DragIcon        =   "FrmInvoicing.frx":0442
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   480
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Fraemp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3495
         Left            =   1680
         TabIndex        =   3
         Top             =   2400
         Width           =   6735
         Begin VB.Frame fraMonth 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Invoices - -  Monthly"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   3495
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   6735
            Begin VB.Frame Frame2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Month-Year"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   1695
               Left            =   600
               TabIndex        =   5
               Top             =   720
               Width           =   5415
               Begin VB.ComboBox cboYear 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   3480
                  Style           =   2  'Dropdown List
                  TabIndex        =   7
                  Top             =   840
                  Width           =   1455
               End
               Begin VB.ComboBox cboMonth 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   6
                  Top             =   840
                  Width           =   1695
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Year"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   3840
                  TabIndex        =   9
                  Top             =   600
                  Width           =   465
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Month"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   1440
                  TabIndex        =   8
                  Top             =   600
                  Width           =   615
               End
            End
         End
         Begin VB.Frame Fraserialno 
            BackColor       =   &H00FFFFFF&
            Caption         =   " Invoices  - - Serial Number Wise"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   3495
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   6735
            Begin VB.Frame Frame4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Serial Number"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   1695
               Left            =   840
               TabIndex        =   14
               Top             =   720
               Width           =   5415
               Begin VB.ComboBox CboSerialTo 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   3360
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.ComboBox CboSerialFrom 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   960
                  Style           =   2  'Dropdown List
                  TabIndex        =   15
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "To"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   3000
                  TabIndex        =   18
                  Top             =   840
                  Width           =   225
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "From"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   360
                  TabIndex        =   17
                  Top             =   840
                  Width           =   465
               End
            End
         End
         Begin VB.Frame FraProduct 
            BackColor       =   &H00FFFFFF&
            Caption         =   " Invoices  - - Product Wise"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   3495
            Left            =   1920
            TabIndex        =   38
            Top             =   0
            Width           =   6735
            Begin VB.Frame Frame9 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Product"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   1935
               Left            =   840
               TabIndex        =   39
               Top             =   720
               Width           =   5415
               Begin VB.ComboBox CboPYear 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   42
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.ComboBox CboPMonth 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   3240
                  Style           =   2  'Dropdown List
                  TabIndex        =   41
                  Top             =   480
                  Width           =   1815
               End
               Begin VB.ComboBox CboProduct 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   40
                  Top             =   1200
                  Width           =   3975
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Product"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   240
                  TabIndex        =   45
                  Top             =   1320
                  Width           =   765
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Month"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   2520
                  TabIndex        =   44
                  Top             =   600
                  Width           =   615
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Year"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   360
                  TabIndex        =   43
                  Top             =   600
                  Width           =   585
               End
            End
         End
         Begin VB.Frame FraAgency 
            BackColor       =   &H00FFFFFF&
            Caption         =   " Invoices  - - Agency Wise"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   3495
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   6735
            Begin VB.Frame Frame7 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Agency"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   1935
               Left            =   840
               TabIndex        =   31
               Top             =   720
               Width           =   5415
               Begin VB.ComboBox CboAgency 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   34
                  Top             =   1200
                  Width           =   3975
               End
               Begin VB.ComboBox CboAmonth 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   3240
                  Style           =   2  'Dropdown List
                  TabIndex        =   33
                  Top             =   480
                  Width           =   1815
               End
               Begin VB.ComboBox CboAyear 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   32
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Year"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   360
                  TabIndex        =   37
                  Top             =   600
                  Width           =   585
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Month"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   2520
                  TabIndex        =   36
                  Top             =   600
                  Width           =   615
               End
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Agency"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   270
                  TabIndex        =   35
                  Top             =   1320
                  Width           =   735
               End
            End
         End
         Begin VB.Frame FraMedia 
            BackColor       =   &H00FFFFFF&
            Caption         =   " Invoices  - - Media Wise"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   3495
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   6735
            Begin VB.Frame Frame3 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Media"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   1935
               Left            =   840
               TabIndex        =   11
               Top             =   720
               Width           =   5415
               Begin VB.ComboBox CboMYear 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   24
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.ComboBox CboMMonth 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   3240
                  Style           =   2  'Dropdown List
                  TabIndex        =   23
                  Top             =   480
                  Width           =   1815
               End
               Begin VB.ComboBox CboMedia 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   12
                  Top             =   1200
                  Width           =   3975
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Media"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   360
                  TabIndex        =   27
                  Top             =   1320
                  Width           =   645
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Month"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   2520
                  TabIndex        =   26
                  Top             =   600
                  Width           =   615
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Year"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   360
                  TabIndex        =   25
                  Top             =   600
                  Width           =   585
               End
            End
         End
         Begin VB.Frame FraParticularNo 
            BackColor       =   &H00FFFFFF&
            Caption         =   " Invoice  - - Particular Number"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   3495
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   6735
            Begin VB.Frame Frame8 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Particular Number"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   1695
               Left            =   840
               TabIndex        =   48
               Top             =   720
               Width           =   5415
               Begin VB.ComboBox CboParticularNo 
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
                  ForeColor       =   &H00000080&
                  Height          =   360
                  Left            =   2160
                  Style           =   1  'Simple Combo
                  TabIndex        =   49
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "From"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404080&
                  Height          =   240
                  Left            =   1560
                  TabIndex        =   50
                  Top             =   840
                  Width           =   465
               End
            End
         End
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<< &Back"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton cmdWork 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6600
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   720
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
End
Attribute VB_Name = "frmInvoicing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sqlqry As String
Dim Sqlqry1 As String
Dim Sqlqry2 As String
Dim sqlqry3 As String
Dim Sqlqry4 As String
Dim Sqlqry5 As String
Dim Sqlqry6 As String
Dim Sqlqry7 As String
Dim Sqlqry8 As String
Dim Sqlqry9 As String
Dim Sqlqry10 As String
Dim SQLQRY11 As String
Dim SQLQRY12 As String
Dim rs As Recordset
Dim rs1 As Recordset
Dim SNo As Long
Dim X, Y, Z As Currency
Dim SelMonth As Integer
Dim FirstDate As Date
Dim LastDate As Date

Private Sub CboParticularNo_Change()
    If CboParticularNo.Text = "" Then
      MsgBox "Select Serial Number"
      CboParticularNo.SetFocus
      Exit Sub
    End If
End Sub
Private Sub CboSerialTo_LostFocus()
If CboSerialFrom.Text = "" Then
  MsgBox "Select Serial # from"
  CboSerialFrom.SetFocus
  Exit Sub
ElseIf Val(CboSerialTo.Text) < Val(CboSerialFrom) Then
  MsgBox " Serial # To cannot be lesser then the Serial # From"
  CboSerialTo.SetFocus
  Exit Sub
End If
End Sub
Private Sub cmdBack_Click()
 Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub
Private Sub Populatemonthinv()
 If ValidateData = False Then Exit Sub

X = MsgBox("Do you want to print invoices for the month of " & cboMonth.Text & ", Year " & cboYear.Text, vbInformation + vbYesNo, "Confirmation")
   
If X = vbNo Then Exit Sub

Dim i As Integer
Dim j As Integer
    j = cboMonth.ListIndex + 1
    FirstDate = Now
    LastDate = Now
    i = DaysinMonth(j, cboYear.Text)
    FirstDate = Format(DateValue("1-" & j & "-" & cboYear), "DD/mm/yyyy")
    LastDate = Format(DateValue(i & "-" & j & "-" & cboYear), "DD/mm/yyyy")
    
    Sqlqry1 = "select * from bo_mas Where STATUS='N' AND invoice_Date Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "# order by media"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        
        If rs.RecordCount <> 0 Then
         rs.MoveFirst
         Do Until rs.EOF
         
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         Sqlqry = " Insert into invrep values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "',' " _
                                     & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" & Trim(rs!tcurrency) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                     & Val(rs!tra_gamount) & "," _
                                     & Val(rs!Tot_free) & "," _
                                     & Val(rs!Tot_barter) & ",'" _
                                     & Val(Trim(rs!disc_percentage)) & "','" _
                                     & Val(Trim(rs!disc_rate)) & "'," _
                                     & Val(Trim(rs!add_discount)) & "," _
                                     & Val(Trim(rs!surcharge)) & "," _
                                     & Val(Trim(rs!tra_namount)) & ",'" & Format(rs!invoice_date, "dd/mm/yyyy") & "')"
                                     
                                     
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
            
            If rs!Media = "Cinema" Then
            
                Sqlqry1 = "Select * from bo_tracin where serial_no ='" & Trim(rs!serial_no) & "'"
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                If rs1.RecordCount <> 0 Then
                   
                   rs1.MoveFirst
                   Do Until rs1.EOF
                   Set ws = DBEngine.Workspaces(0)
                   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                     If rs1!Type = "Paid" Then
                         Sqlqry2 = " Insert into dumbo_tracin values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                               & Trim(rs1!Month) & "','" _
                                               & findfirstfixup(rs1!Product) & "','" _
                                               & findfirstfixup(rs1!client) & "','" _
                                               & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                               & Trim(rs1!sub_Media) & "','" _
                                               & Trim(rs1!DATEFROM) & "','" _
                                               & Trim(rs1!Dateto) & "','" _
                                               & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                               & Trim(rs1!Day) & "','" _
                                               & Trim(rs1!Length) & "','" _
                                               & findfirstfixup(Trim(rs1!Description)) & "','" _
                                               & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                               & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                               & Val(Trim(rs1!tra_amount)) & ", " _
                                               & Val(Trim(rs1!tra_amount)) & " )"
                                               
                      Else
                          Sqlqry2 = " Insert into dumbo_tracin values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                               & Trim(rs1!Month) & "','" _
                                               & findfirstfixup(rs1!Product) & "','" _
                                               & findfirstfixup(rs1!client) & "','" _
                                               & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                               & Trim(rs1!sub_Media) & "','" _
                                               & Trim(rs1!DATEFROM) & "','" _
                                               & Trim(rs1!Dateto) & "','" _
                                               & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                               & Trim(rs1!Day) & "','" _
                                               & Trim(rs1!Length) & "','" _
                                               & findfirstfixup(Trim(rs1!Description)) & "','" _
                                               & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                               & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0', '0' )"
                       End If
                       
                       
                      ws.BeginTrans
                      db.Execute (Sqlqry2)
                      ws.CommitTrans
                    rs1.MoveNext
                   Loop
                  End If
               ElseIf rs!Media = "Online" Then
               
                  Sqlqry1 = "Select * from bo_traol where serial_no ='" & Trim(rs!serial_no) & "'"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                      If rs1.RecordCount <> 0 Then
                         rs1.MoveFirst
                         Do Until rs1.EOF
                         Set ws = DBEngine.Workspaces(0)
                         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         If rs1!Type = "Paid" Then
                           Sqlqry2 = " Insert into dumBo_traol values('" & rs1!serial_no & "','" & Trim(rs1!date_From) & "','" & Trim(rs1!DATE_TO) & "','" & rs1!Year & "','" _
                                                     & Trim(rs1!Month) & "','" _
                                                     & findfirstfixup(rs1!Product) & "','" _
                                                     & findfirstfixup(rs1!client) & "','" _
                                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                     & Trim(rs1!sub_Media) & "','" _
                                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                                     & Trim(rs1!Type) & "'," _
                                                     & Val(Trim(rs1!impressions)) & "," _
                                                     & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                                     & Val(Trim(rs1!tra_amount)) & "," _
                                                     & Val(Trim(rs1!tra_amount)) & ")"
                                                     
                           Else
                             Sqlqry2 = " Insert into dumBo_traol values('" & rs1!serial_no & "','" & Trim(rs1!date_From) & "','" & Trim(rs1!DATE_TO) & "','" & rs1!Year & "','" _
                                                     & Trim(rs1!Month) & "','" _
                                                     & findfirstfixup(rs1!Product) & "','" _
                                                     & findfirstfixup(rs1!client) & "','" _
                                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                     & Trim(rs1!sub_Media) & "','" _
                                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                                     & Trim(rs1!Type) & "'," _
                                                     & Val(Trim(rs1!impressions)) & "," _
                                                     & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0')"
                           End If
                                                     
                            ws.BeginTrans
                            db.Execute (Sqlqry2)
                            ws.CommitTrans
                          rs1.MoveNext
                         Loop
                       End If
         ElseIf rs!Media = "Television" Then
            Sqlqry1 = "Select * from bo_traTv where serial_no ='" & Trim(rs!serial_no) & "'"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                  
                  rs1.MoveFirst
                  Do Until rs1.EOF
                  Set ws = DBEngine.Workspaces(0)
                  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                   If rs1!Type = "Paid" Then
                    Sqlqry2 = " Insert into dumbo_tratv values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                              & Trim(rs1!Month) & "','" _
                                              & findfirstfixup(rs1!Product) & "','" _
                                              & findfirstfixup(rs1!client) & "','" _
                                              & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                              & Trim(rs1!sub_Media) & "','" _
                                              & findfirstfixup(Trim(rs1!bo_ref)) & "','" & Trim(rs1!code) & "','" & Trim(rs1!sec) & "','" _
                                              & Trim(rs1!Day) & "','" _
                                              & Trim(rs1!Time) & "','" _
                                              & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                              & Trim(rs1!Type) & "'," _
                                              & Val(Trim(rs1!spots)) & "," _
                                              & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                              & Val(Trim(rs1!tra_amount)) & "," _
                                              & Val(Trim(rs1!tra_amount)) & ")"
                    Else
                      Sqlqry2 = " Insert into dumbo_tratv values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                              & Trim(rs1!Month) & "','" _
                                              & findfirstfixup(rs1!Product) & "','" _
                                              & findfirstfixup(rs1!client) & "','" _
                                              & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                              & Trim(rs1!sub_Media) & "','" _
                                              & findfirstfixup(Trim(rs1!bo_ref)) & "','" & Trim(rs1!code) & "','" & Trim(rs1!sec) & "','" _
                                              & Trim(rs1!Day) & "','" _
                                              & Trim(rs1!Time) & "','" _
                                              & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                              & Trim(rs1!Type) & "'," _
                                              & Val(Trim(rs1!spots)) & "," _
                                              & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0')"
                    End If
                                              
                     ws.BeginTrans
                     db.Execute (Sqlqry2)
                     ws.CommitTrans
                   rs1.MoveNext
                  Loop
                End If
                
           ElseIf rs!Media = "Magazine" And rs!sub_Media = "ALAM ASSAYARRAT" Or rs!sub_Media = "ZEINA" Then
              Sqlqry1 = "Select * from bo_tramag where serial_no ='" & Trim(rs!serial_no) & "'"
              Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
              If rs1.RecordCount <> 0 Then
                 rs1.MoveFirst
                 Do Until rs1.EOF
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 If rs1!Type = "Paid" Then
                   Sqlqry2 = " Insert into dumbo_tramagext values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                             & Val(Trim(rs1!tra_amount)) & "," _
                                             & Val(Trim(rs1!tra_amount)) & ",' " _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 Else
                      Sqlqry2 = " Insert into dumbo_tramagext values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0','" _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 End If
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
                  rs1.MoveNext
                 Loop
               End If
        
                  
       
          Else
            Sqlqry1 = "Select * from bo_tramag where serial_no ='" & Trim(rs!serial_no) & "'"
            Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
              If rs1.RecordCount <> 0 Then
                 rs1.MoveFirst
                 Do Until rs1.EOF
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 If rs1!Type = "Paid" Then
                   Sqlqry2 = " Insert into dumbo_tramag values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                             & Val(Trim(rs1!tra_amount)) & "," _
                                             & Val(Trim(rs1!tra_amount)) & ",' " _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 Else
                      Sqlqry2 = " Insert into dumbo_tramag values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0','" _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 End If
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
                  rs1.MoveNext
                 Loop
               End If
           End If
                  
       rs.MoveNext
       Loop
      End If
    
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
     CrystalReport1.ReportFileName = App.Path & "\bocininv.rpt"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.DetailCopies = 1
     CrystalReport1.Action = 1
     
    
     
     
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
     CrystalReport1.ReportFileName = App.Path & "\bomaginv.rpt"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Action = 1
     
    
     
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
     CrystalReport1.ReportFileName = App.Path & "\bomaginvext.rpt"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Action = 1
    
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
     CrystalReport1.ReportFileName = App.Path & "\botelinv1.rpt"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Action = 1
        
        
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
     CrystalReport1.ReportFileName = App.Path & "\boolinv.rpt"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Action = 1
     
     cboYear.Text = Year(Now())
     cboMonth.ListIndex = Month(Now) - 1


End Sub
Private Sub populatemediainv()
If ValidateDatamedia = False Then Exit Sub
 
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from invrep"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans

 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tracin"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tramag"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tramagext"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
 
 
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_traol"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tratv"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans

X = MsgBox("Do you want to print invoices for the month of " & CboMMonth.Text & ", Year " & CboMYear.Text & ", Media " & CboMedia.Text, vbInformation + vbYesNo, "Confirmation")
   
If X = vbNo Then Exit Sub

Dim i As Integer
Dim j As Integer
    j = CboMMonth.ListIndex + 1
    FirstDate = Now
    LastDate = Now
    i = DaysinMonth(j, CboMYear.Text)
    FirstDate = Format(DateValue("1-" & j & "-" & CboMYear), "dd/mm/yyyy")
    LastDate = Format(DateValue(i & "-" & j & "-" & CboMYear), "DD/mm/yyyy")
    
            If CboMedia = "Cinema" Then
            
               Sqlqry1 = "select * from bo_mas Where STATUS='N' AND media='Cinema' and invoice_Date Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "# order by media"
              Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        
               If rs.RecordCount <> 0 Then
                rs.MoveFirst
                Do Until rs.EOF
                
                Set ws = DBEngine.Workspaces(0)
                Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                Sqlqry = " Insert into invrep values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "',' " _
                                            & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                            & findfirstfixup(rs!Product) & "','" _
                                            & findfirstfixup(rs!client) & "','" _
                                            & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                            & Trim(rs!sub_Media) & "','" & Trim(rs!tcurrency) & "','" _
                                            & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                            & Val(rs!tra_gamount) & "," _
                                            & Val(rs!Tot_free) & "," _
                                            & Val(rs!Tot_barter) & ",'" _
                                            & Val(Trim(rs!disc_percentage)) & "','" _
                                            & Val(Trim(rs!disc_rate)) & "'," _
                                            & Val(Trim(rs!add_discount)) & "," _
                                            & Val(Trim(rs!surcharge)) & "," _
                                            & Val(Trim(rs!tra_namount)) & ",'" & Format(rs!invoice_date, "dd/mm/yyyy") & "')"
                                            
                                            
                   ws.BeginTrans
                   db.Execute (Sqlqry)
                   ws.CommitTrans
                   
            
                  Sqlqry1 = "Select * from bo_tracin where serial_no ='" & Trim(rs!serial_no) & "'"
                  Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                  If rs1.RecordCount <> 0 Then
                   
                   rs1.MoveFirst
                   Do Until rs1.EOF
                   Set ws = DBEngine.Workspaces(0)
                   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                     If rs1!Type = "Paid" Then
                         Sqlqry2 = " Insert into dumbo_tracin values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                               & Trim(rs1!Month) & "','" _
                                               & findfirstfixup(rs1!Product) & "','" _
                                               & findfirstfixup(rs1!client) & "','" _
                                               & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                               & Trim(rs1!sub_Media) & "','" _
                                               & Trim(rs1!DATEFROM) & "','" _
                                               & Trim(rs1!Dateto) & "','" _
                                               & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                               & Trim(rs1!Day) & "','" _
                                               & Trim(rs1!Length) & "','" _
                                               & findfirstfixup(Trim(rs1!Description)) & "','" _
                                               & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                               & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                               & Val(Trim(rs1!tra_amount)) & ", " _
                                               & Val(Trim(rs1!tra_amount)) & " )"
                                               
                      Else
                          Sqlqry2 = " Insert into dumbo_tracin values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                               & Trim(rs1!Month) & "','" _
                                               & findfirstfixup(rs1!Product) & "','" _
                                               & findfirstfixup(rs1!client) & "','" _
                                               & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                               & Trim(rs1!sub_Media) & "','" _
                                               & Trim(rs1!DATEFROM) & "','" _
                                               & Trim(rs1!Dateto) & "','" _
                                               & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                               & Trim(rs1!Day) & "','" _
                                               & Trim(rs1!Length) & "','" _
                                               & findfirstfixup(Trim(rs1!Description)) & "','" _
                                               & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                               & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0', '0' )"
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
            
                  
                  
               ElseIf CboMedia.Text = "Online" Then
               
                    Sqlqry = "select * from bo_mas Where STATUS='N' AND media='Online' and invoice_Date Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "# order by media"
                    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                        
                        If rs.RecordCount <> 0 Then
                         rs.MoveFirst
                         Do Until rs.EOF
                         
                         Set ws = DBEngine.Workspaces(0)
                         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         Sqlqry1 = " Insert into invrep values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "',' " _
                                                     & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                                     & findfirstfixup(rs!Product) & "','" _
                                                     & findfirstfixup(rs!client) & "','" _
                                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                                     & Trim(rs!sub_Media) & "','" & Trim(rs!tcurrency) & "','" _
                                                     & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                                     & Val(rs!tra_gamount) & "," _
                                                     & Val(rs!Tot_free) & "," _
                                                     & Val(rs!Tot_barter) & ",'" _
                                                     & Val(Trim(rs!disc_percentage)) & "','" _
                                                     & Val(Trim(rs!disc_rate)) & "'," _
                                                     & Val(Trim(rs!add_discount)) & "," _
                                                     & Val(Trim(rs!surcharge)) & "," _
                                                     & Val(Trim(rs!tra_namount)) & ",'" & Format(rs!invoice_date, "dd/mm/yyyy") & "')"
                                                     
                                                     
                            ws.BeginTrans
                            db.Execute (Sqlqry1)
                            ws.CommitTrans
                            
                    
                  Sqlqry1 = "Select * from bo_traol where serial_no ='" & Trim(rs!serial_no) & "'"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                      If rs1.RecordCount <> 0 Then
                         
                         rs1.MoveFirst
                         Do Until rs1.EOF
                         Set ws = DBEngine.Workspaces(0)
                         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         If rs1!Type = "Paid" Then
                           Sqlqry2 = " Insert into dumBo_traol values('" & rs1!serial_no & "','" & Trim(rs1!date_From) & "','" & Trim(rs1!DATE_TO) & "','" & rs1!Year & "','" _
                                                     & Trim(rs1!Month) & "','" _
                                                     & findfirstfixup(rs1!Product) & "','" _
                                                     & findfirstfixup(rs1!client) & "','" _
                                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                     & Trim(rs1!sub_Media) & "','" _
                                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                                     & Trim(rs1!Type) & "'," _
                                                     & Val(Trim(rs1!impressions)) & "," _
                                                     & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                                     & Val(Trim(rs1!tra_amount)) & "," _
                                                     & Val(Trim(rs1!tra_amount)) & ")"
                                                     
                           Else
                             Sqlqry2 = " Insert into dumBo_traol values('" & rs1!serial_no & "','" & Trim(rs1!date_From) & "','" & Trim(rs1!DATE_TO) & "','" & rs1!Year & "','" _
                                                     & Trim(rs1!Month) & "','" _
                                                     & findfirstfixup(rs1!Product) & "','" _
                                                     & findfirstfixup(rs1!client) & "','" _
                                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                     & Trim(rs1!sub_Media) & "','" _
                                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                                     & Trim(rs1!Type) & "'," _
                                                     & Val(Trim(rs1!impressions)) & "," _
                                                     & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0')"
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
             
         ElseIf CboMedia = "Television" Then
            Sqlqry = "select * from bo_mas Where STATUS='N' AND media='Television' and invoice_Date Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "# order by media"
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                 Do Until rs.EOF
                 
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 Sqlqry1 = " Insert into invrep values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "',' " _
                                             & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                             & Trim(rs!sub_Media) & "','" & Trim(rs!tcurrency) & "','" _
                                             & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & ",'" _
                                             & Val(Trim(rs!disc_percentage)) & "','" _
                                             & Val(Trim(rs!disc_rate)) & "'," _
                                             & Val(Trim(rs!add_discount)) & "," _
                                             & Val(Trim(rs!surcharge)) & "," _
                                             & Val(Trim(rs!tra_namount)) & ",'" & Format(rs!invoice_date, "dd/mm/yyyy") & "')"
                                             
                                             
                    ws.BeginTrans
                    db.Execute (Sqlqry1)
                    ws.CommitTrans
                    
            
            Sqlqry1 = "Select * from bo_traTv where serial_no ='" & Trim(rs!serial_no) & "'"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                  
                  rs1.MoveFirst
                  Do Until rs1.EOF
                  Set ws = DBEngine.Workspaces(0)
                  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                   If rs1!Type = "Paid" Then
                    Sqlqry2 = " Insert into dumbo_tratv values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                              & Trim(rs1!Month) & "','" _
                                              & findfirstfixup(rs1!Product) & "','" _
                                              & findfirstfixup(rs1!client) & "','" _
                                              & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                              & Trim(rs1!sub_Media) & "','" _
                                              & findfirstfixup(Trim(rs1!bo_ref)) & "','" & Trim(rs1!code) & "','" & Trim(rs1!sec) & "','" _
                                              & Trim(rs1!Day) & "','" _
                                              & Trim(rs1!Time) & "','" _
                                              & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                              & Trim(rs1!Type) & "'," _
                                              & Val(Trim(rs1!spots)) & "," _
                                              & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                              & Val(Trim(rs1!tra_amount)) & "," _
                                              & Val(Trim(rs1!tra_amount)) & ")"
                    Else
                      Sqlqry2 = " Insert into dumbo_tratv values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                              & Trim(rs1!Month) & "','" _
                                              & findfirstfixup(rs1!Product) & "','" _
                                              & findfirstfixup(rs1!client) & "','" _
                                              & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                              & Trim(rs1!sub_Media) & "','" _
                                              & findfirstfixup(Trim(rs1!bo_ref)) & "','" & Trim(rs1!code) & "','" & Trim(rs1!sec) & "','" _
                                              & Trim(rs1!Day) & "','" _
                                              & Trim(rs1!Time) & "','" _
                                              & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                              & Trim(rs1!Type) & "'," _
                                              & Val(Trim(rs1!spots)) & "," _
                                              & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0')"
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
            
           'Elseif rs!media='Ma
            ElseIf CboMedia = "Magazine" Then
            
            
                Sqlqry = "select * from bo_mas Where STATUS='N' AND media='Magazine' and invoice_Date Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "# order by media"
                Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                    If rs.RecordCount <> 0 Then
                     rs.MoveFirst
                     Do Until rs.EOF
                     
                     Set ws = DBEngine.Workspaces(0)
                     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                     Sqlqry1 = " Insert into invrep values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "',' " _
                                                 & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                                 & findfirstfixup(rs!Product) & "','" _
                                                 & findfirstfixup(rs!client) & "','" _
                                                 & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                                 & Trim(rs!sub_Media) & "','" & Trim(rs!tcurrency) & "','" _
                                                 & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                                 & Val(rs!tra_gamount) & "," _
                                                 & Val(rs!Tot_free) & "," _
                                                 & Val(rs!Tot_barter) & ",'" _
                                                 & Val(Trim(rs!disc_percentage)) & "','" _
                                                 & Val(Trim(rs!disc_rate)) & "'," _
                                                 & Val(Trim(rs!add_discount)) & "," _
                                                 & Val(Trim(rs!surcharge)) & "," _
                                                 & Val(Trim(rs!tra_namount)) & ",'" & Format(rs!invoice_date, "dd/mm/yyyy") & "')"
                                                 
                                                 
                        ws.BeginTrans
                        db.Execute (Sqlqry1)
                        ws.CommitTrans
                        
                
                 Sqlqry1 = "Select * from bo_tramag where serial_no ='" & Trim(rs!serial_no) & "'"
                 Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                   If rs1.RecordCount <> 0 Then
                      rs1.MoveFirst
                      Do Until rs1.EOF
                      Set ws = DBEngine.Workspaces(0)
                      Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                      If rs1!Type = "Paid" Then
                        Sqlqry2 = " Insert into dumbo_tramagext values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                                  & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                                  & findfirstfixup(rs1!Product) & "','" _
                                                  & findfirstfixup(rs1!client) & "','" _
                                                  & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                  & Trim(rs1!sub_Media) & "','" _
                                                  & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                  & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                                  & Trim(rs1!Page) & "','" _
                                                  & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                  & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                                  & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                                  & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                                  & Val(Trim(rs1!tra_amount)) & "," _
                                                  & Val(Trim(rs1!tra_amount)) & ",' " _
                                                  & Val(Trim(rs1!agcom)) & "','" _
                                                  & Val(Trim(rs1!adper)) & "'," _
                                                  & Val(Trim(rs1!addisc)) & "," _
                                                  & Val(Trim(rs1!surcharge)) & ")"
                      Else
                           Sqlqry2 = " Insert into dumbo_tramagext values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                                  & Trim(rs1!Month) & "', " & Val(rs1!monthind) & ",'" _
                                                  & findfirstfixup(rs1!Product) & "','" _
                                                  & findfirstfixup(rs1!client) & "','" _
                                                  & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                  & Trim(rs1!sub_Media) & "','" _
                                                  & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                  & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                                  & Trim(rs1!Page) & "','" _
                                                  & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                  & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                                  & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                                  & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0','" _
                                                  & Val(Trim(rs1!agcom)) & "','" _
                                                  & Val(Trim(rs1!adper)) & "'," _
                                                  & Val(Trim(rs1!addisc)) & "," _
                                                  & Val(Trim(rs1!surcharge)) & ")"
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
           ElseIf Mid(CboMedia, 1, 3) = "Mag" Then
           'Else
            Sqlqry = "select * from bo_mas Where STATUS='N' AND media='Magazine' and sub_media='" & Mid(CboMedia.Text, 10, 30) & "' and invoice_Date Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "# order by sub_media"
        '   MsgBox Sqlqry
            Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
                
                If rs.RecordCount <> 0 Then
                 rs.MoveFirst
                 Do Until rs.EOF
                 
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 Sqlqry1 = " Insert into invrep values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "',' " _
                                             & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                             & findfirstfixup(rs!Product) & "','" _
                                             & findfirstfixup(rs!client) & "','" _
                                             & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                             & Trim(rs!sub_Media) & "','" & Trim(rs!tcurrency) & "','" _
                                             & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                             & Val(rs!tra_gamount) & "," _
                                             & Val(rs!Tot_free) & "," _
                                             & Val(rs!Tot_barter) & ",'" _
                                             & Val(Trim(rs!disc_percentage)) & "','" _
                                             & Val(Trim(rs!disc_rate)) & "'," _
                                             & Val(Trim(rs!add_discount)) & "," _
                                             & Val(Trim(rs!surcharge)) & "," _
                                             & Val(Trim(rs!tra_namount)) & ",'" & Format(rs!invoice_date, "dd/mm/yyyy") & "')"
                                             
                                             
                    ws.BeginTrans
                    db.Execute (Sqlqry1)
                    ws.CommitTrans
                    
            Sqlqry1 = "Select * from bo_tramag where serial_no ='" & Trim(rs!serial_no) & "'"
            Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
              If rs1.RecordCount <> 0 Then
                 rs1.MoveFirst
                 Do Until rs1.EOF
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 If rs1!Type = "Paid" Then
                   Sqlqry2 = " Insert into dumbo_tramag values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                             & Val(Trim(rs1!tra_amount)) & "," _
                                             & Val(Trim(rs1!tra_amount)) & ",' " _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 Else
                      Sqlqry2 = " Insert into dumbo_tramag values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "', " & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0','" _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
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
                  
       
      
    If CboMedia.Text = "Cinema" Then
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
     CrystalReport1.ReportFileName = App.Path & "\bocininv.rpt"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.DetailCopies = 1
     CrystalReport1.Action = 1
     

    ElseIf CboMedia.Text = "Magazine" Then
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
     CrystalReport1.ReportFileName = App.Path & "\bomaginv.rpt"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Action = 1
     
     
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
     CrystalReport1.ReportFileName = App.Path & "\bomaginvext.rpt"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Action = 1
     
    ElseIf Mid(CboMedia.Text, 1, 3) = "Mag" Then
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
     CrystalReport1.ReportFileName = App.Path & "\bomaginv.rpt"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Action = 1
     
     
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
     CrystalReport1.ReportFileName = App.Path & "\bomaginvext.rpt"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Action = 1
     
    ElseIf CboMedia.Text = "Television" Then
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
     CrystalReport1.ReportFileName = App.Path & "\botelinv1.rpt"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Action = 1
   ElseIf CboMedia.Text = "Online" Then
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
     CrystalReport1.ReportFileName = App.Path & "\boolinv.rpt"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Action = 1
   End If
     CboMYear.Text = Year(Now())
     CboMMonth.ListIndex = Month(Now) - 1
     CboMedia.ListIndex = -1

End Sub

Private Sub PopulateSerialInv()
If ValidateDataserial = False Then Exit Sub


 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from invrep"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans

 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tracin"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tramag"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
 
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tramagext"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_traol"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tratv"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
X = MsgBox("Do you want to print invoices from the serial No. " & CboSerialFrom.Text & ", To " & CboSerialTo.Text, vbInformation, "Confirmation")
   
If X = vbNo Then Exit Sub

Dim i As Integer
Dim j As Integer
    
    Sqlqry1 = "select * from bo_mas Where STATUS='N' AND serial_no>='" & CboSerialFrom & "' and serial_no<='" & CboSerialTo & "'  order by media"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        
        If rs.RecordCount <> 0 Then
         rs.MoveFirst
         Do Until rs.EOF
         
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         Sqlqry = " Insert into invrep values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "',' " _
                                     & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" & Trim(rs!tcurrency) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                     & Val(rs!tra_gamount) & "," _
                                     & Val(rs!Tot_free) & "," _
                                     & Val(rs!Tot_barter) & ",'" _
                                     & Val(Trim(rs!disc_percentage)) & "','" _
                                     & Val(Trim(rs!disc_rate)) & "'," _
                                     & Val(Trim(rs!add_discount)) & "," _
                                     & Val(Trim(rs!surcharge)) & "," _
                                     & Val(Trim(rs!tra_namount)) & ",'" & Format(rs!invoice_date, "dd/mm/yyyy") & "')"
                                     
                                     
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
            
            If rs!Media = "Cinema" Then
            
                Sqlqry1 = "Select * from bo_tracin where serial_no ='" & Trim(rs!serial_no) & "'"
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                If rs1.RecordCount <> 0 Then
                   
                   rs1.MoveFirst
                   Do Until rs1.EOF
                   Set ws = DBEngine.Workspaces(0)
                   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                     If rs1!Type = "Paid" Then
                         Sqlqry2 = " Insert into dumbo_tracin values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                               & Trim(rs1!Month) & "','" _
                                               & findfirstfixup(rs1!Product) & "','" _
                                               & findfirstfixup(rs1!client) & "','" _
                                               & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                               & Trim(rs1!sub_Media) & "','" _
                                               & Trim(rs1!DATEFROM) & "','" _
                                               & Trim(rs1!Dateto) & "','" _
                                               & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                               & Trim(rs1!Day) & "','" _
                                               & Trim(rs1!Length) & "','" _
                                               & findfirstfixup(Trim(rs1!Description)) & "','" _
                                               & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                               & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                               & Val(Trim(rs1!tra_amount)) & ", " _
                                               & Val(Trim(rs1!tra_amount)) & " )"
                                               
                      Else
                          Sqlqry2 = " Insert into dumbo_tracin values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                               & Trim(rs1!Month) & "','" _
                                               & findfirstfixup(rs1!Product) & "','" _
                                               & findfirstfixup(rs1!client) & "','" _
                                               & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                               & Trim(rs1!sub_Media) & "','" _
                                               & Trim(rs1!DATEFROM) & "','" _
                                               & Trim(rs1!Dateto) & "','" _
                                               & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                               & Trim(rs1!Day) & "','" _
                                               & Trim(rs1!Length) & "','" _
                                               & findfirstfixup(Trim(rs1!Description)) & "','" _
                                               & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                               & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0', '0' )"
                       End If
                       
                       
                      ws.BeginTrans
                      db.Execute (Sqlqry2)
                      ws.CommitTrans
                    rs1.MoveNext
                   Loop
                  End If
               ElseIf rs!Media = "Online" Then
               
                  Sqlqry1 = "Select * from bo_traol where serial_no ='" & Trim(rs!serial_no) & "'"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                      If rs1.RecordCount <> 0 Then
                         
                         rs1.MoveFirst
                         Do Until rs1.EOF
                         Set ws = DBEngine.Workspaces(0)
                         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         If rs1!Type = "Paid" Then
                           Sqlqry2 = " Insert into dumBo_traol values('" & rs1!serial_no & "','" & Trim(rs1!date_From) & "','" & Trim(rs1!DATE_TO) & "','" & rs1!Year & "','" _
                                                     & Trim(rs1!Month) & "','" _
                                                     & findfirstfixup(rs1!Product) & "','" _
                                                     & findfirstfixup(rs1!client) & "','" _
                                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                     & Trim(rs1!sub_Media) & "','" _
                                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                                     & Trim(rs1!Type) & "'," _
                                                     & Val(Trim(rs1!impressions)) & "," _
                                                     & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                                     & Val(Trim(rs1!tra_amount)) & "," _
                                                     & Val(Trim(rs1!tra_amount)) & ")"
                                                     
                           Else
                             Sqlqry2 = " Insert into dumBo_traol values('" & rs1!serial_no & "','" & Trim(rs1!date_From) & "','" & Trim(rs1!DATE_TO) & "','" & rs1!Year & "','" _
                                                     & Trim(rs1!Month) & "','" _
                                                     & findfirstfixup(rs1!Product) & "','" _
                                                     & findfirstfixup(rs1!client) & "','" _
                                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                     & Trim(rs1!sub_Media) & "','" _
                                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                                     & Trim(rs1!Type) & "'," _
                                                     & Val(Trim(rs1!impressions)) & "," _
                                                     & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0')"
                           End If
                                                     
                            ws.BeginTrans
                            db.Execute (Sqlqry2)
                            ws.CommitTrans
                          rs1.MoveNext
                         Loop
                       End If
         ElseIf rs!Media = "Television" Then
            Sqlqry1 = "Select * from bo_traTv where serial_no ='" & Trim(rs!serial_no) & "'"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                  
                  rs1.MoveFirst
                  Do Until rs1.EOF
                  Set ws = DBEngine.Workspaces(0)
                  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                   If rs1!Type = "Paid" Then
                    Sqlqry2 = " Insert into dumbo_tratv values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                              & Trim(rs1!Month) & "','" _
                                              & findfirstfixup(rs1!Product) & "','" _
                                              & findfirstfixup(rs1!client) & "','" _
                                              & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                              & Trim(rs1!sub_Media) & "','" _
                                              & findfirstfixup(Trim(rs1!bo_ref)) & "','" & Trim(rs1!code) & "','" & Trim(rs1!sec) & "','" _
                                              & Trim(rs1!Day) & "','" _
                                              & Trim(rs1!Time) & "','" _
                                              & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                              & Trim(rs1!Type) & "'," _
                                              & Val(Trim(rs1!spots)) & "," _
                                              & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                              & Val(Trim(rs1!tra_amount)) & "," _
                                              & Val(Trim(rs1!tra_amount)) & ")"
                    Else
                      Sqlqry2 = " Insert into dumbo_tratv values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                              & Trim(rs1!Month) & "','" _
                                              & findfirstfixup(rs1!Product) & "','" _
                                              & findfirstfixup(rs1!client) & "','" _
                                              & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                              & Trim(rs1!sub_Media) & "','" _
                                              & findfirstfixup(Trim(rs1!bo_ref)) & "','" & Trim(rs1!code) & "','" & Trim(rs1!sec) & "','" _
                                              & Trim(rs1!Day) & "','" _
                                              & Trim(rs1!Time) & "','" _
                                              & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                              & Trim(rs1!Type) & "'," _
                                              & Val(Trim(rs1!spots)) & "," _
                                              & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0')"
                    End If
                                              
                     ws.BeginTrans
                     db.Execute (Sqlqry2)
                     ws.CommitTrans
                   rs1.MoveNext
                  Loop
                End If
           ElseIf rs!Media = "Magazine" And rs!sub_Media = "ALAM ASSAYARRAT" Or rs!sub_Media = "ZEINA" Then
             Sqlqry1 = "Select * from bo_tramag where serial_no ='" & Trim(rs!serial_no) & "'"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
              If rs1.RecordCount <> 0 Then
                 
                 rs1.MoveFirst
                 Do Until rs1.EOF
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 If rs1!Type = "Paid" Then
                   Sqlqry2 = " Insert into dumbo_tramagext values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                             & Val(Trim(rs1!tra_amount)) & "," _
                                             & Val(Trim(rs1!tra_amount)) & ",' " _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 Else
                      Sqlqry2 = " Insert into dumbo_tramagext values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "', " & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0','" _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 End If
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
                  rs1.MoveNext
                 Loop
               End If
           Else
            Sqlqry1 = "Select * from bo_tramag where serial_no ='" & Trim(rs!serial_no) & "'"
            Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
              If rs1.RecordCount <> 0 Then
                 
                 rs1.MoveFirst
                 Do Until rs1.EOF
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 If rs1!Type = "Paid" Then
                   Sqlqry2 = " Insert into dumbo_tramag values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                             & Val(Trim(rs1!tra_amount)) & "," _
                                             & Val(Trim(rs1!tra_amount)) & ",' " _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 Else
                      Sqlqry2 = " Insert into dumbo_tramag values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "', " & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0','" _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 End If
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
                  rs1.MoveNext
                 Loop
               End If
           End If
                  
       rs.MoveNext
       Loop
      End If
      
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_traCin"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\bocininv.rpt"
        CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.DetailCopies = 1
        CrystalReport1.Action = 1
     End If
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_tramag"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
         CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
         CrystalReport1.ReportFileName = App.Path & "\bomaginv.rpt"
         CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
         CrystalReport1.WindowState = crptMaximized
         CrystalReport1.Action = 1
     End If
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_tramagext"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
         CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
         CrystalReport1.ReportFileName = App.Path & "\bomaginvext.rpt"
         CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
         CrystalReport1.WindowState = crptMaximized
         CrystalReport1.Action = 1
     End If
     
     
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_tratv"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
          CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
          CrystalReport1.ReportFileName = App.Path & "\botelinv1.rpt"
          CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
          CrystalReport1.WindowState = crptMaximized
          CrystalReport1.Action = 1
     End If
        
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_traol"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
          CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
          CrystalReport1.ReportFileName = App.Path & "\boolinv.rpt"
          CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
          CrystalReport1.WindowState = crptMaximized
          CrystalReport1.Action = 1
     End If
        
     CboSerialFrom.ListIndex = -1
     CboSerialTo.ListIndex = -1

End Sub

Private Sub populateParticularno()

If ValidateDataParticular = False Then Exit Sub


 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from invrep"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans

 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tracin"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tramag"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
 
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tramagext"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_traol"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tratv"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
'X = MsgBox("Do you want to print invoices from the serial No. " & CboSerialFrom.Text & ", To " & CboSerialTo.Text, vbInformation, "Confirmation")
 X = MsgBox("Do you want to print invoice for the serial No. " & CboParticularNo.Text & ", vbInformation, Confirmation")
If X = vbNo Then Exit Sub

Dim i As Integer
Dim j As Integer
    
    Sqlqry1 = "select * from bo_mas Where STATUS='N' AND serial_no='" & CboParticularNo & "' order by media"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        
        If rs.RecordCount <> 0 Then
         rs.MoveFirst
         Do Until rs.EOF
         
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         Sqlqry = " Insert into invrep values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "',' " _
                                     & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" & Trim(rs!tcurrency) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                     & Val(rs!tra_gamount) & "," _
                                     & Val(rs!Tot_free) & "," _
                                     & Val(rs!Tot_barter) & ",'" _
                                     & Val(Trim(rs!disc_percentage)) & "','" _
                                     & Val(Trim(rs!disc_rate)) & "'," _
                                     & Val(Trim(rs!add_discount)) & "," _
                                     & Val(Trim(rs!surcharge)) & "," _
                                     & Val(Trim(rs!tra_namount)) & ",'" & Format(rs!invoice_date, "dd/mm/yyyy") & "')"
                                     
                                     
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
            
            If rs!Media = "Cinema" Then
            
                Sqlqry1 = "Select * from bo_tracin where serial_no ='" & Trim(rs!serial_no) & "'"
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                If rs1.RecordCount <> 0 Then
                   
                   rs1.MoveFirst
                   Do Until rs1.EOF
                   Set ws = DBEngine.Workspaces(0)
                   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                     If rs1!Type = "Paid" Then
                         Sqlqry2 = " Insert into dumbo_tracin values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                               & Trim(rs1!Month) & "','" _
                                               & findfirstfixup(rs1!Product) & "','" _
                                               & findfirstfixup(rs1!client) & "','" _
                                               & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                               & Trim(rs1!sub_Media) & "','" _
                                               & Trim(rs1!DATEFROM) & "','" _
                                               & Trim(rs1!Dateto) & "','" _
                                               & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                               & Trim(rs1!Day) & "','" _
                                               & Trim(rs1!Length) & "','" _
                                               & findfirstfixup(Trim(rs1!Description)) & "','" _
                                               & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                               & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                               & Val(Trim(rs1!tra_amount)) & ", " _
                                               & Val(Trim(rs1!tra_amount)) & " )"
                                               
                      Else
                          Sqlqry2 = " Insert into dumbo_tracin values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                               & Trim(rs1!Month) & "','" _
                                               & findfirstfixup(rs1!Product) & "','" _
                                               & findfirstfixup(rs1!client) & "','" _
                                               & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                               & Trim(rs1!sub_Media) & "','" _
                                               & Trim(rs1!DATEFROM) & "','" _
                                               & Trim(rs1!Dateto) & "','" _
                                               & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                               & Trim(rs1!Day) & "','" _
                                               & Trim(rs1!Length) & "','" _
                                               & findfirstfixup(Trim(rs1!Description)) & "','" _
                                               & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                               & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0', '0' )"
                       End If
                       
                       
                      ws.BeginTrans
                      db.Execute (Sqlqry2)
                      ws.CommitTrans
                    rs1.MoveNext
                   Loop
                  End If
               ElseIf rs!Media = "Online" Then
               
                  Sqlqry1 = "Select * from bo_traol where serial_no ='" & Trim(rs!serial_no) & "'"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                      If rs1.RecordCount <> 0 Then
                         
                         rs1.MoveFirst
                         Do Until rs1.EOF
                         Set ws = DBEngine.Workspaces(0)
                         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         If rs1!Type = "Paid" Then
                           Sqlqry2 = " Insert into dumBo_traol values('" & rs1!serial_no & "','" & Trim(rs1!date_From) & "','" & Trim(rs1!DATE_TO) & "','" & rs1!Year & "','" _
                                                     & Trim(rs1!Month) & "','" _
                                                     & findfirstfixup(rs1!Product) & "','" _
                                                     & findfirstfixup(rs1!client) & "','" _
                                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                     & Trim(rs1!sub_Media) & "','" _
                                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                                     & Trim(rs1!Type) & "'," _
                                                     & Val(Trim(rs1!impressions)) & "," _
                                                     & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                                     & Val(Trim(rs1!tra_amount)) & "," _
                                                     & Val(Trim(rs1!tra_amount)) & ")"
                                                     
                           Else
                             Sqlqry2 = " Insert into dumBo_traol values('" & rs1!serial_no & "','" & Trim(rs1!date_From) & "','" & Trim(rs1!DATE_TO) & "','" & rs1!Year & "','" _
                                                     & Trim(rs1!Month) & "','" _
                                                     & findfirstfixup(rs1!Product) & "','" _
                                                     & findfirstfixup(rs1!client) & "','" _
                                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                     & Trim(rs1!sub_Media) & "','" _
                                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                                     & Trim(rs1!Type) & "'," _
                                                     & Val(Trim(rs1!impressions)) & "," _
                                                     & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0')"
                           End If
                                                     
                            ws.BeginTrans
                            db.Execute (Sqlqry2)
                            ws.CommitTrans
                          rs1.MoveNext
                         Loop
                       End If
         ElseIf rs!Media = "Television" Then
            Sqlqry1 = "Select * from bo_traTv where serial_no ='" & Trim(rs!serial_no) & "'"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                  
                  rs1.MoveFirst
                  Do Until rs1.EOF
                  Set ws = DBEngine.Workspaces(0)
                  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                   If rs1!Type = "Paid" Then
                    Sqlqry2 = " Insert into dumbo_tratv values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                              & Trim(rs1!Month) & "','" _
                                              & findfirstfixup(rs1!Product) & "','" _
                                              & findfirstfixup(rs1!client) & "','" _
                                              & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                              & Trim(rs1!sub_Media) & "','" _
                                              & findfirstfixup(Trim(rs1!bo_ref)) & "','" & Trim(rs1!code) & "','" & Trim(rs1!sec) & "','" _
                                              & Trim(rs1!Day) & "','" _
                                              & Trim(rs1!Time) & "','" _
                                              & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                              & Trim(rs1!Type) & "'," _
                                              & Val(Trim(rs1!spots)) & "," _
                                              & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                              & Val(Trim(rs1!tra_amount)) & "," _
                                              & Val(Trim(rs1!tra_amount)) & ")"
                    Else
                      Sqlqry2 = " Insert into dumbo_tratv values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                              & Trim(rs1!Month) & "','" _
                                              & findfirstfixup(rs1!Product) & "','" _
                                              & findfirstfixup(rs1!client) & "','" _
                                              & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                              & Trim(rs1!sub_Media) & "','" _
                                              & findfirstfixup(Trim(rs1!bo_ref)) & "','" & Trim(rs1!code) & "','" & Trim(rs1!sec) & "','" _
                                              & Trim(rs1!Day) & "','" _
                                              & Trim(rs1!Time) & "','" _
                                              & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                              & Trim(rs1!Type) & "'," _
                                              & Val(Trim(rs1!spots)) & "," _
                                              & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0')"
                    End If
                                              
                     ws.BeginTrans
                     db.Execute (Sqlqry2)
                     ws.CommitTrans
                   rs1.MoveNext
                  Loop
                End If
           ElseIf rs!Media = "Magazine" And rs!sub_Media = "ALAM ASSAYARRAT" Or rs!sub_Media = "ZEINA" Then
             Sqlqry1 = "Select * from bo_tramag where serial_no ='" & Trim(rs!serial_no) & "'"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
              If rs1.RecordCount <> 0 Then
                 
                 rs1.MoveFirst
                 Do Until rs1.EOF
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 If rs1!Type = "Paid" Then
                   Sqlqry2 = " Insert into dumbo_tramagext values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                             & Val(Trim(rs1!tra_amount)) & "," _
                                             & Val(Trim(rs1!tra_amount)) & ",' " _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 Else
                      Sqlqry2 = " Insert into dumbo_tramagext values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "', " & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0','" _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 End If
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
                  rs1.MoveNext
                 Loop
               End If
           Else
            Sqlqry1 = "Select * from bo_tramag where serial_no ='" & Trim(rs!serial_no) & "'"
            Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
              If rs1.RecordCount <> 0 Then
                 
                 rs1.MoveFirst
                 Do Until rs1.EOF
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 If rs1!Type = "Paid" Then
                   Sqlqry2 = " Insert into dumbo_tramag values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                             & Val(Trim(rs1!tra_amount)) & "," _
                                             & Val(Trim(rs1!tra_amount)) & ",' " _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 Else
                      Sqlqry2 = " Insert into dumbo_tramag values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "', " & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0','" _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 End If
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
                  rs1.MoveNext
                 Loop
               End If
           End If
                  
       rs.MoveNext
       Loop
      End If
      
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_traCin"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\bocininv2.rpt"
        CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.DetailCopies = 1
        CrystalReport1.Action = 1
     End If
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_tramag"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
         CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
         CrystalReport1.ReportFileName = App.Path & "\bomaginv2.rpt"
         CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
         CrystalReport1.WindowState = crptMaximized
         CrystalReport1.Action = 1
     End If
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_tramagext"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
         CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
         CrystalReport1.ReportFileName = App.Path & "\bomaginvext2.rpt"
         CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
         CrystalReport1.WindowState = crptMaximized
         CrystalReport1.Action = 1
     End If
     
     
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_tratv"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
          CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
          CrystalReport1.ReportFileName = App.Path & "\botelinv2.rpt"
          CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
          CrystalReport1.WindowState = crptMaximized
          CrystalReport1.Action = 1
     End If
        
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_traol"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
          CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
          CrystalReport1.ReportFileName = App.Path & "\boolinv2.rpt"
          CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
          CrystalReport1.WindowState = crptMaximized
          CrystalReport1.Action = 1
     End If
        
     CboParticularNo.ListIndex = -1
     


End Sub
Private Sub cmdWork_Click()
Dim X

If OptMonth.Value = True Then
  Populatemonthinv
ElseIf Optmedia.Value = True Then
  populatemediainv
ElseIf OptSerialNo.Value = True Then
 PopulateSerialInv
ElseIf OptAgency.Value = True Then
 populateagencyinv
ElseIf OptProduct.Value = True Then
 populateproductinv
ElseIf OptParticularno.Value = True Then
 populateParticularno
End If
           
               
End Sub
Private Sub populateproductinv()

If validateDataProduct = False Then Exit Sub


 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from invrep"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans

 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tracin"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tramag"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tramagext"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
 
 
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_traol"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tratv"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans

X = MsgBox("Do you want to print invoices for the month of " & CboPMonth.Text & ", Year " & CboPYear.Text & ", Product " & CboProduct.Text, vbInformation + vbYesNo, "Confirmation")
   
If X = vbNo Then Exit Sub

Dim i As Integer
Dim j As Integer
    j = CboPMonth.ListIndex + 1
    FirstDate = Now
    LastDate = Now
    i = DaysinMonth(j, CboPYear.Text)
    FirstDate = Format(DateValue("1-" & j & "-" & CboPYear), "dd/mm/yyyy")
    LastDate = Format(DateValue(i & "-" & j & "-" & CboPYear), "DD/mm/yyyy")
    
    Sqlqry1 = "select * from bo_mas Where STATUS='N' AND Product='" & CboProduct.Text & "' and invoice_Date Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "# order by media"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        
        If rs.RecordCount <> 0 Then
         rs.MoveFirst
         Do Until rs.EOF
         
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         Sqlqry = " Insert into invrep values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "',' " _
                                     & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" & Trim(rs!tcurrency) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                     & Val(rs!tra_gamount) & "," _
                                     & Val(rs!Tot_free) & "," _
                                     & Val(rs!Tot_barter) & ",'" _
                                     & Val(Trim(rs!disc_percentage)) & "','" _
                                     & Val(Trim(rs!disc_rate)) & "'," _
                                     & Val(Trim(rs!add_discount)) & "," _
                                     & Val(Trim(rs!surcharge)) & "," _
                                     & Val(Trim(rs!tra_namount)) & ",'" & Format(rs!invoice_date, "dd/mm/yyyy") & "')"
                                     
                                     
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
            
            If rs!Media = "Cinema" Then
            
                Sqlqry1 = "Select * from bo_tracin where serial_no ='" & Trim(rs!serial_no) & "'"
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                If rs1.RecordCount <> 0 Then
                   
                   rs1.MoveFirst
                   Do Until rs1.EOF
                   Set ws = DBEngine.Workspaces(0)
                   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                     If rs1!Type = "Paid" Then
                         Sqlqry2 = " Insert into dumbo_tracin values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                               & Trim(rs1!Month) & "','" _
                                               & findfirstfixup(rs1!Product) & "','" _
                                               & findfirstfixup(rs1!client) & "','" _
                                               & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                               & Trim(rs1!sub_Media) & "','" _
                                               & Trim(rs1!DATEFROM) & "','" _
                                               & Trim(rs1!Dateto) & "','" _
                                               & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                               & Trim(rs1!Day) & "','" _
                                               & Trim(rs1!Length) & "','" _
                                               & findfirstfixup(Trim(rs1!Description)) & "','" _
                                               & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                               & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                               & Val(Trim(rs1!tra_amount)) & ", " _
                                               & Val(Trim(rs1!tra_amount)) & " )"
                                               
                      Else
                          Sqlqry2 = " Insert into dumbo_tracin values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                               & Trim(rs1!Month) & "','" _
                                               & findfirstfixup(rs1!Product) & "','" _
                                               & findfirstfixup(rs1!client) & "','" _
                                               & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                               & Trim(rs1!sub_Media) & "','" _
                                               & Trim(rs1!DATEFROM) & "','" _
                                               & Trim(rs1!Dateto) & "','" _
                                               & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                               & Trim(rs1!Day) & "','" _
                                               & Trim(rs1!Length) & "','" _
                                               & findfirstfixup(Trim(rs1!Description)) & "','" _
                                               & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                               & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0', '0' )"
                       End If
                       
                       
                      ws.BeginTrans
                      db.Execute (Sqlqry2)
                      ws.CommitTrans
                    rs1.MoveNext
                   Loop
                  End If
               ElseIf rs!Media = "Online" Then
               
                  Sqlqry1 = "Select * from bo_traol where serial_no ='" & Trim(rs!serial_no) & "'"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                      If rs1.RecordCount <> 0 Then
                         
                         rs1.MoveFirst
                         Do Until rs1.EOF
                         Set ws = DBEngine.Workspaces(0)
                         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         If rs1!Type = "Paid" Then
                           Sqlqry2 = " Insert into dumBo_traol values('" & rs1!serial_no & "','" & Trim(rs1!date_From) & "','" & Trim(rs1!DATE_TO) & "','" & rs1!Year & "','" _
                                                     & Trim(rs1!Month) & "','" _
                                                     & findfirstfixup(rs1!Product) & "','" _
                                                     & findfirstfixup(rs1!client) & "','" _
                                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                     & Trim(rs1!sub_Media) & "','" _
                                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                                     & Trim(rs1!Type) & "'," _
                                                     & Val(Trim(rs1!impressions)) & "," _
                                                     & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                                     & Val(Trim(rs1!tra_amount)) & "," _
                                                     & Val(Trim(rs1!tra_amount)) & ")"
                                                     
                           Else
                             Sqlqry2 = " Insert into dumBo_traol values('" & rs1!serial_no & "','" & Trim(rs1!date_From) & "','" & Trim(rs1!DATE_TO) & "','" & rs1!Year & "','" _
                                                     & Trim(rs1!Month) & "','" _
                                                     & findfirstfixup(rs1!Product) & "','" _
                                                     & findfirstfixup(rs1!client) & "','" _
                                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                     & Trim(rs1!sub_Media) & "','" _
                                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                                     & Trim(rs1!Type) & "'," _
                                                     & Val(Trim(rs1!impressions)) & "," _
                                                     & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0')"
                           End If
                                                     
                            ws.BeginTrans
                            db.Execute (Sqlqry2)
                            ws.CommitTrans
                          rs1.MoveNext
                         Loop
                       End If
         ElseIf rs!Media = "Television" Then
            Sqlqry1 = "Select * from bo_traTv where serial_no ='" & Trim(rs!serial_no) & "'"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                  
                  rs1.MoveFirst
                  Do Until rs1.EOF
                  Set ws = DBEngine.Workspaces(0)
                  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                   If rs1!Type = "Paid" Then
                    Sqlqry2 = " Insert into dumbo_tratv values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                              & Trim(rs1!Month) & "','" _
                                              & findfirstfixup(rs1!Product) & "','" _
                                              & findfirstfixup(rs1!client) & "','" _
                                              & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                              & Trim(rs1!sub_Media) & "','" _
                                              & findfirstfixup(Trim(rs1!bo_ref)) & "','" & Trim(rs1!code) & "','" & Trim(rs1!sec) & "','" _
                                              & Trim(rs1!Day) & "','" _
                                              & Trim(rs1!Time) & "','" _
                                              & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                              & Trim(rs1!Type) & "'," _
                                              & Val(Trim(rs1!spots)) & "," _
                                              & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                              & Val(Trim(rs1!tra_amount)) & "," _
                                              & Val(Trim(rs1!tra_amount)) & ")"
                    Else
                      Sqlqry2 = " Insert into dumbo_tratv values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                              & Trim(rs1!Month) & "','" _
                                              & findfirstfixup(rs1!Product) & "','" _
                                              & findfirstfixup(rs1!client) & "','" _
                                              & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                              & Trim(rs1!sub_Media) & "','" _
                                              & findfirstfixup(Trim(rs1!bo_ref)) & "','" & Trim(rs1!code) & "','" & Trim(rs1!sec) & "','" _
                                              & Trim(rs1!Day) & "','" _
                                              & Trim(rs1!Time) & "','" _
                                              & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                              & Trim(rs1!Type) & "'," _
                                              & Val(Trim(rs1!spots)) & "," _
                                              & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0')"
                    End If
                                              
                     ws.BeginTrans
                     db.Execute (Sqlqry2)
                     ws.CommitTrans
                   rs1.MoveNext
                  Loop
                End If
           ElseIf rs!Media = "Magazine" And rs!sub_Media = "ALAM ASSAYARRAT" Or rs!sub_Media = "ZEINA" Then
              Sqlqry1 = "Select * from bo_tramag where serial_no ='" & Trim(rs!serial_no) & "'"
              Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
              If rs1.RecordCount <> 0 Then
                 
                 rs1.MoveFirst
                 Do Until rs1.EOF
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 If rs1!Type = "Paid" Then
                   Sqlqry2 = " Insert into dumbo_tramagext values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                             & Val(Trim(rs1!tra_amount)) & "," _
                                             & Val(Trim(rs1!tra_amount)) & ",' " _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 Else
                      Sqlqry2 = " Insert into dumbo_tramagext values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "', " & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0','" _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 End If
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
                  rs1.MoveNext
                 Loop
               End If
           Else
            Sqlqry1 = "Select * from bo_tramag where serial_no ='" & Trim(rs!serial_no) & "'"
            Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
              If rs1.RecordCount <> 0 Then
                 
                 rs1.MoveFirst
                 Do Until rs1.EOF
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 If rs1!Type = "Paid" Then
                   Sqlqry2 = " Insert into dumbo_tramag values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                             & Val(Trim(rs1!tra_amount)) & "," _
                                             & Val(Trim(rs1!tra_amount)) & ",' " _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 Else
                      Sqlqry2 = " Insert into dumbo_tramag values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "', " & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0','" _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 End If
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
                  rs1.MoveNext
                 Loop
               End If
           End If
                  
       rs.MoveNext
       Loop
      End If
      
          Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_traCin"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\bocininv.rpt"
        CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.DetailCopies = 1
        CrystalReport1.Action = 1
     End If
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_tramag"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
         CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
         CrystalReport1.ReportFileName = App.Path & "\bomaginv.rpt"
         CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
         CrystalReport1.WindowState = crptMaximized
         CrystalReport1.Action = 1
     End If
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_tramagext"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
         CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
         CrystalReport1.ReportFileName = App.Path & "\bomaginvext.rpt"
         CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
         CrystalReport1.WindowState = crptMaximized
         CrystalReport1.Action = 1
     End If
     
     
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_tratv"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
          CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
          CrystalReport1.ReportFileName = App.Path & "\botelinv1.rpt"
          CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
          CrystalReport1.WindowState = crptMaximized
          CrystalReport1.Action = 1
     End If
        
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_traol"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
          CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
          CrystalReport1.ReportFileName = App.Path & "\boolinv.rpt"
          CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
          CrystalReport1.WindowState = crptMaximized
          CrystalReport1.Action = 1
     End If

     CboPYear.Text = Year(Now())
     CboPMonth.ListIndex = Month(Now) - 1
     CboProduct.ListIndex = -1

End Sub
Private Sub populateagencyinv()

If validateDataAgency = False Then Exit Sub


 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from invrep"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans

 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tracin"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tramag"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_traol"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans
    
 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from dumbo_tratv"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans

X = MsgBox("Do you want to print invoices for the month of " & CboAmonth.Text & ", Year " & CboAyear.Text & ", Agency " & CboAgency.Text, vbInformation + vbYesNo, "Confirmation")
   
If X = vbNo Then Exit Sub

Dim i As Integer
Dim j As Integer
    j = CboAmonth.ListIndex + 1
    FirstDate = Now
    LastDate = Now
    i = DaysinMonth(j, CboAyear.Text)
    FirstDate = Format(DateValue("1-" & j & "-" & CboAyear), "dd/mm/yyyy")
    LastDate = Format(DateValue(i & "-" & j & "-" & CboAyear), "DD/mm/yyyy")
    
    Sqlqry1 = "select * from bo_mas Where STATUS='N' AND Agency='" & CboAgency.Text & "' and invoice_Date Between #" & Format(FirstDate, "mm-dd-yyyy") & "# and #" & Format(LastDate, "mm-dd-yyyy") & "# order by media"
    Set rs = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
        
        If rs.RecordCount <> 0 Then
         rs.MoveFirst
         Do Until rs.EOF
         
         Set ws = DBEngine.Workspaces(0)
         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
         Sqlqry = " Insert into invrep values('" & Trim(rs!serial_no) & "','" & Trim(rs!Year) & "',' " _
                                     & Trim(rs!Month) & " '," & Val(rs!monthind) & ",'" _
                                     & findfirstfixup(rs!Product) & "','" _
                                     & findfirstfixup(rs!client) & "','" _
                                     & findfirstfixup(rs!Agency) & "','" & Trim(rs!Media) & "','" _
                                     & Trim(rs!sub_Media) & "','" & Trim(rs!tcurrency) & "','" _
                                     & findfirstfixup(Trim(rs!bo_ref)) & "'," _
                                     & Val(rs!tra_gamount) & "," _
                                     & Val(rs!Tot_free) & "," _
                                     & Val(rs!Tot_barter) & ",'" _
                                     & Val(Trim(rs!disc_percentage)) & "','" _
                                     & Val(Trim(rs!disc_rate)) & "'," _
                                     & Val(Trim(rs!add_discount)) & "," _
                                     & Val(Trim(rs!surcharge)) & "," _
                                     & Val(Trim(rs!tra_namount)) & ",'" & Format(rs!invoice_date, "dd/mm/yyyy") & "')"
                                     
                                     
            ws.BeginTrans
            db.Execute (Sqlqry)
            ws.CommitTrans
            
            If rs!Media = "Cinema" Then
            
                Sqlqry1 = "Select * from bo_tracin where serial_no ='" & Trim(rs!serial_no) & "'"
                Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                If rs1.RecordCount <> 0 Then
                   
                   rs1.MoveFirst
                   Do Until rs1.EOF
                   Set ws = DBEngine.Workspaces(0)
                   Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                     If rs1!Type = "Paid" Then
                         Sqlqry2 = " Insert into dumbo_tracin values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                               & Trim(rs1!Month) & "','" _
                                               & findfirstfixup(rs1!Product) & "','" _
                                               & findfirstfixup(rs1!client) & "','" _
                                               & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                               & Trim(rs1!sub_Media) & "','" _
                                               & Trim(rs1!DATEFROM) & "','" _
                                               & Trim(rs1!Dateto) & "','" _
                                               & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                               & Trim(rs1!Day) & "','" _
                                               & Trim(rs1!Length) & "','" _
                                               & findfirstfixup(Trim(rs1!Description)) & "','" _
                                               & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                               & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                               & Val(Trim(rs1!tra_amount)) & ", " _
                                               & Val(Trim(rs1!tra_amount)) & " )"
                                               
                      Else
                          Sqlqry2 = " Insert into dumbo_tracin values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                               & Trim(rs1!Month) & "','" _
                                               & findfirstfixup(rs1!Product) & "','" _
                                               & findfirstfixup(rs1!client) & "','" _
                                               & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                               & Trim(rs1!sub_Media) & "','" _
                                                & Trim(rs1!DATEFROM) & "','" _
                                               & Trim(rs1!Dateto) & "','" _
                                               & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                               & Trim(rs1!Day) & "','" _
                                               & Trim(rs1!Length) & "','" _
                                               & findfirstfixup(Trim(rs1!Description)) & "','" _
                                               & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                               & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0', '0' )"
                       End If
                       
                       
                      ws.BeginTrans
                      db.Execute (Sqlqry2)
                      ws.CommitTrans
                    rs1.MoveNext
                   Loop
                  End If
               ElseIf rs!Media = "Online" Then
               
                  Sqlqry1 = "Select * from bo_traol where serial_no ='" & Trim(rs!serial_no) & "'"
                    Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
                      If rs1.RecordCount <> 0 Then
                         
                         rs1.MoveFirst
                         Do Until rs1.EOF
                         Set ws = DBEngine.Workspaces(0)
                         Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                         If rs1!Type = "Paid" Then
                           Sqlqry2 = " Insert into dumBo_traol values('" & rs1!serial_no & "','" & Trim(rs1!date_From) & "','" & Trim(rs1!DATE_TO) & "','" & rs1!Year & "','" _
                                                     & Trim(rs1!Month) & "','" _
                                                     & findfirstfixup(rs1!Product) & "','" _
                                                     & findfirstfixup(rs1!client) & "','" _
                                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                     & Trim(rs1!sub_Media) & "','" _
                                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                                     & Trim(rs1!Type) & "'," _
                                                     & Val(Trim(rs1!impressions)) & "," _
                                                     & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                                     & Val(Trim(rs1!tra_amount)) & "," _
                                                     & Val(Trim(rs1!tra_amount)) & ")"
                                                     
                           Else
                             Sqlqry2 = " Insert into dumBo_traol values('" & rs1!serial_no & "','" & Trim(rs1!date_From) & "','" & Trim(rs1!DATE_TO) & "','" & rs1!Year & "','" _
                                                     & Trim(rs1!Month) & "','" _
                                                     & findfirstfixup(rs1!Product) & "','" _
                                                     & findfirstfixup(rs1!client) & "','" _
                                                     & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                                     & Trim(rs1!sub_Media) & "','" _
                                                     & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                                     & findfirstfixup(Trim(rs1!Description)) & "','" _
                                                     & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                                     & Trim(rs1!Type) & "'," _
                                                     & Val(Trim(rs1!impressions)) & "," _
                                                     & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0')"
                           End If
                                                     
                            ws.BeginTrans
                            db.Execute (Sqlqry2)
                            ws.CommitTrans
                          rs1.MoveNext
                         Loop
                       End If
         ElseIf rs!Media = "Television" Then
            Sqlqry1 = "Select * from bo_traTv where serial_no ='" & Trim(rs!serial_no) & "'"
             Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
               If rs1.RecordCount <> 0 Then
                  
                  rs1.MoveFirst
                  Do Until rs1.EOF
                  Set ws = DBEngine.Workspaces(0)
                  Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                   If rs1!Type = "Paid" Then
                    Sqlqry2 = " Insert into dumbo_tratv values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                              & Trim(rs1!Month) & "','" _
                                              & findfirstfixup(rs1!Product) & "','" _
                                              & findfirstfixup(rs1!client) & "','" _
                                              & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                              & Trim(rs1!sub_Media) & "','" _
                                              & findfirstfixup(Trim(rs1!bo_ref)) & "','" & Trim(rs1!code) & "','" & Trim(rs1!sec) & "','" _
                                              & Trim(rs1!Day) & "','" _
                                              & Trim(rs1!Time) & "','" _
                                              & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                              & Trim(rs1!Type) & "'," _
                                              & Val(Trim(rs1!spots)) & "," _
                                              & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                              & Val(Trim(rs1!tra_amount)) & "," _
                                              & Val(Trim(rs1!tra_amount)) & ")"
                    Else
                      Sqlqry2 = " Insert into dumbo_tratv values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                              & Trim(rs1!Month) & "','" _
                                              & findfirstfixup(rs1!Product) & "','" _
                                              & findfirstfixup(rs1!client) & "','" _
                                              & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                              & Trim(rs1!sub_Media) & "','" _
                                              & findfirstfixup(Trim(rs1!bo_ref)) & "','" & Trim(rs1!code) & "','" & Trim(rs1!sec) & "','" _
                                              & Trim(rs1!Day) & "','" _
                                              & Trim(rs1!Time) & "','" _
                                              & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!mat_code)) & "','" _
                                              & Trim(rs1!Type) & "'," _
                                              & Val(Trim(rs1!spots)) & "," _
                                              & Val(Trim(rs1!Rate)) & ",'" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0')"
                    End If
                                              
                     ws.BeginTrans
                     db.Execute (Sqlqry2)
                     ws.CommitTrans
                   rs1.MoveNext
                  Loop
                End If
            ElseIf rs!Media = "Magazine" And rs!sub_Media = "ALAM ASSAYARRAT" Or rs!sub_Media = "ZEINA" Then
              Sqlqry1 = "Select * from bo_tramag where serial_no ='" & Trim(rs!serial_no) & "'"
              Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
              If rs1.RecordCount <> 0 Then
                 rs1.MoveFirst
                 Do Until rs1.EOF
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 If rs1!Type = "Paid" Then
                   Sqlqry2 = " Insert into dumbo_tramagext values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                             & Val(Trim(rs1!tra_amount)) & "," _
                                             & Val(Trim(rs1!tra_amount)) & ",' " _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 Else
                      Sqlqry2 = " Insert into dumbo_tramagext values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                              & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0','" _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 End If
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
                  rs1.MoveNext
                 Loop
               End If
             
           Else
            Sqlqry1 = "Select * from bo_tramag where serial_no ='" & Trim(rs!serial_no) & "'"
            Set rs1 = db.OpenRecordset(Sqlqry1, dbOpenDynaset)
              If rs1.RecordCount <> 0 Then
                 rs1.MoveFirst
                 Do Until rs1.EOF
                 Set ws = DBEngine.Workspaces(0)
                 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
                 If rs1!Type = "Paid" Then
                   Sqlqry2 = " Insert into dumbo_tramag values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "'," & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "'," _
                                             & Val(Trim(rs1!tra_amount)) & "," _
                                             & Val(Trim(rs1!tra_amount)) & ",' " _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 Else
                      Sqlqry2 = " Insert into dumbo_tramag values('" & rs1!serial_no & "','" & rs1!Year & "','" _
                                             & Trim(rs1!Month) & "', " & Val(rs1!monthind) & ",'" _
                                             & findfirstfixup(rs1!Product) & "','" _
                                             & findfirstfixup(rs1!client) & "','" _
                                             & findfirstfixup(rs1!Agency) & "','" & Trim(rs1!Media) & "','" _
                                             & Trim(rs1!sub_Media) & "','" _
                                             & findfirstfixup(Trim(rs1!bo_ref)) & "','" _
                                             & Trim(rs1!issue_no) & "','" & Trim(rs1!tDate) & "','" _
                                             & Trim(rs1!Page) & "','" _
                                             & findfirstfixup(Trim(rs1!Description)) & "','" _
                                             & findfirstfixup(Trim(rs1!Comments)) & "','" _
                                             & findfirstfixup(Trim(rs1!mat_code)) & "','" & Trim(rs1!Space) & "','" _
                                             & Trim(rs1!Type) & "','" & Trim(rs1!tcurrency) & "','" & Trim(rs1!tconvertion) & "','0','0','" _
                                             & Val(Trim(rs1!agcom)) & "','" _
                                             & Val(Trim(rs1!adper)) & "'," _
                                             & Val(Trim(rs1!addisc)) & "," _
                                             & Val(Trim(rs1!surcharge)) & ")"
                 End If
                    ws.BeginTrans
                    db.Execute (Sqlqry2)
                    ws.CommitTrans
                  rs1.MoveNext
                 Loop
               End If
           End If
                  
       rs.MoveNext
       Loop
      End If
           Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_traCin"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
        CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
        CrystalReport1.ReportFileName = App.Path & "\bocininv.rpt"
        CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.DetailCopies = 1
        CrystalReport1.Action = 1
     End If
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_tramag"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
         CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
         CrystalReport1.ReportFileName = App.Path & "\bomaginv.rpt"
         CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
         CrystalReport1.WindowState = crptMaximized
         CrystalReport1.Action = 1
     End If
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_tramagext"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
     
     CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
      CrystalReport1.ReportFileName = App.Path & "\bomaginvext.rpt"
     'CrystalReport1.SelectionFormula = "{invrep.media}='" & "Magazine" & "'"
     CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Action = 1
    End If
    
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_tratv"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
          CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
          CrystalReport1.ReportFileName = App.Path & "\botelinv1.rpt"
          CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
          CrystalReport1.WindowState = crptMaximized
          CrystalReport1.Action = 1
     End If
        
     
     Set ws = DBEngine.Workspaces(0)
     Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
     Sqlqry = "Select * from dumBo_traol"
     Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
     If rs.RecordCount <> 0 Then
          CrystalReport1.DataFiles(0) = App.Path & "\misov.mdb"
          CrystalReport1.ReportFileName = App.Path & "\boolinv.rpt"
          CrystalReport1.Formulas(0) = "xxx1=' ONLY'"
          CrystalReport1.WindowState = crptMaximized
          CrystalReport1.Action = 1
     End If

     CboAyear.Text = Year(Now())
     CboAmonth.ListIndex = Month(Now) - 1
     CboAgency.ListIndex = -1
End Sub
Private Function ValidateData() As Boolean
  ValidateData = False
    If cboYear.ListCount = False Then
        MsgBox "Invalid Year. Select the Year", vbInformation, "Invalid Entry"
        cboYear.SetFocus
        SendKeys "{Home}+{End}"
        Exit Function
    Else
         ValidateData = True
    End If
End Function
Private Function ValidateDataserial() As Boolean
  ValidateDataserial = False
    If CboSerialTo.Text = "" Then
        MsgBox "Invalid Serial # To.  Select the Number", vbInformation, "Invalid Entry"
        CboSerialTo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Function
    Else
         ValidateDataserial = True
    End If
End Function

Private Function ValidateDataParticular() As Boolean
  ValidateDataParticular = False
    If CboParticularNo.Text = "" Then
        MsgBox "Invalid Number,  Select the Number", vbInformation, "Invalid Entry"
        CboParticularNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Function
    Else
         ValidateDataParticular = True
    End If
End Function

Private Function ValidateDatamedia() As Boolean
  ValidateDatamedia = False
    If CboMYear.ListCount = False Then
        MsgBox "Invalid Year. Select the Year", vbInformation, "Invalid Entry"
        CboMYear.SetFocus
        SendKeys "{Home}+{End}"
        Exit Function
    ElseIf CboMMonth.Text = "" Then
       MsgBox "Invalid Month. Select the Month", vbInformation, "Invalid Entry"
       CboMMonth.SetFocus
       SendKeys "{Home} + {End}"
       Exit Function
    ElseIf CboMedia.Text = "" Then
         MsgBox "Invalid Media. Select the Media", vbInformation, "Invalid Entry"
         CboMedia.SetFocus
         SendKeys "{Home} + {End}"
         Exit Function
    Else
         ValidateDatamedia = True
    End If
End Function
Private Function validateDataAgency() As Boolean
validateDataAgency = False
    If CboAyear.ListCount = False Then
        MsgBox "Invalid Year. Select the Year", vbInformation, "Invalid Entry"
        CboAyear.SetFocus
        SendKeys "{Home}+{End}"
        Exit Function
    ElseIf CboAmonth.Text = "" Then
       MsgBox "Invalid Month. Select the Month", vbInformation, "Invalid Entry"
       CboAmonth.SetFocus
       SendKeys "{Home} + {End}"
       Exit Function
    ElseIf CboAgency.Text = "" Then
         MsgBox "Invalid Agency. Select the Agency", vbInformation, "Invalid Entry"
         CboAgency.SetFocus
         SendKeys "{Home} + {End}"
         Exit Function
    Else
         validateDataAgency = True
    End If
End Function

Private Function validateDataProduct() As Boolean
validateDataProduct = False
    If CboPYear.ListCount = False Then
        MsgBox "Invalid Year. Select the Year", vbInformation, "Invalid Entry"
        CboPYear.SetFocus
        SendKeys "{Home}+{End}"
        Exit Function
    ElseIf CboPMonth.Text = "" Then
       MsgBox "Invalid Month. Select the Month", vbInformation, "Invalid Entry"
       CboPMonth.SetFocus
       SendKeys "{Home} + {End}"
       Exit Function
    ElseIf CboProduct.Text = "" Then
         MsgBox "Invalid Product. Select the Product", vbInformation, "Invalid Entry"
         CboProduct.SetFocus
         SendKeys "{Home} + {End}"
         Exit Function
    Else
         validateDataProduct = True
    End If
End Function
Private Sub Form_Load()
Dim i As Integer
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
cboMonth.Clear
cboMonth.AddItem "January"
cboMonth.AddItem "February"
cboMonth.AddItem "March"
cboMonth.AddItem "April"
cboMonth.AddItem "May"
cboMonth.AddItem "June"
cboMonth.AddItem "July"
cboMonth.AddItem "August"
cboMonth.AddItem "September"
cboMonth.AddItem "October"
cboMonth.AddItem "November"
cboMonth.AddItem "December"
cboMonth.ListIndex = Month(Now) - 1

CboMMonth.Clear
CboMMonth.AddItem "January"
CboMMonth.AddItem "February"
CboMMonth.AddItem "March"
CboMMonth.AddItem "April"
CboMMonth.AddItem "May"
CboMMonth.AddItem "June"
CboMMonth.AddItem "July"
CboMMonth.AddItem "August"
CboMMonth.AddItem "September"
CboMMonth.AddItem "October"
CboMMonth.AddItem "November"
CboMMonth.AddItem "December"
CboMMonth.ListIndex = Month(Now) - 1


CboAmonth.Clear
CboAmonth.AddItem "January"
CboAmonth.AddItem "February"
CboAmonth.AddItem "March"
CboAmonth.AddItem "April"
CboAmonth.AddItem "May"
CboAmonth.AddItem "June"
CboAmonth.AddItem "July"
CboAmonth.AddItem "August"
CboAmonth.AddItem "September"
CboAmonth.AddItem "October"
CboAmonth.AddItem "November"
CboAmonth.AddItem "December"
CboAmonth.ListIndex = Month(Now) - 1

CboPMonth.Clear
CboPMonth.AddItem "January"
CboPMonth.AddItem "February"
CboPMonth.AddItem "March"
CboPMonth.AddItem "April"
CboPMonth.AddItem "May"
CboPMonth.AddItem "June"
CboPMonth.AddItem "July"
CboPMonth.AddItem "August"
CboPMonth.AddItem "September"
CboPMonth.AddItem "October"
CboPMonth.AddItem "November"
CboPMonth.AddItem "December"
CboPMonth.ListIndex = Month(Now) - 1

cboYear.Clear
CboMYear.Clear
CboAyear.Clear
CboPYear.Clear

For i = 2000 To 2200
    cboYear.AddItem i
    CboMYear.AddItem i
    CboAyear.AddItem i
    CboPYear.AddItem i
Next i

OptMonth.Value = True
Optmedia.Value = False
OptSerialNo.Value = False
OptParticularno.Value = False

cboYear.Text = Year(Now)
CboMYear.Text = Year(Now)
CboAyear.Text = Year(Now)
CboPYear.Text = Year(Now)


populateserialnos
populatePartno
populateagency
populateproduct
populateMagazine
'PopulateMagagine


 Set ws = DBEngine.Workspaces(0)
 Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
 Sqlqry = "Delete * from invrep"
    ws.BeginTrans
    db.Execute (Sqlqry)
    ws.CommitTrans

End Sub
Private Sub populatePartno()
     Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select serial_no from BO_MAS order by serial_no"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount <> 0 Then
       
     rs.MoveFirst
      Do Until rs.EOF
        CboParticularNo.AddItem Val(rs!serial_no)
        
       rs.MoveNext
      Loop
        
       
    End If

End Sub
Private Sub populateserialnos()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select serial_no from BO_MAS order by serial_no"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
       MsgBox "No Transactions are recorded"
       Exit Sub
    Else
     rs.MoveFirst
      Do Until rs.EOF
        CboSerialFrom.AddItem Val(rs!serial_no)
        CboSerialTo.AddItem Val(rs!serial_no)
       rs.MoveNext
      Loop
        
       
    End If

End Sub
Private Sub populateagency()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select agentname from agndtls order by agentname"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
       MsgBox "No Transactions are recorded"
       Exit Sub
    Else
     rs.MoveFirst
      Do Until rs.EOF
        CboAgency.AddItem Trim(rs!agentname)
       rs.MoveNext
      Loop
    End If

End Sub

Private Sub populateMagazine()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select sub_Media from Media where Media_type='Magazine' order by Sub_Media"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
       MsgBox "No Transactions are recorded"
       Exit Sub
    Else
     rs.MoveFirst
       CboMedia.Clear
       CboMedia.AddItem "Cinema"
       CboMedia.AddItem "Magazine"
       CboMedia.AddItem "Online"
       CboMedia.AddItem "Television"
      
      Do Until rs.EOF
        CboMedia.AddItem "Magazine" & " " & Trim(rs!sub_Media)
       rs.MoveNext
      Loop
    End If

End Sub

Private Sub populateproduct()
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\misov.mdb")
    Sqlqry = "Select product_name from products order by Product_name"
    Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
    If rs.RecordCount = 0 Then
       MsgBox "No Transactions are recorded"
       Exit Sub
    Else
     rs.MoveFirst
      Do Until rs.EOF
        CboProduct.AddItem Trim(rs!product_name)
       rs.MoveNext
      Loop
        
   End If
End Sub

Private Sub PopulateUnFreezed()
 Dim Last As Date
 Dim i As Integer
 Dim j As Integer
 Sqlqry = "Select TDate from cpmt_mas Where Status='Y' order by tdate"
 Set rs = db.OpenRecordset(Sqlqry, dbOpenDynaset)
  If rs.RecordCount <> 0 Then
    rs.MoveLast
    Last = rs!tDate
    i = Month(Last)
    SelMonth = i
    cboMonth.Clear
    Select Case i
    Case 12
        cboMonth.AddItem "January"
    Case 1
        cboMonth.AddItem "February"
    Case 2
        cboMonth.AddItem "March"
    Case 3
        cboMonth.AddItem "April"
    Case 4
        cboMonth.AddItem "May"
    Case 5
        cboMonth.AddItem "June"
    Case 6
        cboMonth.AddItem "July"
    Case 7
        cboMonth.AddItem "August"
    Case 8
        cboMonth.AddItem "September"
    Case 9
        cboMonth.AddItem "October"
    Case 10
        cboMonth.AddItem "November"
    Case 11
        cboMonth.AddItem "December"
    End Select
    j = Year(Last)
    cboYear.Clear
    If i = 12 Then j = j + 1
    cboYear.AddItem j
    cboMonth.ListIndex = 0
    cboYear.ListIndex = 0
  Else
    cboMonth.ListIndex = 0
  End If
End Sub

Private Sub OptAgency_Click()
fraMonth.Visible = False
FraMedia.Visible = False
FraProduct.Visible = False
FraAgency.Visible = True
Fraserialno.Visible = False
FraParticularNo.Visible = False
End Sub

Private Sub Optmedia_Click()
fraMonth.Visible = False
FraMedia.Visible = True
FraProduct.Visible = False
FraAgency.Visible = False
Fraserialno.Visible = False
FraParticularNo.Visible = False
End Sub

Private Sub OptMonth_Click()
fraMonth.Visible = True
FraMedia.Visible = False
FraProduct.Visible = False
FraAgency.Visible = False
FraParticularNo.Visible = False
Fraserialno.Visible = False
End Sub

Private Sub OptParticularno_Click()
fraMonth.Visible = False
FraMedia.Visible = False
FraProduct.Visible = False
FraAgency.Visible = False
Fraserialno.Visible = False
FraParticularNo.Visible = True
End Sub

Private Sub OptProduct_Click()
fraMonth.Visible = False
FraMedia.Visible = False
FraProduct.Visible = True
FraAgency.Visible = False
Fraserialno.Visible = False
FraParticularNo.Visible = False
End Sub

Private Sub OptSerialNo_Click()
fraMonth.Visible = False
FraMedia.Visible = False
FraProduct.Visible = False
FraAgency.Visible = False
Fraserialno.Visible = True
FraParticularNo.Visible = False
End Sub
