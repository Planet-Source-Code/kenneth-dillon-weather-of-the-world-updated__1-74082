VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmWeatherMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Weather Of The World"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   15090
   Icon            =   "frmWeatherMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   15090
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbPhCode 
      Height          =   315
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   124
      Top             =   10680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdNext 
      DisabledPicture =   "frmWeatherMain.frx":0CCA
      Enabled         =   0   'False
      Height          =   375
      Left            =   12740
      MousePointer    =   99  'Custom
      Picture         =   "frmWeatherMain.frx":1D4C
      Style           =   1  'Graphical
      TabIndex        =   123
      ToolTipText     =   "Next"
      Top             =   9860
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      DisabledPicture =   "frmWeatherMain.frx":244E
      Enabled         =   0   'False
      Height          =   375
      Left            =   12230
      MousePointer    =   99  'Custom
      Picture         =   "frmWeatherMain.frx":34D0
      Style           =   1  'Graphical
      TabIndex        =   122
      ToolTipText     =   "Previous"
      Top             =   9860
      Width           =   375
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Find A City"
      Height          =   375
      Left            =   13240
      MousePointer    =   99  'Custom
      TabIndex        =   121
      ToolTipText     =   "Search For City"
      Top             =   9860
      Width           =   1575
   End
   Begin VB.CommandButton cmbZipCode 
      Caption         =   "Get Weather"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10480
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   9860
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   13680
      MousePointer    =   99  'Custom
      TabIndex        =   120
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdCel 
      Caption         =   "&C"
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
      Left            =   13080
      MousePointer    =   99  'Custom
      TabIndex        =   119
      ToolTipText     =   "Change To Celcus"
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   475
   End
   Begin VB.CommandButton cmdFar 
      Caption         =   "&F"
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
      Left            =   12500
      MousePointer    =   99  'Custom
      TabIndex        =   118
      ToolTipText     =   "Change To Far"
      Top             =   4320
      Width           =   475
   End
   Begin VB.ComboBox cmbAnthem 
      Height          =   315
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   117
      Top             =   10440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picHidden 
      AutoSize        =   -1  'True
      Height          =   1095
      Left            =   14160
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   116
      Top             =   10920
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ImageList imgMapFlag 
      Left            =   1320
      Top             =   9840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   112
      ImageHeight     =   168
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5961
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2490
      TabIndex        =   114
      Top             =   9940
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Max             =   1
   End
   Begin VB.TextBox txtCountryStat 
      Height          =   3855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   113
      Text            =   "frmWeatherMain.frx":13D3D
      Top             =   13440
      Visible         =   0   'False
      Width           =   12375
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   6240
      Top             =   10560
   End
   Begin VB.ComboBox cmbCode 
      Height          =   315
      Left            =   12120
      Style           =   2  'Dropdown List
      TabIndex        =   107
      Top             =   10440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cmboZip 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9190
      Sorted          =   -1  'True
      TabIndex        =   106
      Top             =   9860
      Width           =   1175
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   9840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":13D43
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":14A1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":15AA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":15DC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1615B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":164F5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   720
      Top             =   9840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   221
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":17E0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1989F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":19BB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1C596
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1E821
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1E9AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1EB3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1F1F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1F384
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1F494
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1FB1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1FCA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1FE2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":1FFB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2013C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":202D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":205E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":20775
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":20C67
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":210FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":21604
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":21790
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":22057
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2250C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2347D
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":238A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2434C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":25575
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":25F33
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":26878
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":26BD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":26D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2743F
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":28035
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":283D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":28861
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2918F
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":29B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2A66C
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2AC23
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2ADB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2B499
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2BB9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2C33F
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2C7F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2CCE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2D2CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2D66C
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2DEA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2E481
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2EAD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2EF8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2F9F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":2FFFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":306C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":30BB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":30FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":31576
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":31FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":32860
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":33191
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":33D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":343EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":34C6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3547E
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":35E48
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3616C
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":369E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":37090
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":37D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3813D
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":38A2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":39322
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":397D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":39C84
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3A025
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3A439
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3A75C
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3AC5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3B005
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3B4EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3BE9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3CAFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3D09A
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3D54F
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3DDBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3E728
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3EC65
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3F006
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3F527
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":3FE01
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":40112
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":41062
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":417CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":41C81
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":41E0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":422C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":42449
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":425D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":42BC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":43A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":443EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":45045
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":454B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4616E
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4666A
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":46967
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4709D
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":47885
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":47DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":47FFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4875C
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":48AFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":48E9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":49A1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":49E59
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4A829
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4B094
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4B5D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4BA8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4C159
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4CC51
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4D2A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4D6AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4E1FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4E772
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4F4AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":4F7BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":501CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":52122
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":52651
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":52EE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":53927
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":53E3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":546E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":54A85
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":55527
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":55BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":56073
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":564A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":56B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":57075
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":57931
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":57FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":584D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":58B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5955B
            Key             =   ""
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":59D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5A908
            Key             =   ""
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5B1AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5B4A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5BF94
            Key             =   ""
         EndProperty
         BeginProperty ListImage153 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5C40E
            Key             =   ""
         EndProperty
         BeginProperty ListImage154 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5C857
            Key             =   ""
         EndProperty
         BeginProperty ListImage155 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5CBF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage156 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5D475
            Key             =   ""
         EndProperty
         BeginProperty ListImage157 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5DC72
            Key             =   ""
         EndProperty
         BeginProperty ListImage158 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5EC94
            Key             =   ""
         EndProperty
         BeginProperty ListImage159 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5F32F
            Key             =   ""
         EndProperty
         BeginProperty ListImage160 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":5FFBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage161 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":605CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage162 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":618D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage163 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6209B
            Key             =   ""
         EndProperty
         BeginProperty ListImage164 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6243C
            Key             =   ""
         EndProperty
         BeginProperty ListImage165 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":62A36
            Key             =   ""
         EndProperty
         BeginProperty ListImage166 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":63529
            Key             =   ""
         EndProperty
         BeginProperty ListImage167 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":63CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage168 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6451D
            Key             =   ""
         EndProperty
         BeginProperty ListImage169 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":64DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage170 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":65285
            Key             =   ""
         EndProperty
         BeginProperty ListImage171 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":65D2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage172 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":66F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage173 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":67832
            Key             =   ""
         EndProperty
         BeginProperty ListImage174 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":67F1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage175 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6866B
            Key             =   ""
         EndProperty
         BeginProperty ListImage176 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":68B63
            Key             =   ""
         EndProperty
         BeginProperty ListImage177 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":690BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage178 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":69C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage179 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6A0AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage180 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6A567
            Key             =   ""
         EndProperty
         BeginProperty ListImage181 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6AA3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage182 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6B0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage183 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6B80F
            Key             =   ""
         EndProperty
         BeginProperty ListImage184 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6BFFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage185 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6C3B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage186 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6C98F
            Key             =   ""
         EndProperty
         BeginProperty ListImage187 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6CD97
            Key             =   ""
         EndProperty
         BeginProperty ListImage188 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6CF2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage189 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6D585
            Key             =   ""
         EndProperty
         BeginProperty ListImage190 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6DB13
            Key             =   ""
         EndProperty
         BeginProperty ListImage191 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6EBA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage192 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6F8D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage193 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":6FF36
            Key             =   ""
         EndProperty
         BeginProperty ListImage194 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":70056
            Key             =   ""
         EndProperty
         BeginProperty ListImage195 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage196 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":707AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage197 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7093C
            Key             =   ""
         EndProperty
         BeginProperty ListImage198 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":70AC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage199 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7114C
            Key             =   ""
         EndProperty
         BeginProperty ListImage200 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":71A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage201 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":726EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage202 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":72CC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage203 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":73173
            Key             =   ""
         EndProperty
         BeginProperty ListImage204 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":737BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage205 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":73945
            Key             =   ""
         EndProperty
         BeginProperty ListImage206 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":73AC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage207 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":73C47
            Key             =   ""
         EndProperty
         BeginProperty ListImage208 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7451A
            Key             =   ""
         EndProperty
         BeginProperty ListImage209 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":74E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage210 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":75544
            Key             =   ""
         EndProperty
         BeginProperty ListImage211 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":76D23
            Key             =   ""
         EndProperty
         BeginProperty ListImage212 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":77798
            Key             =   ""
         EndProperty
         BeginProperty ListImage213 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7929E
            Key             =   ""
         EndProperty
         BeginProperty ListImage214 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7A49F
            Key             =   ""
         EndProperty
         BeginProperty ListImage215 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7B446
            Key             =   ""
         EndProperty
         BeginProperty ListImage216 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7BC41
            Key             =   ""
         EndProperty
         BeginProperty ListImage217 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7BF15
            Key             =   ""
         EndProperty
         BeginProperty ListImage218 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7CB92
            Key             =   ""
         EndProperty
         BeginProperty ListImage219 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7D49F
            Key             =   ""
         EndProperty
         BeginProperty ListImage220 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7E364
            Key             =   ""
         EndProperty
         BeginProperty ListImage221 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":7F6B9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   65535
      Left            =   9120
      Top             =   10300
   End
   Begin VB.Frame fmFlag 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Country Flag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   12480
      TabIndex        =   81
      Top             =   2400
      Width           =   2295
      Begin VB.Image imgFlag 
         Height          =   1200
         Left            =   120
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   345
         Width           =   2055
      End
   End
   Begin VB.Frame fmMap 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Country Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2175
      Left            =   12480
      TabIndex        =   80
      Top             =   75
      Width           =   2295
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   150
         ScaleHeight     =   1695
         ScaleWidth      =   1965
         TabIndex        =   115
         Top             =   360
         Visible         =   0   'False
         Width           =   1960
      End
      Begin VB.Image imgPicture 
         Height          =   1680
         Left            =   180
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1935
      End
      Begin VB.Image imgMap 
         Height          =   1695
         Left            =   300
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   72
      Top             =   9855
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4234
            MinWidth        =   4234
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   9701
            MinWidth        =   9701
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2090
            MinWidth        =   2082
            Text            =   "Enter Zip"
            TextSave        =   "Enter Zip"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frmToday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Today Forecast"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3975
      Left            =   9020
      TabIndex        =   12
      Top             =   120
      Width           =   3235
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   3015
         TabIndex        =   112
         Top             =   2760
         Width           =   3015
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   3015
         TabIndex        =   111
         Top             =   1440
         Width           =   3015
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   3015
         TabIndex        =   110
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblTodayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   21
         Top             =   3240
         Width           =   45
      End
      Begin VB.Label lblTodayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   20
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label lblTodayTime 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   19
         Top             =   2880
         Width           =   45
      End
      Begin VB.Label lblTodayTime 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   18
         Top             =   1560
         Width           =   45
      End
      Begin VB.Label lblTodayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Top             =   840
         Width           =   45
      End
      Begin VB.Label lblTodayTime 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   480
         Width           =   45
      End
      Begin VB.Label lblTodayDeg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   2580
         TabIndex        =   15
         Top             =   3000
         Width           =   60
      End
      Begin VB.Label lblTodayDeg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   2580
         TabIndex        =   14
         Top             =   1800
         Width           =   60
      End
      Begin VB.Label lblTodayDeg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   2580
         TabIndex        =   13
         Top             =   600
         Width           =   60
      End
      Begin VB.Image imgToday 
         Height          =   855
         Index           =   2
         Left            =   1455
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   855
      End
      Begin VB.Image imgToday 
         Height          =   855
         Index           =   1
         Left            =   1460
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   855
      End
      Begin VB.Image imgToday 
         Height          =   855
         Index           =   0
         Left            =   1460
         Stretch         =   -1  'True
         Top             =   320
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         Index           =   1
         X1              =   0
         X2              =   4080
         Y1              =   2600
         Y2              =   2600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         Index           =   0
         X1              =   0
         X2              =   4080
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "10 Day Forecast"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4960
      Left            =   4320
      TabIndex        =   11
      Top             =   4800
      Width           =   10455
      Begin VB.CommandButton cmdDay 
         Caption         =   "WED"
         Height          =   270
         Index           =   9
         Left            =   9480
         MousePointer    =   99  'Custom
         TabIndex        =   95
         Top             =   320
         Width           =   600
      End
      Begin VB.CommandButton cmdDay 
         Caption         =   "WED"
         Height          =   270
         Index           =   8
         Left            =   8450
         MousePointer    =   99  'Custom
         TabIndex        =   94
         Top             =   320
         Width           =   600
      End
      Begin VB.CommandButton cmdDay 
         Caption         =   "WED"
         Height          =   270
         Index           =   7
         Left            =   7400
         MousePointer    =   99  'Custom
         TabIndex        =   93
         Top             =   320
         Width           =   600
      End
      Begin VB.CommandButton cmdDay 
         Caption         =   "WED"
         Height          =   270
         Index           =   6
         Left            =   6380
         MousePointer    =   99  'Custom
         TabIndex        =   92
         Top             =   320
         Width           =   600
      End
      Begin VB.CommandButton cmdDay 
         Caption         =   "WED"
         Height          =   270
         Index           =   5
         Left            =   5380
         MousePointer    =   99  'Custom
         TabIndex        =   91
         Top             =   320
         Width           =   600
      End
      Begin VB.CommandButton cmdDay 
         Caption         =   "WED"
         Height          =   270
         Index           =   4
         Left            =   4370
         MousePointer    =   99  'Custom
         TabIndex        =   96
         Top             =   320
         Width           =   600
      End
      Begin VB.CommandButton cmdDay 
         Caption         =   "WED"
         Height          =   270
         Index           =   3
         Left            =   3340
         MousePointer    =   99  'Custom
         TabIndex        =   90
         Top             =   320
         Width           =   600
      End
      Begin VB.CommandButton cmdDay 
         Caption         =   "WED"
         Height          =   270
         Index           =   2
         Left            =   2300
         MousePointer    =   99  'Custom
         TabIndex        =   89
         Top             =   320
         Width           =   600
      End
      Begin VB.CommandButton cmdDay 
         Caption         =   "WED"
         Height          =   270
         Index           =   1
         Left            =   1270
         MousePointer    =   99  'Custom
         TabIndex        =   88
         Top             =   320
         Width           =   600
      End
      Begin VB.CommandButton cmdDay 
         Caption         =   "WED"
         Height          =   270
         Index           =   0
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   87
         Top             =   320
         Width           =   600
      End
      Begin VB.PictureBox picDetail 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   8400
         ScaleHeight     =   345
         ScaleWidth      =   1785
         TabIndex        =   86
         Top             =   3600
         Width           =   1815
      End
      Begin VB.PictureBox picDetail 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   6380
         ScaleHeight     =   345
         ScaleWidth      =   1785
         TabIndex        =   85
         Top             =   3600
         Width           =   1815
      End
      Begin VB.PictureBox picDetail 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   4340
         ScaleHeight     =   345
         ScaleWidth      =   1785
         TabIndex        =   84
         Top             =   3600
         Width           =   1815
      End
      Begin VB.PictureBox picDetail 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   2280
         ScaleHeight     =   345
         ScaleWidth      =   1785
         TabIndex        =   83
         Top             =   3600
         Width           =   1815
      End
      Begin VB.PictureBox picDetail 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   240
         ScaleHeight     =   345
         ScaleMode       =   0  'User
         ScaleWidth      =   1785
         TabIndex        =   82
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   1335
         Left            =   240
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2280
         TabIndex        =   105
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblSpeed 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8205
         TabIndex        =   104
         Top             =   4485
         Width           =   45
      End
      Begin VB.Label lblDirection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8205
         TabIndex        =   103
         Top             =   4215
         Width           =   45
      End
      Begin VB.Label lblWaxing 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5895
         TabIndex        =   102
         Top             =   4485
         Width           =   45
      End
      Begin VB.Label lblMoonPhase 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5895
         TabIndex        =   101
         Top             =   4215
         Width           =   45
      End
      Begin VB.Label lblMoonSet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3360
         TabIndex        =   100
         Top             =   4530
         Width           =   45
      End
      Begin VB.Label lblMoonRise 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3360
         TabIndex        =   99
         Top             =   4200
         Width           =   45
      End
      Begin VB.Label lbSunSet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   98
         Top             =   4530
         Width           =   45
      End
      Begin VB.Label lblSunRise 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   97
         Top             =   4200
         Width           =   45
      End
      Begin VB.Image imgWind 
         Height          =   480
         Left            =   7600
         Picture         =   "frmWeatherMain.frx":7FAC9
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image imgMoon 
         Height          =   480
         Left            =   5320
         Picture         =   "frmWeatherMain.frx":8014A
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image imgSunRise 
         Appearance      =   0  'Flat
         Height          =   750
         Index           =   0
         Left            =   240
         Picture         =   "frmWeatherMain.frx":80899
         Top             =   4095
         Width           =   3000
      End
      Begin VB.Label lblDetail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   8120
         TabIndex        =   77
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label lblDetail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   6100
         TabIndex        =   76
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label lblDetail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   4100
         TabIndex        =   75
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label lblDetail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   74
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label lblDetail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   10
         TabIndex        =   73
         Top             =   2960
         Width           =   45
      End
      Begin VB.Image imgDetail 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   4
         Left            =   8400
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Image imgDetail 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   6380
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Image imgDetail 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   4340
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Image imgDetail 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Image imgDetail 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label lblTenDayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   9480
         TabIndex        =   71
         Top             =   2055
         Width           =   45
      End
      Begin VB.Label lblTenDayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   8445
         TabIndex        =   70
         Top             =   2055
         Width           =   45
      End
      Begin VB.Label lblTenDayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   7395
         TabIndex        =   69
         Top             =   2055
         Width           =   45
      End
      Begin VB.Label lblTenDayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   6375
         TabIndex        =   68
         Top             =   2055
         Width           =   45
      End
      Begin VB.Label lblTenDayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   5385
         TabIndex        =   67
         Top             =   2055
         Width           =   45
      End
      Begin VB.Label lblTenDayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   4365
         TabIndex        =   66
         Top             =   2055
         Width           =   45
      End
      Begin VB.Label lblTenDayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3345
         TabIndex        =   65
         Top             =   2055
         Width           =   45
      End
      Begin VB.Label lblTenDayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2295
         TabIndex        =   64
         Top             =   2055
         Width           =   45
      End
      Begin VB.Label lblTenDayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1275
         TabIndex        =   63
         Top             =   2055
         Width           =   45
      End
      Begin VB.Label lblTenDayCon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   62
         Top             =   2055
         Width           =   45
      End
      Begin VB.Label lblTenDayL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   9480
         TabIndex        =   61
         Top             =   2840
         Width           =   735
      End
      Begin VB.Label lblTenDayL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   8450
         TabIndex        =   60
         Top             =   2840
         Width           =   735
      End
      Begin VB.Label lblTenDayL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   7400
         TabIndex        =   59
         Top             =   2840
         Width           =   735
      End
      Begin VB.Label lblTenDayL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   6380
         TabIndex        =   58
         Top             =   2840
         Width           =   735
      End
      Begin VB.Label lblTenDayL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   5380
         TabIndex        =   57
         Top             =   2840
         Width           =   735
      End
      Begin VB.Label lblTenDayL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   4370
         TabIndex        =   56
         Top             =   2840
         Width           =   735
      End
      Begin VB.Label lblTenDayL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3340
         TabIndex        =   55
         Top             =   2840
         Width           =   735
      End
      Begin VB.Label lblTenDayL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2300
         TabIndex        =   54
         Top             =   2840
         Width           =   735
      End
      Begin VB.Label lblTenDayL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1270
         TabIndex        =   53
         Top             =   2840
         Width           =   735
      End
      Begin VB.Label lblTenDayH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   9480
         TabIndex        =   52
         Top             =   2540
         Width           =   615
      End
      Begin VB.Label lblTenDayH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   8450
         TabIndex        =   51
         Top             =   2540
         Width           =   615
      End
      Begin VB.Label lblTenDayH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   7400
         TabIndex        =   50
         Top             =   2540
         Width           =   615
      End
      Begin VB.Label lblTenDayH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   6380
         TabIndex        =   49
         Top             =   2540
         Width           =   615
      End
      Begin VB.Label lblTenDayH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   5380
         TabIndex        =   48
         Top             =   2540
         Width           =   615
      End
      Begin VB.Label lblTenDayH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   4370
         TabIndex        =   47
         Top             =   2540
         Width           =   615
      End
      Begin VB.Label lblTenDayH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3340
         TabIndex        =   46
         Top             =   2540
         Width           =   615
      End
      Begin VB.Label lblTenDayH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2300
         TabIndex        =   45
         Top             =   2540
         Width           =   615
      End
      Begin VB.Label lblTenDayH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1270
         TabIndex        =   44
         Top             =   2540
         Width           =   615
      End
      Begin VB.Label lblTenDayL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   43
         Top             =   2840
         Width           =   735
      End
      Begin VB.Label lblTenDayH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   42
         Top             =   2540
         Width           =   615
      End
      Begin VB.Label lblTenDayD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   9480
         TabIndex        =   41
         Top             =   880
         Width           =   615
      End
      Begin VB.Label lblTenDayD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   8450
         TabIndex        =   40
         Top             =   880
         Width           =   615
      End
      Begin VB.Label lblTenDayD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   7400
         TabIndex        =   39
         Top             =   880
         Width           =   615
      End
      Begin VB.Label lblTenDayD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   6380
         TabIndex        =   38
         Top             =   880
         Width           =   615
      End
      Begin VB.Label lblTenDayD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   5400
         TabIndex        =   37
         Top             =   880
         Width           =   615
      End
      Begin VB.Label lblTenDayD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   4370
         TabIndex        =   36
         Top             =   880
         Width           =   615
      End
      Begin VB.Label lblTenDayD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3340
         TabIndex        =   35
         Top             =   880
         Width           =   615
      End
      Begin VB.Label lblTenDayD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2300
         TabIndex        =   34
         Top             =   880
         Width           =   615
      End
      Begin VB.Label lblTenDayD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1270
         TabIndex        =   33
         Top             =   880
         Width           =   615
      End
      Begin VB.Label lblTenDayM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   9480
         TabIndex        =   32
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblTenDayM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   8450
         TabIndex        =   31
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblTenDayM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   7400
         TabIndex        =   30
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblTenDayM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   6360
         TabIndex        =   29
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblTenDayM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   5400
         TabIndex        =   28
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblTenDayM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   4370
         TabIndex        =   27
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblTenDayM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3340
         TabIndex        =   26
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblTenDayM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2300
         TabIndex        =   25
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblTenDayM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1270
         TabIndex        =   24
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblTenDayD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   880
         Width           =   615
      End
      Begin VB.Label lblTenDayM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.Image imgTenDay 
         Height          =   735
         Index           =   9
         Left            =   9480
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgTenDay 
         Height          =   735
         Index           =   8
         Left            =   8450
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgTenDay 
         Height          =   735
         Index           =   7
         Left            =   7400
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgTenDay 
         Height          =   735
         Index           =   6
         Left            =   6380
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgTenDay 
         Height          =   735
         Index           =   5
         Left            =   5380
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgTenDay 
         Height          =   735
         Index           =   4
         Left            =   4370
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgTenDay 
         Height          =   735
         Index           =   3
         Left            =   3340
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgTenDay 
         Height          =   735
         Index           =   2
         Left            =   2300
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgTenDay 
         Height          =   735
         Index           =   1
         Left            =   1270
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgTenDay 
         Height          =   735
         Index           =   0
         Left            =   240
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgSunRise 
         Height          =   750
         Index           =   1
         Left            =   2760
         Picture         =   "frmWeatherMain.frx":810B5
         Top             =   4080
         Width           =   3000
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   13680
      Top             =   10200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5175
      Left            =   360
      TabIndex        =   4
      Top             =   11020
      Visible         =   0   'False
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   9128
      _Version        =   393217
      TextRTF         =   $"frmWeatherMain.frx":818E7
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Current Conditions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3015
      Left            =   4320
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
      Begin MSComctlLib.ListView lstCurCondition 
         Height          =   1095
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1400
         Visible         =   0   'False
         Width           =   4135
         _ExtentX        =   7303
         _ExtentY        =   1931
         View            =   3
         Arrange         =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "1"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "2"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "3"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "4"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Image imgFire 
         Height          =   700
         Left            =   4000
         Stretch         =   -1  'True
         Top             =   320
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblNoReport 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblNoReport"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   109
         Top             =   960
         Visible         =   0   'False
         Width           =   4200
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNoWeather 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   120
         TabIndex        =   108
         Top             =   360
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label lblTimeCondition 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   600
         Left            =   120
         TabIndex        =   78
         Top             =   2360
         Width           =   4200
      End
      Begin VB.Label lblFeel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2760
         TabIndex        =   7
         Top             =   900
         Width           =   45
      End
      Begin VB.Label lblCondition 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblMainTmp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   135
      End
      Begin VB.Image imgMain 
         Height          =   780
         Left            =   240
         Picture         =   "frmWeatherMain.frx":81972
         Top             =   360
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Countries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   9645
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin MSComctlLib.TreeView TView 
         Height          =   9255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   16325
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   471
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList2"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1200
      Top             =   9840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   59
      ImageHeight     =   156
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":829EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":83D60
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":84F00
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":86213
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":8763D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":88C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":8A2A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":8B908
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":8CE5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":8E175
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":8F2AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":9058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":91A9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":93112
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWeatherMain.frx":94768
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgLrgMap 
      Height          =   1095
      Left            =   2640
      Top             =   9860
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblDayDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Height          =   615
      Left            =   4320
      TabIndex        =   79
      Top             =   4180
      Width           =   8155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weather Report For"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   4320
      TabIndex        =   9
      Top             =   120
      Width           =   3180
   End
   Begin VB.Label lblCity 
      BackStyle       =   0  'Transparent
      Caption         =   "Toronto, Canada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   4320
      TabIndex        =   8
      Top             =   650
      Width           =   4665
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuBook 
      Caption         =   "Bookmarks"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add To Bookmark"
      End
      Begin VB.Menu mnuRemoveBookMark 
         Caption         =   "Remove From Book Marks"
         Begin VB.Menu mnuRemove 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuRemove 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRemove 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRemove 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRemove 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuspace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFavorite 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuFavorite 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFavorite 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFavorite 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFavorite 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Local Weather"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find A City"
      End
      Begin VB.Menu mnuTemp 
         Caption         =   "Temperture"
         Begin VB.Menu mnuCel 
            Caption         =   "Celsius"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuFar 
            Caption         =   "Farenheit"
         End
      End
   End
   Begin VB.Menu mnuStorm 
      Caption         =   "Storms"
      Begin VB.Menu mnuSevereAlert 
         Caption         =   "Severe Weather"
         Begin VB.Menu mnuSevere 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuSevere 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuSevere 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuSevere 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuSevere 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuSevere 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuSevere 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuSevere 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuSevere 
            Caption         =   ""
            Index           =   8
         End
      End
      Begin VB.Menu mnuAlertState 
         Caption         =   "Weather Alerts By States"
         Enabled         =   0   'False
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   10
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   12
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   13
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   14
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   15
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   16
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   17
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   18
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   19
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   20
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   21
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   22
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   23
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   24
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   25
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   26
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   27
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   28
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   29
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   30
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   31
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   32
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   33
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   34
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   35
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   36
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   37
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   38
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   39
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   40
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   41
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   42
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   43
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   44
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   45
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   46
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   47
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   48
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   49
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   50
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   51
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   52
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   53
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   54
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   55
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   56
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   57
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   58
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   59
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   60
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStateAlert 
            Caption         =   ""
            Index           =   61
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuHurricane 
         Caption         =   "Hurricane"
         Begin VB.Menu mnuHur 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuHur 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuHur 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuHur 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuHur 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuHur 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuHur 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuHur 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuHur 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu mnuHur 
            Caption         =   ""
            Index           =   9
         End
      End
      Begin VB.Menu menuActiveStorm 
         Caption         =   ""
         Visible         =   0   'False
         Begin VB.Menu mnuCurTract 
            Caption         =   "Current Track"
            Index           =   0
         End
         Begin VB.Menu mnuCurTract 
            Caption         =   "Visible Satellite"
            Index           =   1
         End
         Begin VB.Menu mnuCurTract 
            Caption         =   "Infrared Satellite"
            Index           =   2
         End
      End
      Begin VB.Menu mnuStorm2 
         Caption         =   ""
         Visible         =   0   'False
         Begin VB.Menu mnuInfrared 
            Caption         =   "Current Track"
            Index           =   0
         End
         Begin VB.Menu mnuInfrared 
            Caption         =   "Visible Satellite"
            Index           =   1
         End
         Begin VB.Menu mnuInfrared 
            Caption         =   "Infrared Satellite"
            Index           =   2
         End
      End
      Begin VB.Menu mnuStorm3 
         Caption         =   ""
         Visible         =   0   'False
         Begin VB.Menu mnuActiveHurricane 
            Caption         =   "Current Track"
            Index           =   0
         End
         Begin VB.Menu mnuActiveHurricane 
            Caption         =   "Visible Satellite"
            Index           =   1
         End
         Begin VB.Menu mnuActiveHurricane 
            Caption         =   "Infrared Satellite"
            Index           =   2
         End
      End
      Begin VB.Menu mnuStormList 
         Caption         =   ""
         Visible         =   0   'False
         Begin VB.Menu mnuStormAdvisory 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuStormAdvisory 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStormAdvisory 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuSeasn 
         Caption         =   "Season Summaries"
         Begin VB.Menu mnuHS 
            Caption         =   " "
            Index           =   0
         End
         Begin VB.Menu mnuHS 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuHS 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuHS 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuHS 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuHS 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuHS 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuHS 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuHS 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu mnuHS 
            Caption         =   ""
            Index           =   9
         End
         Begin VB.Menu mnuHS 
            Caption         =   ""
            Index           =   10
         End
         Begin VB.Menu mnuHS 
            Caption         =   ""
            Index           =   11
         End
      End
   End
   Begin VB.Menu mnuRadar 
      Caption         =   "Radar"
      Begin VB.Menu mnuRCurrent 
         Caption         =   "Current"
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   9
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   10
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   11
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   12
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   13
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   14
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   15
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   16
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   17
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   18
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   19
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   20
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   21
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   22
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   23
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   24
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   25
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   26
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   27
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   28
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   29
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   30
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   31
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   32
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   33
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   34
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   35
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   36
         End
         Begin VB.Menu mnuRadCur 
            Caption         =   ""
            Index           =   37
         End
      End
      Begin VB.Menu mnuCurLoop 
         Caption         =   "Current Loops"
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   9
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   10
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   11
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   12
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   13
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   14
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   15
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   16
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   17
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   18
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   19
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   20
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   21
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   22
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   23
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   24
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   25
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   26
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   27
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   28
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   29
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   30
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   31
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   32
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   33
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   34
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   35
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   36
         End
         Begin VB.Menu mnuCurLp 
            Caption         =   ""
            Index           =   37
         End
      End
      Begin VB.Menu mnuRadForeCase 
         Caption         =   "Forcast"
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   9
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   10
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   11
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   12
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   13
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   14
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   15
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   16
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   17
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   18
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   19
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   20
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   21
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   22
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   23
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   24
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   25
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   26
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   27
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   28
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   29
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   30
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   31
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   32
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   33
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   34
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   35
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   36
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   37
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   38
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   39
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   40
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   41
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   42
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   43
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   44
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   45
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   46
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   47
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   48
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   49
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   50
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   51
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   52
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   53
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   54
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   55
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   56
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   57
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   58
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   59
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   60
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   61
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   62
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   63
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   64
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   65
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   66
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   67
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   68
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   69
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   70
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   71
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   72
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   73
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   74
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   75
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   76
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   77
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   78
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   79
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   80
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   81
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   82
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   83
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   84
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   85
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   86
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   87
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   88
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   89
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   90
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   91
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   92
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   93
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   94
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   95
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   96
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   97
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   98
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   99
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   100
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   101
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   102
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   103
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   104
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   105
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   106
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   107
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   108
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   109
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   110
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   111
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   112
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   113
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   114
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   115
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   116
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   117
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   118
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   119
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   120
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   121
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   122
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   123
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   124
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   125
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   126
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   127
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   128
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   129
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   130
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   131
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   132
         End
         Begin VB.Menu mnuRadFor 
            Caption         =   ""
            Index           =   133
         End
      End
      Begin VB.Menu Mnu1K 
         Caption         =   "Regional (1Km)"
      End
      Begin VB.Menu mnuReg1K 
         Caption         =   "Regional (1Km) Loop"
      End
      Begin VB.Menu mnuMetro 
         Caption         =   "Metro"
      End
      Begin VB.Menu mnuMetroloop 
         Caption         =   "Metro Loop"
      End
      Begin VB.Menu mnuSummery 
         Caption         =   "Summary"
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   10
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   12
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   13
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   14
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   15
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   16
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   17
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   18
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   19
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   20
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   21
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   22
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   23
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   24
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   25
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   26
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   27
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   28
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   29
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   30
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   31
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   32
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   33
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   34
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   35
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   36
            Visible         =   0   'False
         End
         Begin VB.Menu muuRadarSummary 
            Caption         =   ""
            Index           =   37
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuJRloop 
         Caption         =   "Java Radar Loop"
      End
   End
   Begin VB.Menu mnuSatellite 
      Caption         =   "Satellite"
      Begin VB.Menu mnuSatGlobal 
         Caption         =   ""
         Begin VB.Menu mnuGbSat 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuGbSat 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuGbSat 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuGbSat 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuGbSat 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuGbSat 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuGbSat 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuGbSat 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuGbSat 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu mnuGbSat 
            Caption         =   ""
            Index           =   9
         End
         Begin VB.Menu mnuGbSat 
            Caption         =   ""
            Index           =   10
         End
         Begin VB.Menu mnuGbSat 
            Caption         =   ""
            Index           =   11
         End
      End
      Begin VB.Menu mnuSat 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuSat 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mnuSat 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu mnuSat 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu mnuSat 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu mnuVisibleSatellite 
         Caption         =   "Visible Satellite"
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   9
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   10
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   11
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   12
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   13
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   14
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   15
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   16
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   17
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   18
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   19
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   20
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   21
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   22
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   23
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   24
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   25
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   26
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   27
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   28
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   29
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   30
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   31
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   32
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   33
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   34
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   35
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   36
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   37
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   38
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   39
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   40
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   41
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   42
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   43
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   44
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   45
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   46
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   47
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   48
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   49
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   50
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   51
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   52
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   53
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   54
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   55
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   56
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   57
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   58
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   59
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   60
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   61
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   62
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   63
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   64
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   65
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   66
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   67
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   68
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   69
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   70
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   71
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   72
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   73
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   74
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   75
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   76
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   77
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   78
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   79
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   80
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   81
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   82
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   83
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   84
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   85
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   86
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   87
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   88
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   89
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   90
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   91
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   92
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   93
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   94
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   95
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   96
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   97
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   98
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   99
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   100
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   101
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   102
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   103
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   104
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   105
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   106
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   107
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   108
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   109
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   110
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   111
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   112
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   113
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   114
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   115
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   116
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   117
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   118
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   119
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   120
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   121
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   122
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   123
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   124
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   125
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   126
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   127
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   128
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   129
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   130
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   131
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   132
         End
         Begin VB.Menu mnuVisSat 
            Caption         =   ""
            Index           =   133
         End
      End
      Begin VB.Menu mnuCurrentSatellite 
         Caption         =   "Current Satellite"
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   9
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   10
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   11
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   12
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   13
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   14
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   15
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   16
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   17
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   18
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   19
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   20
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   21
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   22
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   23
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   24
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   25
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   26
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   27
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   28
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   29
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   30
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   31
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   32
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   33
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   34
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   35
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   36
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   37
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   38
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   39
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   40
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   41
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   42
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   43
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   44
         End
         Begin VB.Menu mnuCurSat 
            Caption         =   ""
            Index           =   45
         End
      End
      Begin VB.Menu mnuWaterVaper 
         Caption         =   ""
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   9
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   10
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   11
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   12
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   13
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   14
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   15
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   16
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   17
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   18
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   19
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   20
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   21
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   22
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   23
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   24
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   25
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   26
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   27
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   28
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   29
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   30
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   31
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   32
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   33
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   34
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   35
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   36
         End
         Begin VB.Menu mnuWV 
            Caption         =   ""
            Index           =   37
         End
      End
   End
   Begin VB.Menu mnuWeather 
      Caption         =   "Weather Alert"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuShowMap 
      Caption         =   "Show GPS Location"
      Begin VB.Menu mnuGPS 
         Caption         =   ""
      End
   End
   Begin VB.Menu mnuCountryStat 
      Caption         =   "Country Information"
      Begin VB.Menu mnuCountryFact 
         Caption         =   ""
      End
      Begin VB.Menu mnuAnthem 
         Caption         =   ""
      End
      Begin VB.Menu mnuStatistics 
         Caption         =   ""
      End
      Begin VB.Menu mnuPhoneCode 
         Caption         =   ""
      End
      Begin VB.Menu mnuTimeDate 
         Caption         =   ""
      End
      Begin VB.Menu mnuCountryHol 
         Caption         =   ""
         Begin VB.Menu mnu2012 
            Caption         =   "Holidays In 2012"
         End
         Begin VB.Menu mnu2011 
            Caption         =   "Holidays In 2011"
         End
         Begin VB.Menu mnu2010 
            Caption         =   "Holidays In  2010"
         End
         Begin VB.Menu mnu2009 
            Caption         =   "Holidays In 2009"
         End
         Begin VB.Menu mnu2008 
            Caption         =   "Holidays In 2008"
         End
      End
      Begin VB.Menu mnuDistance 
         Caption         =   "Hotel - Flight And Distance Between Countries"
      End
   End
   Begin VB.Menu mnuWorld 
      Caption         =   "World Statistics"
      Begin VB.Menu mnuWorldCap 
         Caption         =   "Capitals Of The World"
      End
      Begin VB.Menu mnuHolDate 
         Caption         =   "Countries National Holidays By Date"
      End
      Begin VB.Menu mnuNatHoliday 
         Caption         =   "National Holidays Around the World"
      End
      Begin VB.Menu mnuRace 
         Caption         =   "Ethnicity and Race by Countries"
      End
      Begin VB.Menu mnuSevenWonders 
         Caption         =   "Seven Wonders of the Modern World"
      End
      Begin VB.Menu mnuTallest 
         Caption         =   "Tallest Buildings in the World"
      End
      Begin VB.Menu mnuEcoStat 
         Caption         =   "Economic Statistics by Country"
         Begin VB.Menu mnuEcoStat2009 
            Caption         =   "Year 2009"
         End
         Begin VB.Menu mnuEcoStat2008 
            Caption         =   "Year 2008"
         End
         Begin VB.Menu mnuEcoStat2005 
            Caption         =   "Year 2005"
         End
      End
      Begin VB.Menu mnuComNation 
         Caption         =   "Members of the Commonwealth of Nations"
      End
   End
   Begin VB.Menu mnuPopStatistics 
      Caption         =   "Population Statistics"
      Begin VB.Menu mnuPopulation 
         Caption         =   "Area and Population"
      End
      Begin VB.Menu mnuPopDensity 
         Caption         =   "Population Density"
      End
      Begin VB.Menu mnu50PopCountries 
         Caption         =   "World's 50 Most Populous Countries:"
         Begin VB.Menu mnuYear2010 
            Caption         =   "Year 2010"
         End
         Begin VB.Menu mnuYear2009 
            Caption         =   "Year 2009"
         End
         Begin VB.Menu mnuYear2008 
            Caption         =   "Year 2008"
         End
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmWeatherMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim phload As Boolean
Dim noFlags As Boolean
Dim sCountryUrl As String
Dim xAni As Integer
Dim anm1 As Integer
Dim FarEnable As Boolean
Dim sSelCountryName As String
Dim sStringToFind As String
Dim sSelCityName As String
Dim SatName As String
Dim bNextState As Boolean
Dim bPreState As Boolean
Dim sStateBoxCode As String
Dim sCountryCode As String
Dim nLen As Integer
Dim bNodeFound As Boolean
Dim sCityName As String
Dim sCityCode As String
Dim iArraycnt As Integer
Dim oldBtIndex As Integer
Dim iredoIndex As Integer
Dim IndexArray() As Long
Dim LinkArray() As String
Dim AnthemArray() As String
Dim PhoneArray() As String
Dim LrgMapAddress As String
Dim fso As FileSystemObject
Dim oldLetterNode As String
Dim oldNameIndex As Long
Dim oldCountryNode As String
Dim IsCelsius As Boolean
Dim curNameIndex As Long
Public itnetCon As Boolean
Dim zipButton As Boolean
Dim bStormBulletins As Boolean
Private Const sPassword = "PasswordIsAGoodThingToHaveArround"
Private Declare Function InternetAttemptConnect Lib "wininet" (ByVal dwReserved As Long) As Long
Private Const FLAG_ICC_FORCE_CONNECTION = &H1

Private Sub cmboZip_Change()
  If Len(cmboZip.Text) = 5 Then
    cmbZipCode.Enabled = True
  Else
    cmbZipCode.Enabled = False
  End If
End Sub

Private Sub cmboZip_Click()
  If Len(cmboZip.Text) = 5 Then
    cmbZipCode.Enabled = True
  Else
    cmbZipCode.Enabled = False
  End If
End Sub

Private Sub cmbZipCode_Click()
  Dim USAcityCode As String
  
  MousePointer = 11
  DisableMenu False
  zipButton = True
  USAcityCode = GetCityCode(cmboZip.Text)
  If Len(USAcityCode) = 0 Then
    MsgBox "Zip Code Does Not Exist", vbInformation, "The Weather Of The World"
    cmboZip.Text = ""
    Exit Sub
  End If
  reMoveIcons
  GetWeather USAcityCode
  GetCountryFagMap "United States"
  GetlargeMap
  sSelCountryName = "United States"
  zipButton = False
  If bNodeFound Then
    TView_DblClick
  Else
    mnuCountryHol.Caption = "United States National Holidays"
    StatusBar1.Panels(2).Text = "Listing For: " & lblCity.Caption & Space(4) & "Region: " & "United States"
  End If
  DisableMenu True
  MousePointer = 0
End Sub

Private Sub cmdCel_Click()
  Dim oFoundNode As Node
  
  cmdFar.Enabled = True
  cmdCel.Enabled = False
  mnuFar.Checked = False
  mnuCel.Checked = True
  IsCelsius = True
  
  sCityCode = QueryValue(HKEY_CURRENT_USER, CityCodeValue, "City_Tag_Name")
  sCityCode = StripTerminator(sCityCode)
  Set oFoundNode = TreeFindNode(TView, sCityName, True, 1)
  
  If Not bNodeFound Then
    zipButton = True
    GetWeather sCityCode
    zipButton = False
  Else
    GetWeather TView.Nodes(curNameIndex).Tag
  End If
  TView.SetFocus
  Set oFoundNode = Nothing
End Sub

Private Sub cmdDay_Click(Index As Integer)
  Dim sDayIndex As Integer
  TView.Enabled = False
  sDayIndex = Index
  'Get Day detail
  GetDayDetails sDayIndex, TView.Nodes(curNameIndex).Tag
  cmdDay(Index).FontBold = True
  cmdDay(oldBtIndex).FontBold = False
  TView.Enabled = True
  oldBtIndex = Index
  TView.SetFocus
  'Reset Timer
  iMinCount = 0
End Sub

Private Sub cmdDay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  cmdDay(Index).ToolTipText = cmdDay(Index).Caption & " Detail Conditions"
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdNext_Click()
  If iredoIndex <= UBound(IndexArray, 1) Then
    If iredoIndex = UBound(IndexArray, 1) Then
      curNameIndex = IndexArray(iredoIndex)
      iredoIndex = UBound(IndexArray, 1)
      cmdNext.Enabled = False
    Else
      curNameIndex = IndexArray(iredoIndex + 1)
      iredoIndex = iredoIndex + 1
      cmdPrevious.Enabled = True
    End If
    If iredoIndex = UBound(IndexArray, 1) Then
      cmdNext.Enabled = False
    End If
    bNextState = cmdNext.Enabled
    bPreState = cmdPrevious.Enabled
    'Close previous  Node
    If TView.Nodes(oldNameIndex).Parent <> TView.Nodes(curNameIndex).Parent Then
      TView.Nodes(oldNameIndex).Parent.Expanded = False
    End If
    oldNameIndex = curNameIndex
    TView_DblClick
    TView.Nodes(curNameIndex).EnsureVisible
    TView.Nodes(curNameIndex).Selected = True
  End If
End Sub

Private Sub cmdPrevious_Click()
  If iredoIndex <= UBound(IndexArray, 1) Then
    If iredoIndex = 0 Then
      curNameIndex = IndexArray(iredoIndex)
      cmdPrevious.Enabled = False
    Else
      curNameIndex = IndexArray(iredoIndex - 1)
      iredoIndex = iredoIndex - 1
      cmdNext.Enabled = True
    End If
    If iredoIndex = 0 Then
      cmdPrevious.Enabled = False
    End If
    bPreState = cmdPrevious.Enabled
    bNextState = cmdNext.Enabled
    'Close previous  Node
    If TView.Nodes(oldNameIndex).Parent <> TView.Nodes(curNameIndex).Parent Then
      TView.Nodes(oldNameIndex).Parent.Expanded = False
    End If
    oldNameIndex = curNameIndex
    TView_DblClick
    TView.Nodes(curNameIndex).EnsureVisible
    TView.Nodes(curNameIndex).Selected = True
  End If
End Sub

Private Sub cmdSearch_Click()
  Dim sFindString As String
  Dim lItemIndex As Long, oFoundNode As Node
  sFindString = InputBox("Enter City To Find", "Weather Of The World", "Toronto", frmWeatherMain.Left + 6000, frmWeatherMain.Top + 4000)
  If Len(sFindString) <> 0 Then
    Do
      lItemIndex = lItemIndex + 1
      Set oFoundNode = TreeViewFindNode(TView, sFindString, True, lItemIndex)
      If oFoundNode Is Nothing Then
        'Didn't find any more items
        MsgBox "No More " & sFindString & " In Countries!", vbInformation, "Weather Of The World City Search"
        Exit Do
      End If
      oFoundNode.EnsureVisible
      If MsgBox("Found " & oFoundNode.Text & " In " & oFoundNode.Parent & vbNewLine & "Find next matching item? ", vbQuestion + vbYesNo, "Weather Of The World City Search") = vbNo Then
        oFoundNode.Selected = True
        Exit Do
      End If
    Loop
  End If
  Set oFoundNode = Nothing
  TView.SetFocus
End Sub

Private Sub Form_Click()
  stopAnimate
End Sub

Private Sub Form_DblClick()
  anm1 = 0
  Image1.Left = anm1
  Timer2.Enabled = True
  Image1.Visible = True
  xAni = 1
End Sub

Private Sub Form_Initialize()
  InitCommonControlsXP
End Sub

Private Sub Form_Load()
  Dim X As Integer
  Dim hMenu   As Long
  Dim lStyle As Long
  Dim oFoundNode As Node
  
  If Check_Connection Then
    'disable MAXIMIZE button
    lStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
    lStyle = lStyle And Not WS_MAXIMIZEBOX
    Call SetWindowLong(Me.hwnd, GWL_STYLE, lStyle)
    frmWeatherMain.Height = 11100
    frmSplash.Show
    DoEvents
    If year(Now) <> 2012 Then
      mnu2012.Visible = False
    End If
    frmWeatherMain.Icon = ImageList1.ListImages(1).Picture
    cmdCel.Caption = Chr(176) & "C"
    cmdFar.Caption = Chr(176) & "F"
    StatusBar1.Panels(1).Text = Format(Date, "Long Date")
    'StatusBar1.Panels(4).Text = "Design By Kenneth Dillon For Kentronics Inc."
    Set cmdExit.MouseIcon = ImageList1.ListImages(3).Picture
    Set cmdCel.MouseIcon = ImageList1.ListImages(3).Picture
    Set cmdFar.MouseIcon = ImageList1.ListImages(3).Picture
    Set cmdSearch.MouseIcon = ImageList1.ListImages(3).Picture
    Set cmdNext.MouseIcon = ImageList1.ListImages(3).Picture
    Set imgFire.Picture = ImageList1.ListImages(6).Picture
    Set cmdPrevious.MouseIcon = ImageList1.ListImages(3).Picture
    Set cmbZipCode.MouseIcon = ImageList1.ListImages(3).Picture
    For X = 0 To 9
      Set cmdDay(X).MouseIcon = ImageList1.ListImages(3).Picture
    Next
    LoadCountryFlags
    LoadTreeView
    LoadComboBox
    LoadCountryHol
    
    DoEvents
    curNameIndex = GetSetting("The Weather Program", "City Information", "Code_Name", "7319")
    sCityName = GetSetting("The Weather Program", "City Information", "City_Name", "Toronto")
    sCityCode = GetSetting("The Weather Program", "City Information", "City_Tag_Name", "CAXX0504")
    IsCelsius = GetSetting("The Weather Program", "Conversion", "Celsius", "True")
    If IsCelsius Then
      cmdCel.Enabled = False
      cmdFar.Enabled = True
    Else
      cmdCel.Enabled = True
      cmdFar.Enabled = False
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(App.Path & "\Icons") = False Then
      fso.CreateFolder App.Path & "\Icons"
    End If
    GetCityTag
    Set oFoundNode = TreeFindNode(TView, sCityName, True, 1)
    If bNodeFound Then
      GetWeather TView.Nodes(curNameIndex).Tag
      GetHurricane
      If InStr(1, TView.Nodes(curNameIndex).Parent, "Saint", vbTextCompare) Then
         GetCountryFagMap "St." & Mid(TView.Nodes(curNameIndex).Parent, InStr(1, TView.Nodes(curNameIndex).Parent, " ", vbTextCompare))
      ElseIf Mid(TView.Nodes(curNameIndex).Tag, 1, 2) = "US" Then
         GetCountryFagMap "United States"
         sSelCountryName = "United States"
      Else
         GetCountryFagMap TView.Nodes(curNameIndex).Parent
      End If
      GetlargeMap
      DoEvents
      If TView.Nodes(curNameIndex).Parent.Parent = "United States" Then
        mnuCountryHol.Caption = "United States National Holidays"
      Else
        mnuCountryHol.Caption = TView.Nodes(curNameIndex).Parent & " National Holidays"
      End If
      'Load Array
      ReDim Preserve IndexArray(iArraycnt)
      IndexArray(iArraycnt) = curNameIndex
      iArraycnt = iArraycnt + 1
      iredoIndex = UBound(IndexArray, 1)
      'Select node
      TView.Nodes(curNameIndex).Expanded = True
      TView.Nodes("ROOT").Expanded = True
      TView.Nodes(curNameIndex).Selected = True
      GetRegion TView.Nodes(curNameIndex).Parent
      ReDim Preserve IndexArray(iArraycnt)
      IndexArray(iArraycnt) = curNameIndex
      oldNameIndex = curNameIndex
      DoEvents
    Else
      zipButton = True
      GetWeather sCityCode
      GetHurricane
      GetCountryFagMap "United States"
      GetRegion "United States"
      sSelCityName = sCityName
      sSelCountryName = "United States"
      Nozip = True
      TView.Nodes(225).Selected = True
      mnuCountryHol.Caption = "United States National Holidays"
      StatusBar1.Panels(2).Text = "Listing For: " & lblCity.Caption & Space(4) & "Region: " & "United States"
    End If
    'Regions
    getSatRegions
    GetSatWaterVaper
    GetCurrentSatellite
    getVisSatellite
    GetRadForcast
    UpdateMenuValues 0, False
    Timer1.Enabled = True
    Unload frmSplash
    Set oFoundNode = Nothing
    LoadNatAnthem
    GetRadSummary
    GerWeatherBulletins
    If bStormBulletins Then
      GetBulletins
    End If
  Else
    MsgBox "No Internet Connection Available", vbInformation, "Weather Of The World"
  End If
End Sub

Public Sub LoadTreeView()
  Dim tmpNode        As Node
  Dim TmpString      As String
  Dim oldCountry As String
  Dim oldTmpString As String
  Dim tmpNameString  As String
  Dim tmpLetter As String
  Dim nX As Long
  Dim IndxCnt As Long
  Dim nFileNum As Integer
  Dim sString As String
  Dim myArray() As String
  
  'On Error GoTo TreeView_error
  'Clear the treeview and node
  Set tmpNode = Nothing
  TView.Visible = False
  TView.Nodes.Clear
  TView.Enabled = False
  
  'This is Used to Add The "ROOT" Node
  Set tmpNode = TView.Nodes.Add(, , "ROOT", "Countries", 4, 4)
  
  'Store Some Information In The Node's Tag
  TView.Nodes("ROOT").Tag = "ROOT"
  TView.Nodes("ROOT").Bold = True
  TView.Nodes("ROOT").ForeColor = vbRed 'Blue
  'Add A-Z
  For nX = 0 To 25
    'Store The Category Name To tmpString
    TmpString = Chr(65 + nX)
    'Add the Relation Nodes
    Set tmpNode = TView.Nodes.Add("ROOT", tvwChild, TmpString, TmpString, 1, 1)
    'Store Some Information In The Node's Tag
    TView.Nodes(TmpString).Tag = TmpString
    TView.Nodes(TmpString).Bold = True
  Next
  nX = 0
  'Add Countries to first letter node
  nFileNum = FreeFile
  Open App.Path & "\Region Cities All.Dat" For Binary Access Read As #nFileNum
  'On Error Resume Next
  Do While Not EOF(nFileNum)
    'read the length of the string
    Get #nFileNum, , nLen
    'initialize the string with the correct number of spaces
    sString = Space$(nLen)
    Get #nFileNum, , sString
    sString = DecryptText((sString), sPassword, True)
    If Len(Trim$(sString)) > 1 Then
      myArray = Split(sString, ",")
      TmpString = myArray(2)
      tmpLetter = UCase(Mid(myArray(2), 1, 1))
      If myArray(3) = "United States" Then
        TmpString = myArray(3)
        tmpLetter = UCase(Mid(myArray(3), 1, 1))
      End If
      'Check for duplicate
      If TmpString <> oldTmpString Then
          'Add the Relation Nodes
        sfndResult = FindStringinListControl(cmbCode, Trim(TmpString))
        If sfndResult <> -1 Then
          Set tmpNode = TView.Nodes.Add(TView.Nodes(tmpLetter).Tag, tvwChild, TmpString, TmpString, sfndResult + 5, sfndResult + 5)
        Else
          Set tmpNode = TView.Nodes.Add(TView.Nodes(tmpLetter).Tag, tvwChild, TmpString, TmpString, 3, 3)
        End If
        'Store Some Information In The Node's Tag
        TView.Nodes(TmpString).Tag = TmpString
        TView.Nodes(TmpString).Bold = True
        TView.Nodes(TmpString).ForeColor = vbBlue
        oldTmpString = TmpString
      End If
    End If
  Loop
  'Add Zip code entry
  Set tmpNode = TView.Nodes.Add(TView.Nodes("United States").Tag, tvwChild, "Zip Code Entry", "Zip Code Entry", 197, 197)
  'Store Some Information In The Node's Tag
  TView.Nodes(TmpString).Tag = TmpString
  TView.Nodes(TmpString).Bold = True
  TView.Nodes(TmpString).ForeColor = vbBlue
  Close #nFileNum
  'Load Cities to countries
  nX = 0
  nFileNum = FreeFile
  Open App.Path & "\Region Cities All.Dat" For Binary Access Read As #nFileNum
  
  Do While Not EOF(nFileNum)
    'read the length of the string
    Get #nFileNum, , nLen
    'initialize the string with the correct number of spaces
    sString = Space$(nLen)
    Get #nFileNum, , sString
    sString = DecryptText((sString), sPassword, True)
    If Len(Trim$(sString)) > 1 Then
      myArray = Split(sString, ",")
      tmpLetter = UCase(Mid(myArray(2), 1, 1))
      TmpString = myArray(2)
      tmpNameString = myArray(1)
      If myArray(3) = "United States" Then
        oldCountry = myArray(3)
        TmpString = myArray(2)
        tmpNameString = myArray(1)
        If oldTmpString <> TmpString Then
          Set tmpNode = TView.Nodes.Add(TView.Nodes(oldCountry).Tag, tvwChild, TmpString, TmpString, 197, 197)
          TView.Nodes(TmpString).Tag = TmpString
        End If
      End If
      oldTmpString = TmpString
      If myArray(3) = "United States" Then
        Set tmpNode = TView.Nodes.Add(TView.Nodes(TmpString).Tag, tvwChild, tmpNameString & IndxCnt, tmpNameString, 197, 197)
        TView.Nodes(tmpNameString & IndxCnt).Tag = myArray(0)
        IndxCnt = IndxCnt + 1
        If IndxCnt > 32500 Then
          Exit Do
        End If
      Else
        sfndResult = FindStringinListControl(cmbCode, TView.Nodes(TmpString).Tag)
        If sfndResult <> -1 Then
          Set tmpNode = TView.Nodes.Add(TView.Nodes(TmpString).Tag, tvwChild, tmpNameString & IndxCnt, tmpNameString, sfndResult + 5, sfndResult + 5)
        Else
          Set tmpNode = TView.Nodes.Add(TView.Nodes(TmpString).Tag, tvwChild, tmpNameString & IndxCnt, tmpNameString, 2, 2)
        End If
        
        TView.Nodes(TmpString).Tag = TmpString
        TView.Nodes(tmpNameString & IndxCnt).Tag = myArray(0)
        IndxCnt = IndxCnt + 1
      End If
    End If
  Loop
  Close #nFileNum
endme:
  TView.Enabled = True
  TView.Nodes("ROOT").Expanded = True
  TView.Nodes("ROOT").Selected = True
  TView.Visible = True
  Set tmpNode = Nothing
  Exit Sub
TreeView_error:
  If Err.Number <> 0 Then
    If Err.Number = 35602 Or Err.Number = 35601 Then
      Set tmpNode = TView.Nodes.Add(TView.Nodes(TmpString).Tag, tvwChild, tmpNameString & nX, tmpNameString, 2, 2)
      TView.Nodes(tmpNameString & nX).Tag = myArray(0)
      nX = nX + 1
      Err.Clear
      'Resume Next
    Else
      MsgBox Err.Number
      MsgBox "Error Loading Treeview : " & Err.Description & vbCrLf & _
              "Error # : " & Str$(Err.Number) & ".", vbCritical + vbOKOnly
        Err.Clear
        Resume Next
    End If
  End If
End Sub

Private Sub cmdFar_Click()
  Dim oFoundNode As Node
  
  cmdFar.Enabled = False
  cmdCel.Enabled = True
  mnuFar.Checked = True
  mnuCel.Checked = False
  FarEnable = False
  IsCelsius = False
  
  sCityCode = QueryValue(HKEY_CURRENT_USER, CityCodeValue, "City_Tag_Name")
  sCityCode = StripTerminator(sCityCode)
  Set oFoundNode = TreeFindNode(TView, sCityName, True, 1)
  
  sCityCode = StripTerminator(sCityCode)
  If Not bNodeFound Then
    zipButton = True
    GetWeather sCityCode
    zipButton = False
  Else
    GetWeather TView.Nodes(curNameIndex).Tag
  End If
  TView.SetFocus
  Set oFoundNode = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting "The Weather Program", "Conversion", "Celsius", IsCelsius
  SaveSetting "The Weather Program", "City Information", "Code_Name", curNameIndex
  SaveSetting "The Weather Program", "City Information", "City_Name", sCityName
  reMoveIcons
  Set frmWeatherMain = Nothing
End Sub

Private Sub Frame3_Click()
  stopAnimate
End Sub

Private Sub imgFlag_Click()
   On Error Resume Next
   stopAnimate
   sFrmName = "Flag Of " & scntName
   picTureName = sFlagPicture
   Load frmCountry
End Sub

Private Sub imgFlag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgFlag.BorderStyle = 1
End Sub

Private Sub imgFlag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgFlag.MouseIcon = ImageList1.ListImages(3).Picture
  imgFlag.ToolTipText = "Click To Enlarge"
End Sub

Private Sub imgFlag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgFlag.BorderStyle = 0
End Sub

Private Sub imgPicture_Click()
  On Error Resume Next
  stopAnimate
  sFrmName = "Map Of " & scntName
  GetlargeMap
  Load frmCountry
End Sub

Private Sub imgPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPicture.BorderStyle = 1
End Sub

Private Sub imgPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPicture.MouseIcon = ImageList1.ListImages(3).Picture
  imgPicture.ToolTipText = "Click To Enlarge"
End Sub

Private Sub imgPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPicture.BorderStyle = 0
End Sub

Private Sub lblDayDetail_Click()
  stopAnimate
End Sub

Private Sub mnu2008_Click()
  sfndResult = FindStringinListControl(cmbCode, Replace(sSelCountryName, "&", "And"))
  If sfndResult <> -1 Then
    GetCountrylHol LinkArray(sfndResult), "2008"
  End If
End Sub

Private Sub mnu2009_Click()
  sfndResult = FindStringinListControl(cmbCode, Replace(sSelCountryName, "&", "And"))
  If sfndResult <> -1 Then
    GetCountrylHol LinkArray(sfndResult), "2009"
  End If
End Sub

Private Sub mnu2010_Click()
  sfndResult = FindStringinListControl(cmbCode, Replace(sSelCountryName, "&", "And"))
  If sfndResult <> -1 Then
    GetCountrylHol LinkArray(sfndResult), "2010"
  End If
End Sub

Private Sub mnu2011_Click()
  sfndResult = FindStringinListControl(cmbCode, Replace(sSelCountryName, "&", "And"))
  If sfndResult <> -1 Then
    GetCountrylHol LinkArray(sfndResult), "2011"
  End If
End Sub

Private Sub mnu2012_Click()
  sfndResult = FindStringinListControl(cmbCode, Replace(sSelCountryName, "&", "And"))
  If sfndResult <> -1 Then
    GetCountrylHol LinkArray(sfndResult), "2011"
  End If
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal
End Sub

Private Sub mnuActiveHurricane_Click(Index As Integer)
  sStatusText = mnuStorm3.Caption & " " & mnuActiveHurricane(Index).Caption
  GetCurrentTrack mnuActiveHurricane(Index).Tag, Index
End Sub

Private Sub mnuAdd_Click()
  Dim X As Integer
  Dim oFoundNode As Node
  
  sCityCode = StripTerminator(QueryValue(HKEY_CURRENT_USER, CityCodeValue, "City_Tag_Name"))
  Set oFoundNode = TreeFindNode(TView, sCityName, True, 1)
  
  For X = 0 To 4
    If Len(mnuFavorite(X).Caption) = 0 Then
      If bNodeFound Then
        mnuFavorite(X).Caption = TView.Nodes(curNameIndex).Text & " - " & TView.Nodes(curNameIndex).Parent
        mnuRemove(X).Caption = TView.Nodes(curNameIndex).Text & " - " & TView.Nodes(curNameIndex).Parent
        SaveSetting "The Weather Program", "BookMark", "City_Name-" & X, TView.Nodes(curNameIndex).Text & " - " & TView.Nodes(curNameIndex).Parent
        SaveSetting "The Weather Program", "BookMark", "City_Tag_Name-" & X, TView.Nodes(curNameIndex).Tag
      Else
        mnuFavorite(X).Caption = Replace(lblCity.Caption, ",", " -")
        mnuRemove(X).Caption = Replace(lblCity.Caption, ",", " -")
        SaveSetting "The Weather Program", "BookMark", "City_Name-" & X, Replace(lblCity.Caption, ",", " -")
        SaveSetting "The Weather Program", "BookMark", "City_Tag_Name-" & X, sCityCode
      End If
      mnuFavorite(X).Tag = X
      mnuRemove(X).Tag = X
      mnuFavorite(X).Enabled = True
      mnuRemove(X).Enabled = True
      mnuFavorite(X).Visible = True
      mnuRemove(X).Visible = True
      mnuRemoveBookMark.Enabled = True
      Exit For
    End If
  Next
  For X = 0 To 4
    If Len(mnuFavorite(X).Caption) = 0 Then
      mnuAdd.Enabled = True
      Exit For
    Else
      mnuAdd.Enabled = False
    End If
  Next
  bNodeFound = False
  Set oFoundNode = Nothing
End Sub

Private Sub mnuAnthem_Click()
  sfndResult = FindStringinListControl(cmbAnthem, sSelCountryName)
  If sfndResult <> -1 Then
    GetCountryAnthem AnthemArray(sfndResult), sSelCountryName
  Else
    MsgBox "Unable To Show " & sSelCountryName & " Anthem", vbInformation, "Weather Of The World"
  End If
End Sub

Private Sub mnuCel_Click()
   If mnuFar.Checked = True Then
      mnuCel.Checked = True
      mnuFar.Checked = False
      cmdCel_Click
   End If
End Sub

Private Sub mnuComNation_Click()
  GetCommNation "http://www.infoplease.com/uk/language/difference-great-britain-england-isles.html"
End Sub

Private Sub mnuCountryFact_Click()
  mnuCountryStat.Enabled = False
  frmWeatherMain.MousePointer = 11
  GetCountryFacts sCountryUrl
  frmWeatherMain.MousePointer = 0
  mnuCountryStat.Enabled = True
End Sub

Private Sub mnuCurLp_Click(Index As Integer)
   PlayRegAnimation = True
   SatName = " Current Radar Loop"
   GetAnimation "/National/Radar/Current.aspx?region=" & mnuCurLp(Index).Tag, mnuCurLp(Index).Caption
   PlayRegAnimation = False
End Sub

Private Sub mnuCurSat_Click(Index As Integer)
   PlayRegAnimation = True
   SatName = " Current Satellite"
   GetSevereWeatherMap "/Global/Satellite/Current.aspx?region=" & mnuCurSat(Index).Tag, mnuCurSat(Index).Caption
   PlayRegAnimation = False
End Sub

Private Sub mnuCurTract_Click(Index As Integer)
  sStatusText = menuActiveStorm.Caption & " " & mnuCurTract(Index).Caption
  GetCurrentTrack mnuCurTract(Index).Tag, Index
End Sub

Private Sub mnuDistance_Click()
  Timer1.Enabled = False
  frmDistance.Show vbModal
  Timer1.Enabled = True
End Sub

Private Sub mnuEcoStat2005_Click()
GetEconomicStats "http://www.infoplease.com/ipa/A0874911.html", "2005"
End Sub

Private Sub mnuEcoStat2008_Click()
  GetEconomicStats "http://www.infoplease.com/world/statistics/economic-statistics-by-country-2008.html", "2008"
End Sub

Private Sub mnuEcoStat2009_Click()
  GetEconomicStats "http://www.infoplease.com/world/statistics/economic-statistics-by-country.html", "2009"
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub

Private Sub mnuFar_Click()
   If mnuCel.Checked = True Then
      mnuFar.Checked = True
      mnuCel.Checked = False
      cmdFar_Click
   End If
End Sub

Private Sub mnuFavorite_Click(Index As Integer)
  Dim oFoundNode As Node
  
  oldNameIndex = curNameIndex
  bNodeFound = False
  ''DoEvents
  sCityName = Mid(mnuFavorite(Index).Caption, 1, InStr(1, mnuFavorite(Index).Caption, " - ", vbTextCompare) - 1)
  sCityCode = StripTerminator(QueryValue(HKEY_CURRENT_USER, FilelistKey, "City_Tag_Name-" & Index))
  Set oFoundNode = TreeFindNode(TView, sCityName, True, 1)
  
  If Not bNodeFound Then
    zipButton = True
    reMoveIcons
    GetWeather sCityCode
    GetCountryFagMap "United States"
    GetlargeMap
    sSelCountryName = "United States"
    mnuCountryHol.Caption = "United States National Holidays"
    StatusBar1.Panels(2).Text = "Listing For: " & lblCity.Caption & Space(4) & "Region: " & "United States"
    If TView.Nodes(224).Expanded = True Then
      TView.Nodes(curNameIndex).Parent.Expanded = False
      TView.Nodes(224).Expanded = False
    End If
    TView.Nodes(225).Selected = True
    zipButton = False
  Else
    If TView.Nodes(oldNameIndex).Parent <> TView.Nodes(curNameIndex).Parent Then
      TView.Nodes(oldNameIndex).Parent.Expanded = False
    End If
    If TView.Nodes(oldNameIndex).Parent.Parent = "United States" Or TView.Nodes(curNameIndex).Parent.Parent = "United States" Then
      TView.Nodes(oldNameIndex).Parent.Parent.Expanded = False
    End If
    TView_DblClick
    TView.Nodes(curNameIndex).EnsureVisible
    TView.Nodes(curNameIndex).Selected = True
  End If
  TView.SetFocus
  bNodeFound = False
  Set oFoundNode = Nothing
End Sub

Private Sub mnuFind_Click()
   cmdSearch_Click
End Sub

Private Sub mnuGbSat_Click(Index As Integer)
   PlayRegAnimation = True
   SatName = " Global Infrared Satellite"
   GetSevereWeatherMap "/Global/Satellite/Infrared.aspx?region=" & mnuGbSat(Index).Tag, mnuGbSat(Index).Caption
   PlayRegAnimation = False
End Sub

Private Sub mnuGPS_Click()
   bGPS = True
   GetLatitude sSelCityName, sSelCountryName
End Sub

Private Sub mnuHolDate_Click()
  Timer1.Enabled = False
  frmDate.Show vbModal
  GetCountriesNatlHol
  Timer1.Enabled = True
End Sub

Private Sub mnuHS_Click(Index As Integer)
   Timer1.Enabled = False
   sFrmName = mnuHS(Index).Caption
   GetHurricaneSumMap mnuHS(Index).Tag
   Timer1.Enabled = True
End Sub

Private Sub mnuHur_Click(Index As Integer)
   Timer1.Enabled = False
   sFrmName = mnuHur(Index).Caption
   If mnuHur(Index).Caption = "Active Track" Then
      frmAlert.lsvStormName.Visible = True
      GetHurricaneTrack mnuHur(Index).Tag
   Else
      GetHurricaneMap mnuHur(Index).Tag
   End If
   Timer1.Enabled = True
End Sub

Private Sub mnuInfrared_Click(Index As Integer)
  sStatusText = mnuStorm2.Caption & " " & mnuInfrared(Index).Caption
  GetCurrentTrack mnuInfrared(Index).Tag, Index
End Sub

Private Sub mnuNatHoliday_Click()
  Timer1.Enabled = False
  GetNatHoliday
  Timer1.Enabled = True
End Sub

Private Sub mnuPhoneCode_Click()
  Timer1.Enabled = False
  If Not phload Then
    LaodPhoneCode
  End If
  MousePointer = 11
  sfndResult = FindStringinListControl(cmbPhCode, sSelCountryName)
  If sfndResult <> -1 Then
    GetCountryPhoneCode PhoneArray(sfndResult)
  Else
    MsgBox "Unable To Show " & sSelCountryName & " Phone Code", vbInformation, "Weather Of The World"
  End If
  phload = True
  MousePointer = 0
  Timer1.Enabled = True
End Sub

Private Sub mnuPopDensity_Click()
  MousePointer = 11
  GetPopDensity
  MousePointer = 0
End Sub

Private Sub mnuPopulation_Click()
  MousePointer = 11
  GetPopulation
  MousePointer = 0
End Sub

Private Sub mnuRace_Click()
  GetRaceofCountry
End Sub

Private Sub mnuRadCur_Click(Index As Integer)
   PlayRegAnimation = True
   SatName = " Current Radar"
   GetSevereWeatherMap "/National/Radar/Current.aspx?region=" & mnuRadCur(Index).Tag, mnuRadCur(Index).Caption
   PlayRegAnimation = False
End Sub

Private Sub mnuRadFor_Click(Index As Integer)
   SatName = " Current Radar"
   GetSevereWeatherMap "/National/Radar/Forecast.aspx?region=" & mnuRadFor(Index).Tag, mnuRadFor(Index).Caption
   PlayRegAnimation = False
End Sub

Private Sub mnuRemove_Click(Index As Integer)
  If MsgBox("Are You Sure You Wish To Delete" & vbCrLf & mnuRemove(Index).Caption & " From Bookmark?", 292, "Weather Of The World Bookmark") = vbYes Then
    UpdateMenuValues Index, True
  End If
End Sub

Private Sub mnuSat_Click(Index As Integer)
   GetSevereWeatherMap mnuSat(Index).Tag, mnuSat(Index).Caption
End Sub

Private Sub mnuSevenWonders_Click()
  Timer1.Enabled = False
  isTallest = False
  frmWorldStat.Show vbModal
  Timer1.Enabled = True
End Sub

Private Sub mnuSevere_Click(Index As Integer)
  If mnuSevere(Index).Caption = "Weather Alerts" Then
    GetWeatherAlerts mnuSevere(Index).Tag
  Else
    GetSevereWeatherMap mnuSevere(Index).Tag, mnuSevere(Index).Caption
  End If
End Sub

Private Sub mnuStateAlert_Click(Index As Integer)
  GetStateAlerts mnuStateAlert(Index).Tag, mnuStateAlert(Index).Caption
End Sub

Private Sub mnuStatistics_Click()
  Timer1.Enabled = False
  If Not phload Then
    LaodPhoneCode
  End If
  MousePointer = 11
  sfndResult = FindStringinListControl(cmbPhCode, sSelCountryName)
  If sfndResult <> -1 Then
    getCountryStatic PhoneArray(sfndResult)
  Else
    MsgBox "Unable To Show " & sSelCountryName & " Phone Code", vbInformation, "Weather Of The World"
  End If
  phload = True
  MousePointer = 0
  Timer1.Enabled = True
End Sub

Private Sub mnuStormAdvisory_Click(Index As Integer)
  GetWeatherAdvisory mnuStormAdvisory(Index).Tag
End Sub

Private Sub mnuTallest_Click()
  Timer1.Enabled = False
  isTallest = True
  frmWorldStat.Show vbModal
  Timer1.Enabled = True
End Sub

Private Sub mnuTimeDate_Click()
  GetCountryTimeDate sSelCityName, sSelCountryName
End Sub

Private Sub mnuVisSat_Click(Index As Integer)
   PlayRegAnimation = True
   SatName = " Visible Satellite"
   GetSevereWeatherMap "/National/Satellite/Visible.aspx?region=" & mnuVisSat(Index).Tag, mnuVisSat(Index).Caption
   PlayRegAnimation = False
End Sub

Private Sub mnuWeather_Click()
   GetWeatherAlert
End Sub

Private Sub mnuWorldCap_Click()
  GetWorldCapital "http://www.infoplease.com/ipa/A0855603.html"
End Sub

Private Sub mnuWV_Click(Index As Integer)
   PlayRegAnimation = True
   SatName = " Water Vapor Satellite"
   GetSevereWeatherMap "/National/Satellite/WaterVapor.aspx?region=" & mnuWV(Index).Tag, mnuWV(Index).Caption
   PlayRegAnimation = False
End Sub

Private Sub mnuYear2008_Click()
  Get50MostPop "http://www.infoplease.com/world/statistics/most-populous-countries-2008.html", "2008"
End Sub

Private Sub mnuYear2009_Click()
  Get50MostPop "http://www.infoplease.com/world/statistics/most-populous-countries-2009.html", "2009"
End Sub

Private Sub mnuYear2010_Click()
  Get50MostPop "http://www.infoplease.com/world/statistics/most-populous-countries.html", "2010"
End Sub

Private Sub muuRadarSummary_Click(Index As Integer)
  slargeMapLink1 = "http://www.intellicast.com/National/Radar/Summary.aspx?region=" & muuRadarSummary(Index).Tag
  GetWebpage slargeMapLink1 '"http://www.intellicast.com/Storm/Hurricane/Active.aspx?storm=1&type=track"
  DisplayRadarMap
End Sub

Private Sub Timer1_Timer()
  On Error Resume Next
  iMinCount = iMinCount + 1
  If iMinCount = 15 And Check_Connection = True Then
    GetCityTag
    Dim oFoundNode As Node
    Set oFoundNode = TreeFindNode(TView, sCityName, True, 1)
    If Not bNodeFound Then
      zipButton = True
      GetWeather sCityCode
      zipButton = False
    Else
      GetWeather TView.Nodes(curNameIndex).Tag
    End If
    cmdDay(oldBtIndex).FontBold = False
    cmdDay(0).FontBold = True
    iMinCount = 0
    anm1 = 0
    Image1.Left = anm1
    Timer2.Enabled = True
    Image1.Visible = True
    xAni = 1
    Set oFoundNode = Nothing
  End If
End Sub

Private Sub Timer2_Timer()
  anm1 = anm1 + 80
  Image1.Left = anm1
  
  If anm1 > Frame3.Width Then
    anm1 = -Image1.Width
    Image1.Left = anm1
  End If
  Image1.Picture = ImageList3.ListImages(xAni).Picture
  xAni = IIf(xAni = ImageList3.ListImages.Count, 1, xAni + 1)
End Sub

Private Sub TView_Click()
  stopAnimate
End Sub

Private Sub TView_DblClick()
  If curNameIndex = 225 Or curNameIndex = 1 Then Exit Sub
  reMoveIcons
  SaveSetting "The Weather Program", "City Information", "Code_Name", curNameIndex
  SaveSetting "The Weather Program", "City Information", "City_Tag_Name", TView.Nodes(curNameIndex).Tag
  SaveSetting "The Weather Program", "City Information", "City_Name", TView.Nodes(curNameIndex).Text
  TView.Enabled = False
  MousePointer = 11
  DisableMenu False
  GetWeather TView.Nodes(curNameIndex).Tag
  sCityName = TView.Nodes(curNameIndex).Text
  If iredoIndex <= 10 And oldNameIndex <> curNameIndex Then
    ReDim Preserve IndexArray(iArraycnt)
    IndexArray(iArraycnt) = curNameIndex
    iArraycnt = iArraycnt + 1
    iredoIndex = UBound(IndexArray, 1)
    cmdPrevious.Enabled = True
    cmdNext.Enabled = False
  End If
  oldNameIndex = curNameIndex
  'DoEvents
  If Nozip Then
    cmdPrevious.Enabled = False
  End If
  Nozip = False
  If InStr(1, TView.Nodes(curNameIndex).Parent, "Saint", vbTextCompare) Then
    GetCountryFagMap "St." & Mid(TView.Nodes(curNameIndex).Parent, InStr(1, TView.Nodes(curNameIndex).Parent, " ", vbTextCompare))
  ElseIf Mid(TView.Nodes(curNameIndex).Tag, 1, 2) = "US" Then
    GetCountryFagMap "United States"
    GetlargeMap
    sSelCountryName = "United States"
    mnuCountryHol.Caption = "United States National Holidays"
    oldLetterNode = TView.Nodes(curNameIndex).Parent
  Else
    GetCountryFagMap TView.Nodes(curNameIndex).Parent
    If Not noFlags Then
      GetlargeMap
    End If
    oldLetterNode = TView.Nodes(curNameIndex).Parent
  End If
  GetRegion TView.Nodes(oldLetterNode).Text
  mnuCountryHol.Caption = TView.Nodes(curNameIndex).Parent & " National Holidays"
  If TView.Nodes(curNameIndex).Parent.Parent = "United States" Then
    mnuCountryHol.Caption = "United States National Holidays"
  Else
    mnuCountryHol.Caption = TView.Nodes(curNameIndex).Parent & " National Holidays"
  End If
  cmdDay(0).FontBold = True
  If oldBtIndex <> 0 Then
    cmdDay(oldBtIndex).FontBold = False
  End If
  oldBtIndex = 0
  'Reset Timer
  iMinCount = 0
  DisableMenu True
  MousePointer = 0
End Sub

Private Sub TView_Expand(ByVal Node As MSComctlLib.Node)
  'LetterNode
  If Node <> "Countries" Or TView.Nodes(Node.Index).Tag <> "ROOT" Then
    If TView.Nodes(Node.Index).Parent = "Countries" Then
      If Node <> oldCountryNode And Len(oldCountryNode) <> 0 Then
        If TView.Nodes(oldCountryNode).Expanded Then
          TView.Nodes(oldCountryNode).Expanded = False
        End If
      End If
      oldCountryNode = Node
    End If
    'Country node
    If Len(TView.Nodes(Node.Index).Parent) < 2 Then
      If Node <> oldLetterNode And Len(oldLetterNode) <> 0 Then
        If TView.Nodes(oldLetterNode).Expanded Then
          TView.Nodes(oldLetterNode).Expanded = False
        End If
      End If
      oldLetterNode = Node
      scntName = Node
    End If
  End If
End Sub

Private Sub TView_NodeClick(ByVal Node As MSComctlLib.Node)
  curNameIndex = Node.Index
End Sub

Private Sub GetWeather(sStateCode As String)
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim sPageName As String
   Dim sStartPos As String
   Dim imageUrl As String
   Dim X As Integer
   Dim sDayDetail As String
   Dim uvPerct As Integer
   Dim iLeftpos As Integer
   Dim oFoundNode As Node
   Dim USAzipTreee As Boolean
   
   On Error GoTo errorHandler
   
   If IsCelsius = False Then
      sPageName = "http://www.intellicast.com/Local/Weather.aspx?unit=F&location=" & sStateCode
   Else
      sPageName = "http://www.intellicast.com/Local/Weather.aspx?unit=C&location=" & sStateCode
   End If
   stopAnimate
   sCountryCode = sStateCode
   cmboZip.Text = ""
   GetWebpage sPageName
   sStartPos = "Primary Header FloatLeft"
   DoEvents
   
   lstCurCondition.ListItems.Clear
   'City And Country
   iIndex3 = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
   iIndexSt = InStr(iIndex3, RichTextBox1.Text, "style=", vbTextCompare)
   iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
   iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "</div>", vbTextCompare)
   lblCity.Caption = Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex2) - (iIndexEnd + 1))
   If Len(Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex2) - (iIndexEnd + 1))) > 33 Then
      lblCity.Top = 530
      lblCity.Caption = Mid(lblCity.Caption, 1, InStr(1, lblCity.Caption, ",")) & vbCrLf & Mid(lblCity.Caption, InStrRev(lblCity.Caption, ",") + 1)
   Else
      lblCity.Top = 650
   End If
   If Mid(sStateCode, 1, 2) = "US" Then
      mnuCountryFact.Caption = "Facts And Figures For United States"
      mnuAnthem.Caption = "United States National Anthem"
      mnuPhoneCode.Caption = "United States Phone Area Codes"
      mnuStatistics.Caption = "United States Statistics"
   Else
      mnuCountryFact.Caption = "Facts And Figures For " & Mid(lblCity.Caption, InStr(1, lblCity.Caption, ",") + 1)
      mnuAnthem.Caption = Mid(lblCity.Caption, InStr(1, lblCity.Caption, ",") + 2) & " National Anthem"
      mnuPhoneCode.Caption = Mid(lblCity.Caption, InStr(1, lblCity.Caption, ",") + 2) & " Phone Area Codes"
      mnuStatistics.Caption = Mid(lblCity.Caption, InStr(1, lblCity.Caption, ",") + 2) & " Statistics"
   End If
   mnuTimeDate.Caption = "Time And Date For " & lblCity.Caption
   mnuGPS.Caption = "GPS Of " & lblCity.Caption
   iIndex = InStr(iIndex2, RichTextBox1.Text, "?location=", vbTextCompare)
   iIndex2 = InStr(iIndex, RichTextBox1.Text, " class=", vbTextCompare)
   bNodeFound = False
   SaveSetting "The Weather Program", "City Information", "City_Tag_Name", sStateCode
   If zipButton Then
      sCityCode = Mid(RichTextBox1.Text, iIndex + 10, iIndex2 - 1 - (iIndex + 10))
      sCityName = Mid(lblCity.Caption, 1, InStr(1, lblCity.Caption, ",") - 1)
      SaveSetting "The Weather Program", "City Information", "City_Tag_Name", sCityCode
      SaveSetting "The Weather Program", "City Information", "City_Name", sCityName
      SaveSetting "The Weather Program", "City Information", "Code_Name", "12345"
      mnuCountryFact.Caption = "Facts And Figures For United States"
      mnuAnthem.Caption = "United States National Anthem"
      Set oFoundNode = TreeFindNode(TView, sCityName, True, 1)
      If bNodeFound Then
         oFoundNode.EnsureVisible
         oFoundNode.Selected = True
         Exit Sub
      Else
        USAzipTreee = True
        sSelCountryName = "United States"
      End If
   End If
   iIndex = InStr(iIndex2, RichTextBox1.Text, "Current Conditions", vbTextCompare)
   'Time of Weather
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "<div style=", vbTextCompare)
   iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
   iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "</div>", vbTextCompare)
   lblTimeCondition.Caption = Mid(RichTextBox1.Text, iIndexEnd + 1, iIndex2 - (iIndexEnd + 1))
   If InStr(1, lblTimeCondition.Caption, "not", vbTextCompare) <> 0 Then
      iIndexSt = InStr(iIndex2, RichTextBox1.Text, "<span class=", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
      iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "</span>", vbTextCompare)
      lblNoWeather.Caption = Replace(Mid(RichTextBox1.Text, iIndexEnd + 1, iIndex2 - (iIndexEnd + 1)), "<br />", "")
      iIndexSt = InStr(iIndex2, RichTextBox1.Text, " />", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</div>", vbTextCompare)
      lblNoReport.Caption = vbCrLf & Replace(Mid(RichTextBox1.Text, iIndexSt + 57, iIndexEnd - (iIndexSt + 57)), "<br />", vbCrLf)
      Picture1.Visible = True
      Picture2.Visible = True
      Picture3.Visible = True
      lblNoWeather.Visible = True
      lblNoReport.Visible = True
      lstCurCondition.Visible = False
      GoTo noReport
   End If
   Picture1.Visible = False
   Picture2.Visible = False
   Picture3.Visible = False
   lblNoReport.Visible = False
   lblNoWeather.Visible = False
   lstCurCondition.Visible = True
   'Current Image
   iIndex3 = InStr(iIndex2 + 1, RichTextBox1.Text, "<img src=", vbTextCompare)
   iIndex = InStr(iIndex3, RichTextBox1.Text, " title=", vbTextCompare)
   imageUrl = Mid(RichTextBox1.Text, iIndex3 + 10, (iIndex - 1) - (iIndex3 + 10))
   SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), imgMain
   'Condition
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "/>", vbTextCompare)
   iIndex2 = InStr(iIndex3, RichTextBox1.Text, "</td>", vbTextCompare)
   lblCondition.Caption = Mid(RichTextBox1.Text, iIndex3 + 2, (iIndex2 - 1) - (iIndex3 + 2))
   'Current temp
   iIndex3 = InStr(iIndex2, RichTextBox1.Text, " title=", vbTextCompare)
   iIndex = InStr(iIndex3, RichTextBox1.Text, "</a>", vbTextCompare)
   lblMainTmp.Caption = Replace(Mid(RichTextBox1.Text, iIndex3 + 21, (iIndex) - (iIndex3 + 21)), "&deg;", Chr(176))
   If IsCelsius And Val(lblMainTmp.Caption) >= 27 Then
      lblMainTmp.ForeColor = vbRed
      lblFeel.ForeColor = vbRed
      imgFire.Visible = True
    ElseIf Not IsCelsius And Val(lblMainTmp.Caption) >= 80 Then
      lblMainTmp.ForeColor = vbRed
      lblFeel.ForeColor = vbRed
      imgFire.Visible = True
    Else
      lblMainTmp.ForeColor = vbBlack
      lblFeel.ForeColor = vbBlack
      imgFire.Visible = False
    End If
   'Feel Like
   iIndex3 = InStr(iIndex, RichTextBox1.Text, " title=", vbTextCompare)
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, ">", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</a>", vbTextCompare)
   lblFeel.Caption = Replace(Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex) - (iIndexEnd + 1)), "&deg;", Chr(176))
   'Wind chill
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<td>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</td>", vbTextCompare)
   lstCurCondition.ListItems.Add , , "Wind Chill:"
   lstCurCondition.ListItems(1).ListSubItems.Add , , Replace(Mid(RichTextBox1.Text, iIndexEnd + 4, (iIndex - (iIndexEnd + 4))), "&deg;", Chr(176))
   'Ceiling
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<td>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</td>", vbTextCompare)
   lstCurCondition.ListItems(1).ListSubItems.Add , , "Ceiling:"
   lstCurCondition.ListItems(1).ListSubItems.Add , , Mid(RichTextBox1.Text, iIndexEnd + 4, (iIndex - (iIndexEnd + 4)))
   'Heat Index
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<td>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</td>", vbTextCompare)
   lstCurCondition.ListItems.Add , , "Heat index:"
   lstCurCondition.ListItems(2).ListSubItems.Add , , Replace(Mid(RichTextBox1.Text, iIndexEnd + 4, (iIndex - (iIndexEnd + 4))), "&deg;", Chr(176))
   'Visibility
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<td>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</td>", vbTextCompare)
   lstCurCondition.ListItems(2).ListSubItems.Add , , "Visibility:"
   lstCurCondition.ListItems(2).ListSubItems.Add , , Mid(RichTextBox1.Text, iIndexEnd + 4, (iIndex - (iIndexEnd + 4)))
   'Dew Point
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<td>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</td>", vbTextCompare)
   lstCurCondition.ListItems.Add , , "Dew Point:"
   lstCurCondition.ListItems(3).ListSubItems.Add , , Replace(Mid(RichTextBox1.Text, iIndexEnd + 4, (iIndex - (iIndexEnd + 4))), "&deg;", Chr(176))
   'Wind
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<td>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</td>", vbTextCompare)
   lstCurCondition.ListItems(3).ListSubItems.Add , , "Wind:"
   lstCurCondition.ListItems(3).ListSubItems.Add , , Mid(RichTextBox1.Text, iIndexEnd + 4, (iIndex - (iIndexEnd + 4)))
   'Humidity
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<td>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</td>", vbTextCompare)
   lstCurCondition.ListItems.Add , , "Humidity:"
   lstCurCondition.ListItems(4).ListSubItems.Add , , Mid(RichTextBox1.Text, iIndexEnd + 4, (iIndex - (iIndexEnd + 4)))
   'Direction
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, " style=", vbTextCompare)
   iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, ">", vbTextCompare)
   iIndex = InStr(iIndexSt, RichTextBox1.Text, "</td>", vbTextCompare)
   lstCurCondition.ListItems(4).ListSubItems.Add , , "Direction:"
   lstCurCondition.ListItems(4).ListSubItems.Add , , Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex - (iIndexSt + 1))), "&deg;", Chr(176))
   'Pressure
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<td>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</td>", vbTextCompare)
   lstCurCondition.ListItems.Add , , "Pressure:"
   lstCurCondition.ListItems(5).ListSubItems.Add , , Mid(RichTextBox1.Text, iIndexEnd + 4, (iIndex - (iIndexEnd + 4)))
   'Gust
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<td>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</td>", vbTextCompare)
   lstCurCondition.ListItems(5).ListSubItems.Add , , "Gusts:"
   lstCurCondition.ListItems(5).ListSubItems.Add , , Mid(RichTextBox1.Text, iIndexEnd + 4, (iIndex - (iIndexEnd + 4)))
   'Today's Forecast
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
   iIndexSt = InStr(iIndex3, RichTextBox1.Text, ">", vbTextCompare)
   iIndex = InStr(iIndex3, RichTextBox1.Text, "</a></div>", vbTextCompare)
   frmToday.Caption = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex - (iIndexSt + 1)))
   For X = 0 To 2
      '1st Time
      iIndex3 = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
      iIndex = InStr(iIndex3, RichTextBox1.Text, "</strong><br />", vbTextCompare)
      lblTodayTime(X).Caption = Mid(RichTextBox1.Text, iIndex3 + 8, (iIndex - (iIndex3 + 8)))
      '1st Image
      iIndex3 = InStr(iIndex, RichTextBox1.Text, "src=", vbTextCompare)
      iIndex = InStr(iIndex3, RichTextBox1.Text, " title=", vbTextCompare)
      imageUrl = Mid(RichTextBox1.Text, iIndex3 + 5, (iIndex - 1) - (iIndex3 + 5))
      SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), imgToday(X)
      '1st Condition
      iIndex3 = InStr(iIndex, RichTextBox1.Text, " title=", vbTextCompare)
      iIndex = InStr(iIndex3, RichTextBox1.Text, " alt", vbTextCompare)
      lblTodayCon(X).Caption = Mid(RichTextBox1.Text, iIndex3 + 8, ((iIndex - 1) - (iIndex3 + 8)))
       '1st degree
      iIndex3 = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
      iIndex = InStr(iIndex3, RichTextBox1.Text, "</strong></td>", vbTextCompare)
      lblTodayDeg(X).Caption = Replace(Mid(RichTextBox1.Text, iIndex3 + 8, ((iIndex) - (iIndex3 + 8))), "&deg;", Chr(176))
   Next
noReport:
   '10 Day Forecast
   iIndex3 = InStr(iIndex, RichTextBox1.Text, "10 Day Forecast", vbTextCompare)
   'Day's of the week
   For X = 0 To 9
      iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "onclick=", vbTextCompare)
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, """>", vbTextCompare)
      iIndex3 = InStr(iIndex, RichTextBox1.Text, "</td>", vbTextCompare)
      cmdDay(X).Caption = UCase(Mid(RichTextBox1.Text, iIndex + 2, ((iIndex3) - (iIndex + 2))))
      cmdDay(0).FontBold = True
   Next
   For X = 0 To 9
      'Month
      iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<div class=", vbTextCompare)
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, ">", vbTextCompare)
      iIndex3 = InStr(iIndex, RichTextBox1.Text, "</div>", vbTextCompare)
      lblTenDayM(X).Caption = Mid(RichTextBox1.Text, iIndex + 1, ((iIndex3) - (iIndex + 1))) ', "&deg;", Chr(176))
      'Date
      iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<div class=", vbTextCompare)
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, ">", vbTextCompare)
      iIndex3 = InStr(iIndex, RichTextBox1.Text, "</div>", vbTextCompare)
      lblTenDayD(X).Caption = Mid(RichTextBox1.Text, iIndex + 1, ((iIndex3) - (iIndex + 1)))
      'Image
      iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<img src=", vbTextCompare)
      iIndex3 = InStr(iIndex, RichTextBox1.Text, " alt=", vbTextCompare)
      imageUrl = Mid(RichTextBox1.Text, iIndexEnd + 10, (iIndex3 - 1) - (iIndexEnd + 10))
      SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), imgTenDay(X) '
      'Condition
      iIndexEnd = InStr(iIndex3, RichTextBox1.Text, " alt=", vbTextCompare)
      iIndex3 = InStr(iIndex, RichTextBox1.Text, " title=", vbTextCompare)
      lblTenDayCon(X).Caption = Mid(RichTextBox1.Text, iIndexEnd + 6, (iIndex3 - 1) - (iIndexEnd + 6))
      'Hi Degree
      iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<div class=", vbTextCompare)
      iIndex3 = InStr(iIndexEnd, RichTextBox1.Text, "</div>", vbTextCompare)
      lblTenDayH(X).Caption = Replace(Replace(Mid(RichTextBox1.Text, iIndexEnd + 12, (iIndex3) - (iIndexEnd + 12)), "&deg;", Chr(176)), """>", " ")
      'low Degree
      iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<div class=", vbTextCompare)
      iIndex3 = InStr(iIndexEnd, RichTextBox1.Text, "</div>", vbTextCompare)
      lblTenDayL(X).Caption = Replace(Replace(Mid(RichTextBox1.Text, iIndexEnd + 12, (iIndex3) - (iIndexEnd + 12)), "&deg;", Chr(176)), """>", " ")
   Next
   'Detail Day
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<strong>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<!-- Check if Day or Evening -->", vbTextCompare)
   sDayDetail = Mid(RichTextBox1.Text, iIndexEnd + 16, ((iIndex - 8) - (iIndexEnd + 16)))
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "-->", vbTextCompare)
   iIndex3 = InStr(iIndexEnd, RichTextBox1.Text, "</strong>", vbTextCompare)
   If iIndexEnd + 18 < iIndex3 - 12 Then
      sDayDetail = sDayDetail & Mid(RichTextBox1.Text, iIndexEnd + 18, ((iIndex3 - 12) - (iIndexEnd + 18)))
   Else
      sDayDetail = sDayDetail & vbCrLf & Mid(RichTextBox1.Text, iIndexEnd + 9, ((iIndex3 - 8) - (iIndexEnd + 9)))
   End If
   iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<br />", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<table", vbTextCompare)
   sDayDetail = sDayDetail & Replace(Mid(RichTextBox1.Text, iIndex3 + 20, ((iIndex) - (iIndex3 + 20))), "</strong>", "")
   lblDayDetail.Caption = Replace(Replace(sDayDetail, "<br />", ""), "<strong>", "")
   For X = 0 To 4
      'UV Detail condition
      iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<td>", vbTextCompare)
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<div class=", vbTextCompare)
      lblDetail(X).Caption = Mid(RichTextBox1.Text, iIndexEnd + 4, ((iIndex) - (iIndexEnd + 4)))
      'UV per
      iIndexEnd = InStr(iIndex, RichTextBox1.Text, "style=", vbTextCompare)
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, "><", vbTextCompare)
      uvPerct = Val(Mid(RichTextBox1.Text, iIndexEnd + 13, ((iIndex) - (iIndexEnd + 13))))
      iLeftpos = imgDetail(X).Left
      picDetail(X).Visible = False
      picDetail(X).Width = 1815
      'UV Image
      iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<img src=", vbTextCompare)
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, " alt=", vbTextCompare)
      imageUrl = Mid(RichTextBox1.Text, iIndexEnd + 10, ((iIndex - 1) - (iIndexEnd + 10)))
      SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), imgDetail(X)
      picDetail(X).Left = ((picDetail(X).Width * uvPerct) / 100) + iLeftpos
      picDetail(X).Width = (picDetail(X).Width * (100 - uvPerct)) / 100
      picDetail(X).Visible = True
   Next
   'Sunrise
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<br />", vbTextCompare)
   lblSunRise.Caption = Replace(Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex) - (iIndexEnd + 8))), "</strong>", "")
   'Sunset
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</div>", vbTextCompare)
   lbSunSet.Caption = Replace(Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex - 1) - (iIndexEnd + 8))), "</strong>", "")
   'Moonrise
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<br />", vbTextCompare)
   lblMoonRise.Caption = Replace(Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex) - (iIndexEnd + 8))), "</strong>", "")
   'Moonset
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</div>", vbTextCompare)
   lblMoonSet.Caption = Replace(Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex - 1) - (iIndexEnd + 8))), "</strong>", "")
   'MoonPhase Image
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<img src=", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, " alt=", vbTextCompare)
   imageUrl = Mid(RichTextBox1.Text, iIndexEnd + 10, ((iIndex - 1) - (iIndexEnd + 10)))
   SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), imgMoon
   'MoonPhse
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</strong>", vbTextCompare)
   lblMoonPhase.Caption = Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex) - (iIndexEnd + 8)))
   'Moon Waxing
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<br />", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</div>", vbTextCompare)
   lblWaxing.Caption = Mid(RichTextBox1.Text, iIndexEnd + 19, ((iIndex - 1) - (iIndexEnd + 19)))
   'Wind Image
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<img src=", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, " alt=", vbTextCompare)
   imageUrl = Mid(RichTextBox1.Text, iIndexEnd + 10, ((iIndex - 1) - (iIndexEnd + 10)))
   SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), imgWind
   'Wind Direction
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<span", vbTextCompare)
   sDayDetail = Replace(Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex) - (iIndexEnd + 8))), "</strong>", "")
   'Wind Degree
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "class=", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</span>", vbTextCompare)
   sDayDetail = sDayDetail & Replace(Mid(RichTextBox1.Text, iIndexEnd + 13, ((iIndex - 1) - (iIndexEnd + 13))), "&deg;", Chr(176))
   lblDirection.Caption = sDayDetail
   'Wind Speed
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<span", vbTextCompare)
   sDayDetail = Replace(Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex) - (iIndexEnd + 8))), "</strong>", "")
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "class=", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</span>", vbTextCompare)
   lblSpeed.Caption = sDayDetail & Space(1) & Mid(RichTextBox1.Text, iIndexEnd + 13, ((iIndex - 1) - (iIndexEnd + 13)))
   'Get Weather Aleart
   GetAlert
   
   If TView.Enabled = False Then
      TView.Enabled = True
      TView.SetFocus
   End If
   If InStr(1, TView.Nodes(curNameIndex).Tag, "US", vbTextCompare) <> 0 Then
      If Not zipButton Then
        sSelCountryName = Mid(TView.Nodes(curNameIndex).Tag, 3, 2)
      End If
   Else
      sSelCountryName = TView.Nodes(curNameIndex).Parent
   End If
   
   If Not zipButton Then
      sSelCityName = TView.Nodes(curNameIndex).Text
   End If
   If USAzipTreee Then
      If TView.Nodes(224).Expanded = True Then
        'DoEvents
        TView.Nodes(curNameIndex).Parent.Expanded = False
        TView.Nodes(224).Expanded = False
      End If
      TView.Nodes(225).Selected = True
    End If
  Set oFoundNode = Nothing
  zipButton = False
  Exit Sub
errorHandler:
  MsgBox "Unable To Display This Weather Report", vbCritical, "World Weather Program"
  TView.Enabled = True
End Sub

'Load Png (Bubbelbilden) to Image Control
Sub PngImageLoad(PathFilename As String, ImageControl As Image)
   Dim Token    As Long
    Token = InitGDIPlus
     ImageControl = LoadPictureGDIPlus(PathFilename, ImageControl.Width / Screen.TwipsPerPixelX, ImageControl.Height / Screen.TwipsPerPixelY)
    FreeGDIPlus Token
End Sub

Private Sub SavePngFille(myUrl As String, pngFile As String, picBox As Object)
  Dim nFileNum As Integer
  Dim myFile As String
  Dim myData() As Byte
  
  On Error Resume Next
  myData() = Inet1.OpenURL(myUrl, icByteArray)
 
  nFileNum = FreeFile
  Open App.Path + "\Icons\" & pngFile For Binary Access Write As #nFileNum
    Put #nFileNum, , myData()
  Close #nFileNum
  If Right(pngFile, 3) = "png" Then
    Call PngImageLoad(App.Path & "\Icons\" & pngFile, picBox)
  Else
    picBox.Picture = LoadPicture(App.Path & "\Icons\" & pngFile)
  End If
End Sub

Private Sub GetWebpage(Page As String)
  RichTextBox1.Text = ""
  RichTextBox1.Text = Inet1.OpenURL(Page)
End Sub

Function TreeViewFindNode(tvFind As TreeView, ByVal sFindItem As String, Optional bSearchAll As Boolean = True, Optional lItemIndex As Long = 1) As Node
    Dim oThisNode As Node, bSearch As Boolean, lInstance As Long
    
    sFindItem = UCase$(sFindItem)
    bSearch = True
    
    For Each oThisNode In tvFind.Nodes
        If bSearchAll = False Then
            'Only Search Top Level Nodes
            If (oThisNode.Parent Is Nothing) = False Then
                bSearch = False
            Else
                bSearch = True
            End If
        End If
        If bSearch Then
            If (UCase$(oThisNode.Text) Like sFindItem) = True Then
                lInstance = lInstance + 1
                If lInstance >= lItemIndex Then
                    'Found matching item
                    Set TreeViewFindNode = oThisNode
                    Exit For
                End If
            End If
        End If
    Next
End Function

Private Sub GetCountriesFlag(sWeblink As String, sCountry As String)
  Dim iIndex, iIndex2, iIndex3 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sRegion As String
  Dim sLimit As Integer
  Dim sCountryStat As String
  Dim Sovereign As String
  Dim sNames As String
  Dim sMoreFacts As String
  Dim sFactsBody As String
  Dim sMoreInfo As String
  Dim sExtraBody As String
  
  On Error GoTo errorHandler
  noFlags = False
  sPageName = sWeblink '"http://www.infoplease.com/ipa/A0107292.html"
  GetWebpage sPageName
  If sCountry = "Mexico" Then
    sStartPos = "Flag of "
    iIndexEnd = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
    GoTo noMap
  Else
    sStartPos = "<div class=""center"""
  End If
  DoEvents
  iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndex = 0 Then
    noFlags = True
    imgMap.Visible = True
    fmFlag.Caption = "No Country Flag"
    fmMap.Caption = "No Country Map"
    imgFlag.Visible = False
    imgPicture.Visible = False
    imgPicture.Enabled = False
    imgFlag.Enabled = False
    Exit Sub
  End If
  'Large Map
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "<a href=", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, " pageno=", vbTextCompare)
  sRegion = Mid(RichTextBox1.Text, iIndexSt + 9, (iIndexEnd - 1) - (iIndexSt + 9))
  LrgMapAddress = "http://www.infoplease.com" & sRegion
  'Map
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "src=", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, " ", vbTextCompare)
  sRegion = Mid(RichTextBox1.Text, iIndexSt + 5, (iIndexEnd - 1) - (iIndexSt + 5))
  sRegion = "http://i.infopls.com" & sRegion
  sMapPicture = App.Path + "\Icons\" & Mid(sRegion, InStrRev(sRegion, "/") + 1)
  'Flag
noMap:
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "src=", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, " ", vbTextCompare)
  sRegion = Mid(RichTextBox1.Text, iIndexSt + 5, (iIndexEnd - 1) - (iIndexSt + 5))
  If InStr(1, sRegion, "http://i.infopls.com", vbTextCompare) = 0 Then
    sRegion = "http://i.infopls.com" & sRegion
    SavePngFille sRegion, Mid(sRegion, InStrRev(sRegion, "/") + 1), imgFlag
    sFlagPicture = App.Path + "\Icons\" & Mid(sRegion, InStrRev(sRegion, "/") + 1)
  End If
  Exit Sub
errorHandler:
End Sub

Private Sub GetCountryFagMap(scntName As String)
  Dim nFileNum As Integer
  Dim sString As String
  Dim myArray() As String
  Dim cnt As Integer
  
  nFileNum = FreeFile
   Open App.Path + "\Map And Flags ByRegion New.txt" For Input As #nFileNum
   Do While Not EOF(nFileNum)
      Line Input #nFileNum, sString
      If Len(sString) > 1 Then
         myArray = Split(sString, ",")
         DoEvents
         If InStr(1, scntName, myArray(1), vbTextCompare) <> 0 Or InStr(1, myArray(1), scntName, vbTextCompare) <> 0 Then
            GetCountriesFlag myArray(0), myArray(1)
            sCountryUrl = myArray(0)
            cnt = cnt + 1
            'DoEvents
            Exit Do
         End If
      End If
   Loop
   If cnt = 0 Or noFlags Then
      imgMap.Visible = True
      fmFlag.Caption = "No Country Flag"
      fmMap.Caption = "No Country Map"
      Set imgFlag.Picture = imgMapFlag.ListImages(1).Picture
      Set imgMap.Picture = imgMapFlag.ListImages(2).Picture
      imgFlag.Visible = True
      imgPicture.Visible = False
      imgPicture.Enabled = False
      imgFlag.Enabled = False
      noFlags = True
   Else
      noFlags = False
      If myArray(1) = "Mexico" Then
        imgMap.Visible = True
        fmMap.Caption = "No Country Map"
        imgPicture.Visible = False
        imgPicture.Enabled = False
      Else
        imgMap.Visible = False
        fmMap.Caption = "Country Map"
        imgPicture.Visible = True
        imgPicture.Enabled = True
      End If
      fmFlag.Caption = "Country Flag"
      imgFlag.Visible = True
      'mnuCountryStat.Enabled = True
      imgFlag.Enabled = True
   End If
   Close #nFileNum
End Sub

Private Sub GetlargeMap()
  Dim iIndex, iIndex2, iIndex3 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sRegion As String
  
  On Error Resume Next
  
  GetWebpage LrgMapAddress
  sStartPos = " align=""center"""
  DoEvents
  iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndex = 0 Then Exit Sub
  'Large Map
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "<img src=", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, " border=", vbTextCompare)
  sRegion = Mid(RichTextBox1.Text, iIndexSt + 10, (iIndexEnd - 1) - (iIndexSt + 10))
  If InStr(1, sRegion, " ", vbTextCompare) <> 0 Then
    sRegion = Mid(sRegion, 1, InStr(1, sRegion, " ", vbTextCompare) - 2)
  End If
  picTureName = App.Path + "\Icons\" & Mid(sRegion, InStrRev(sRegion, "/") + 1)
  SavePngFille sRegion, Mid(sRegion, InStrRev(sRegion, "/") + 1), imgLrgMap
  SetPictureBox frmWeatherMain, picTureName, 0
End Sub

Public Function Check_Connection() As Boolean
  Dim result1 As Boolean
  
  result1 = InternetConnectionExists
  result1 = (InternetCheckConnection("http://www.microsoft.com/", FLAG_ICC_FORCE_CONNECTION, 0&) <> 0)
  If result1 Then
    itnetCon = True
    Check_Connection = True
  Else
    itnetCon = False
    Check_Connection = False
  End If
End Function

Public Function InternetConnectionExists() As Boolean
  InternetConnectionExists = (InternetAttemptConnect(ByVal 0&) = 0)
End Function

Private Sub GetRegion(sFindString As String)
   Dim nFileNum As Integer
   Dim sString As String
   Dim myArray() As String
   
   nFileNum = FreeFile
   Open App.Path & "\Region Cities All.Dat" For Binary Access Read As #nFileNum
   'On Error Resume Next
   Do While Not EOF(nFileNum)
      'read the length of the string
      Get #nFileNum, , nLen
      'initialize the string with the correct number of spaces
      sString = Space$(nLen)
      Get #nFileNum, , sString
      sString = DecryptText((sString), sPassword, True)
      If Len(Trim$(sString)) > 1 Then
         myArray = Split(sString, ",")
         If sFindString = myArray(2) Then
            StatusBar1.Panels(2).Text = "Listing For: " & lblCity.Caption & Space(4) & "Region: " & myArray(3)
            Exit Do
         End If
      End If
   Loop
   Close #nFileNum
End Sub

Private Sub GetDayDetails(sDyIndex As Integer, sStateCode As String)
  Dim iIndex, iIndex2, iIndex3 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sLimit As Integer
  Dim sPageName As String
  Dim sStartPos As String
  Dim bolEvn As Boolean
  Dim imageUrl As String
  Dim X As Integer
  Dim sDayDetail As String
  Dim uvPerct As Integer
  Dim iLeftpos As Integer
  
  On Error GoTo errorHandler
  
  If cmdFar.Enabled = False Then
    sPageName = "http://www.intellicast.com/Local/Weather.aspx?unit=F&location=" & sStateCode
  Else
    sPageName = "http://www.intellicast.com/Local/Weather.aspx?unit=C&location=" & sStateCode
  End If
  
  GetWebpage sPageName
  
  sStartPos = "Details for"
  DoEvents
  
  iIndex = 1
  For X = 0 To 9
    iIndexSt = InStr(iIndex, RichTextBox1.Text, sStartPos, vbTextCompare)
    If X = sDyIndex Then
      'Detail Day
      'Detail condition
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, sStartPos, vbTextCompare)
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<!-- Check if Day or Evening -->", vbTextCompare)
      sDayDetail = Mid(RichTextBox1.Text, iIndexEnd, ((iIndex - 8) - (iIndexEnd)))
      iIndexEnd = InStr(iIndex, RichTextBox1.Text, "-->", vbTextCompare)
      iIndex3 = InStr(iIndexEnd, RichTextBox1.Text, "</strong>", vbTextCompare)
      If iIndexEnd + 18 < iIndex3 - 12 Then
        sDayDetail = sDayDetail & Mid(RichTextBox1.Text, iIndexEnd + 18, ((iIndex3 - 12) - (iIndexEnd + 18)))
        bolEvn = True
      Else
        sDayDetail = sDayDetail & Mid(RichTextBox1.Text, iIndexEnd + 17, ((iIndex3) - (iIndexEnd + 17)))
      End If
      iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<br />", vbTextCompare)
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<table", vbTextCompare)
      If bolEvn Then
        sDayDetail = sDayDetail & Replace(Mid(RichTextBox1.Text, iIndexEnd + 12, ((iIndex) - (iIndexEnd + 12))), "<br />", "")
      Else
        sDayDetail = sDayDetail & vbCrLf & Replace(Mid(RichTextBox1.Text, iIndexEnd + 12, ((iIndex) - (iIndexEnd + 12))), "<br />", "")
      End If
      sDayDetail = Replace(Replace(Replace(sDayDetail, "</strong>", ""), "<strong>", ""), "  ", "")
      lblDayDetail.Caption = sDayDetail
      For sLimit = 0 To 4
        'UV Detail condition
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<td>", vbTextCompare)
        iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<div class=", vbTextCompare)
        lblDetail(sLimit).Caption = Mid(RichTextBox1.Text, iIndexEnd + 4, ((iIndex) - (iIndexEnd + 4)))
        'UV per
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "style=", vbTextCompare)
        iIndex = InStr(iIndexEnd, RichTextBox1.Text, "><", vbTextCompare)
        uvPerct = Val(Mid(RichTextBox1.Text, iIndexEnd + 13, ((iIndex) - (iIndexEnd + 13))))
        iLeftpos = imgDetail(sLimit).Left
        picDetail(sLimit).Visible = False
        picDetail(sLimit).Width = 1815
        'UV Image
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<img src=", vbTextCompare)
        iIndex = InStr(iIndexEnd, RichTextBox1.Text, " alt=", vbTextCompare)
        imageUrl = Mid(RichTextBox1.Text, iIndexEnd + 10, ((iIndex - 1) - (iIndexEnd + 10)))
        SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), imgDetail(sLimit)
        picDetail(sLimit).Left = ((picDetail(sLimit).Width * uvPerct) / 100) + iLeftpos
        picDetail(sLimit).Width = (picDetail(sLimit).Width * (100 - uvPerct)) / 100
        picDetail(sLimit).Visible = True
      Next
      Exit For
    End If
    iIndex3 = InStr(iIndexSt, RichTextBox1.Text, "<!--", vbTextCompare)
    iIndex = iIndex3
  Next
  'Sunrise
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<br />", vbTextCompare)
  lblSunRise.Caption = Replace(Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex) - (iIndexEnd + 8))), "</strong>", "")
  'Sunset
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</div>", vbTextCompare)
  lbSunSet.Caption = Replace(Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex - 1) - (iIndexEnd + 8))), "</strong>", "")
  'Moonrise
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<br />", vbTextCompare)
  lblMoonRise.Caption = Replace(Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex) - (iIndexEnd + 8))), "</strong>", "")
  'Moonset
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</div>", vbTextCompare)
  lblMoonSet.Caption = Replace(Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex - 1) - (iIndexEnd + 8))), "</strong>", "")
  'MoonPhase Image
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<img src=", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, " alt=", vbTextCompare)
  imageUrl = Mid(RichTextBox1.Text, iIndexEnd + 10, ((iIndex - 1) - (iIndexEnd + 10)))
  SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), imgMoon
  'MoonPhse
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</strong>", vbTextCompare)
  lblMoonPhase.Caption = Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex) - (iIndexEnd + 8)))
  'Moon Waxing
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<br />", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</div>", vbTextCompare)
  lblWaxing.Caption = Mid(RichTextBox1.Text, iIndexEnd + 19, ((iIndex - 1) - (iIndexEnd + 19)))
  'Wind Image
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<img src=", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, " alt=", vbTextCompare)
  imageUrl = Mid(RichTextBox1.Text, iIndexEnd + 10, ((iIndex - 1) - (iIndexEnd + 10)))
  SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), imgWind
  'Wind Direction
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<span", vbTextCompare)
  sDayDetail = Replace(Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex) - (iIndexEnd + 8))), "</strong>", "")
  'Wind Degree
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "class=", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</span>", vbTextCompare)
  sDayDetail = sDayDetail & Replace(Mid(RichTextBox1.Text, iIndexEnd + 13, ((iIndex - 1) - (iIndexEnd + 13))), "&deg;", Chr(176))
  lblDirection.Caption = sDayDetail
  'Wind Speed
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<span", vbTextCompare)
  sDayDetail = Replace(Mid(RichTextBox1.Text, iIndexEnd + 8, ((iIndex) - (iIndexEnd + 8))), "</strong>", "")
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "class=", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</span>", vbTextCompare)
  lblSpeed.Caption = sDayDetail & Space(1) & Mid(RichTextBox1.Text, iIndexEnd + 13, ((iIndex - 1) - (iIndexEnd + 13)))
  If TView.Enabled = False Then
    TView.Enabled = True
    TView.SetFocus
  End If
  Exit Sub
errorHandler:
  MsgBox "Unable To Display This Weather Report", vbCritical, "World Weather Program"
  TView.Enabled = True
End Sub

Private Sub LoadComboBox()
   Dim nFileNum As Integer
   Dim sString As String
   Dim myArray() As String
  
   nFileNum = FreeFile
   Open App.Path & "\USA Zipcode & City.Dat" For Binary Access Read As #nFileNum
   'On Error Resume Next
   Do While Not EOF(nFileNum)
      'read the length of the string
      Get #nFileNum, , nLen
      'initialize the string with the correct number of spaces
      sString = Space$(nLen)
      Get #nFileNum, , sString
      sString = DecryptText((sString), sPassword, True)
      If Len(Trim$(sString)) > 1 Then
         myArray() = Split(sString, ",")
         DoEvents
         cmboZip.AddItem myArray(0)
         'DoEvents
      End If
   Loop
   Close #nFileNum
End Sub

Function TreeFindNode(tvFind As TreeView, ByVal sFindItem As String, Optional bSearchAll As Boolean = True, Optional lItemIndex As Long = 1) As Node
   Dim oThisNode As Node, bSearch As Boolean, lInstance As Long
    
   sFindItem = UCase$(sFindItem)
   bSearch = True
    
   For Each oThisNode In tvFind.Nodes
      If bSearchAll = False Then
         'Only Search Top Level Nodes
         If (oThisNode.Parent Is Nothing) = False Then
            bSearch = False
         Else
            bSearch = True
         End If
      End If
      If bSearch Then
         If (UCase$(oThisNode.Text) Like sFindItem) = True And sCityCode = oThisNode.Tag Then
            lInstance = lInstance + 1
            If lInstance >= lItemIndex Then
               'Found matching item
               curNameIndex = oThisNode.Index
               Set TreeFindNode = oThisNode
               bNodeFound = True
               Exit For
            End If
         Else
            bNodeFound = False
         End If
      End If
   Next
End Function

'Decrypt text encrypted with EncryptText
Public Function DecryptText(strText As String, ByVal strPwd As String, CASE_SENSITIVE_PASSWORD As Boolean)
   Dim I As Integer, C As Integer
   Dim strBuff As String
  
   If Not CASE_SENSITIVE_PASSWORD Then
      'Convert password to upper case
      'if not case-sensitive
      strPwd = UCase$(strPwd)
   End If
  
   'Decrypt string
   If Len(strPwd) Then
      For I = 1 To Len(strText)
         C = Asc(Mid$(strText, I, 1))
         C = C - Asc(Mid$(strPwd, (I Mod Len(strPwd)) + 1, 1))
         strBuff = strBuff & Chr$(C And &HFF)
      Next I
   Else
      strBuff = strText
   End If
   DecryptText = strBuff
End Function

Private Sub GetWeatherAlert()
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iCode, iCode1 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim sPageName As String
   Dim sStartPos As String
   Dim imageUrl As String
   Dim sHeading As String
   Dim sSubHeading As String
   
   On Error Resume Next
   sPageName = "http://www.intellicast.com/Local/Weather.aspx?location=" & sCountryCode
   GetWebpage sPageName
   sStartPos = "Local Information"
   DoEvents
  
   iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
   iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "<a href=", vbTextCompare)
   iCode = InStr(iIndex2, RichTextBox1.Text, "/Severe/", vbTextCompare)
   iCode1 = InStr(iCode, RichTextBox1.Text, "=", vbTextCompare)
   iCode = InStr(iCode1, RichTextBox1.Text, ">", vbTextCompare)
   'State code
   sStateBoxCode = Mid(RichTextBox1.Text, iCode1 + 1, (iCode - 1) - (iCode1 + 1))
   If Not bNodeFound Then
      sCountryCode = sStateBoxCode
   End If
   iIndex3 = InStr(iIndexSt, RichTextBox1.Text, "<img", vbTextCompare)
   If InStr(1, Mid(RichTextBox1.Text, iIndex2 + 1, (iIndex3) - (iIndex2 + 1)), "/Severe/", vbTextCompare) <> 0 Then
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "<strong", vbTextCompare)
      iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, ">", vbTextCompare)
      iIndex = InStr(iIndexSt, RichTextBox1.Text, "</strong", vbTextCompare)
      If InStr(1, Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1)), "No", vbTextCompare) = 0 Then
         sCityName = Mid(lblCity.Caption, 1, InStr(1, lblCity, ",", vbTextCompare) - 1)
         If MsgBox(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1)) & vbCrLf & "View current alerts?", vbDefaultButton2 + vbQuestion + vbYesNo, sCityName & " Weather Alert") = vbYes Then
            sPageName = "http://www.intellicast.com/Storm/Severe/Bulletins.aspx?location=" & sCountryCode
            GetWebpage sPageName
            sStartPos = "Weather Alerts:"
            'DoEvents
  
            iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
            iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</div>", vbTextCompare)
            sHeading = Mid(RichTextBox1.Text, iIndexSt, (iIndexEnd) - (iIndexSt))
            iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<strong class=", vbTextCompare)
            iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
            iIndexSt = InStr(iIndex, RichTextBox1.Text, "</strong>", vbTextCompare)
            sSubHeading = sHeading & vbCrLf & Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1)) '
            
            iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "<br/>", vbTextCompare)
            iIndex = InStr(iIndexEnd, RichTextBox1.Text, ">$$<", vbTextCompare)
            frmAlert.txtAlert.Text = sSubHeading & vbCrLf & Replace(Replace(Replace(Mid(RichTextBox1.Text, iIndexEnd + 5, (iIndex - 11) - (iIndexEnd + 5)), "<br />", vbCrLf), "-", ", "), "  ", " ")
            iMinCount = 0
            frmAlert.txtAlert.Visible = True
            frmAlert.Caption = sCityName & " Weather Alert"
            frmAlert.Show vbModal
         End If
      Else
         MsgBox Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1)), "<br/>", " "), vbInformation, sCityName & " Weather Alert"
      End If
   Else
      MsgBox "No Weather Alerts for this location", vbInformation, sCityName & " Weather Alert"
   End If
End Sub

Private Sub GetAlert()
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim X As Integer
   
   On Error Resume Next
   iIndex3 = InStr(1, RichTextBox1.Text, "Local Information", vbTextCompare)
   iIndex2 = InStr(iIndex3, RichTextBox1.Text, "<img", vbTextCompare)
   If InStr(1, Mid(RichTextBox1.Text, iIndex3 + 1, (iIndex2) - (iIndex3 + 1)), "/Severe/", vbTextCompare) <> 0 Then
      iIndexEnd = InStr(iIndex2, RichTextBox1.Text, "<strong", vbTextCompare)
      If iIndexEnd = 0 Then GoTo TrackList
      iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, ">", vbTextCompare)
      iIndex = InStr(iIndexSt, RichTextBox1.Text, "</strong", vbTextCompare)
      If InStr(1, Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1)), "No", vbTextCompare) = 0 Then
         mnuWeather.Visible = True
      Else
         mnuWeather.Visible = False
      End If
   Else
      mnuWeather.Visible = False
   End If
   
  iIndexSt = InStr(1, RichTextBox1.Text, "Tropical Storm Tracking", vbTextCompare)
  If iIndexSt <> 0 Then
    iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "Alert", vbTextCompare)
    iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
    iIndex3 = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
    menuActiveStorm.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndex3) - (iIndex + 1))
    If Trim(Len(menuActiveStorm.Caption)) = 0 Then Exit Sub
    menuActiveStorm.Visible = True
    iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<img src=", vbTextCompare)
    For X = 0 To 2
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, "href=", vbTextCompare)
      iIndexEnd = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
      mnuCurTract(X).Tag = Mid(RichTextBox1.Text, iIndex + 6, (iIndexEnd - 1) - (iIndex + 6))
    Next
    
    iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "Title", vbTextCompare)
    
    If InStr(1, Mid(RichTextBox1.Text, iIndex2, 100), "Bulletins:", vbTextCompare) <> 0 Then
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, "href=", vbTextCompare)
      iIndexEnd = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
      mnuStormList.Tag = Mid(RichTextBox1.Text, iIndex + 6, (iIndexEnd - 1) - (iIndex + 6))
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, " ", vbTextCompare)
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
      mnuStormList.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
      mnuStormList.Visible = True
      bStormBulletins = True
    End If
    
    If InStr(1, Mid(RichTextBox1.Text, iIndex2, 100), "Bulletins:", vbTextCompare) = 0 Then
      iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "Alert", vbTextCompare)
      iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
      iIndex3 = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
      mnuStorm2.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndex3) - (iIndex + 1))
      mnuStorm2.Visible = True
   
      iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<img src=", vbTextCompare)
      For X = 0 To 2
        iIndex = InStr(iIndexEnd, RichTextBox1.Text, "href=", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
        mnuInfrared(X).Tag = Mid(RichTextBox1.Text, iIndex + 6, (iIndexEnd - 1) - (iIndex + 6))
      Next
    Else
      Exit Sub
    End If
    
    iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "Alert Bold", vbTextCompare)
    If iIndex2 = 0 Then GoTo TrackList 'Exit Sub
    If InStr(1, Mid(RichTextBox1.Text, iIndex2, 100), "/Storm/Hurricane/Active.aspx", vbTextCompare) <> 0 Then
      iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
      iIndex3 = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
      mnuStorm3.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndex3) - (iIndex + 1))
      mnuStorm3.Visible = True
   
      iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "<img src=", vbTextCompare)
      For X = 0 To 2
        iIndex = InStr(iIndexEnd, RichTextBox1.Text, "href=", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
        mnuActiveHurricane(X).Tag = Mid(RichTextBox1.Text, iIndex + 6, (iIndexEnd - 1) - (iIndex + 6))
      Next
    End If
TrackList:
    If InStr(iIndexEnd, RichTextBox1.Text, "Storm Track List", vbTextCompare) <> 0 Then
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, "href=", vbTextCompare)
      iIndexEnd = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
      mnuStormList.Tag = Mid(RichTextBox1.Text, iIndex + 6, (iIndexEnd - 1) - (iIndex + 6))
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, " &", vbTextCompare)
      mnuStormList.Caption = Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex) - (iIndexEnd + 1))
      mnuStormList.Visible = True
      bStormBulletins = True
    End If
  End If
End Sub

Private Sub GetHurricane()
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim cnt As Integer
   
   iIndex3 = InStr(1, RichTextBox1.Text, ">Severe Weather</a>", vbTextCompare)
   Do While cnt < 9
      iIndex2 = InStr(iIndex3, RichTextBox1.Text, "href=", vbTextCompare)
      iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
      mnuSevere(cnt).Tag = Mid(RichTextBox1.Text, iIndex2 + 6, (iIndex - 1) - (iIndex2 + 6))
      iIndex3 = InStr(iIndex, RichTextBox1.Text, "</a>", vbTextCompare)
      mnuSevere(cnt).Caption = Replace(Mid(RichTextBox1.Text, iIndex + 1, (iIndex3) - (iIndex + 1)), "&amp;", Chr(38))
      cnt = cnt + 1
   Loop
   cnt = 0
   iIndex3 = InStr(1, RichTextBox1.Text, "Hurricane", vbTextCompare)
   Do While cnt < 10
      iIndex2 = InStr(iIndex3, RichTextBox1.Text, "href=", vbTextCompare)
      iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
      mnuHur(cnt).Tag = Mid(RichTextBox1.Text, iIndex2 + 6, (iIndex - 1) - (iIndex2 + 6))
      iIndex3 = InStr(iIndex, RichTextBox1.Text, "</a>", vbTextCompare)
      mnuHur(cnt).Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndex3) - (iIndex + 1))
      cnt = cnt + 1
   Loop
   cnt = 0
   iIndex3 = InStr(1, RichTextBox1.Text, "Season Summaries", vbTextCompare)
   Do While cnt < 12
      iIndex2 = InStr(iIndex3, RichTextBox1.Text, "href=", vbTextCompare)
      iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
      mnuHS(cnt).Tag = Mid(RichTextBox1.Text, iIndex2 + 6, (iIndex - 1) - (iIndex2 + 6))
      iIndex3 = InStr(iIndex, RichTextBox1.Text, "</a>", vbTextCompare)
      mnuHS(cnt).Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndex3) - (iIndex + 1))
      cnt = cnt + 1
   Loop
   cnt = 0
   iIndexSt = InStr(iIndex3, RichTextBox1.Text, "Satellite", vbTextCompare)
   iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
   iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
   mnuSatGlobal.Tag = Mid(RichTextBox1.Text, iIndex2 + 6, (iIndex - 1) - (iIndex2 + 6))
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "</a>", vbTextCompare)
   mnuSatGlobal.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
   Do While cnt < 5
      iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
      iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
      mnuSat(cnt).Tag = Mid(RichTextBox1.Text, iIndex2 + 6, (iIndex - 1) - (iIndex2 + 6))
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "</a>", vbTextCompare)
      mnuSat(cnt).Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
      cnt = cnt + 1
   Loop
   
   'Visible Satellite
   iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
   iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
   mnuVisibleSatellite.Tag = Mid(RichTextBox1.Text, iIndex2 + 6, (iIndex - 1) - (iIndex2 + 6))
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "</a>", vbTextCompare)
   mnuVisibleSatellite.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
   'Current Satellite
   iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
   iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
   mnuCurrentSatellite.Tag = Mid(RichTextBox1.Text, iIndex2 + 6, (iIndex - 1) - (iIndex2 + 6))
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "</a>", vbTextCompare)
   mnuCurrentSatellite.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
   'Water Vaper
   iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
   iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
   mnuWaterVaper.Tag = Mid(RichTextBox1.Text, iIndex2 + 6, (iIndex - 1) - (iIndex2 + 6))
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "</a>", vbTextCompare)
   mnuWaterVaper.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
End Sub

Private Sub GetHurricaneSumMap(sHurLink As String)
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim sPageName As String
   Dim sStartPos As String
   Dim sRegion As String
  
   On Error Resume Next
   GetWebpage "http://www.intellicast.com" & sHurLink
   If "/Storm/Summary/Hurricane1998.aspx" = sHurLink Then
      sRegion = "http://images.intellicast.com/WxImages/CustomGraphic/hursum98.gif"
   ElseIf "/Storm/Summary/Hurricane1999.aspx" = sHurLink Then
      sRegion = "http://images.intellicast.com/WxImages/CustomGraphic/hursum99.gif"
   Else
      sStartPos = "Hurricane Summary Maps"
      'DoEvents
      iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
      If iIndexSt = 0 Then Exit Sub
      'Large Map
      iIndex = InStr(iIndexSt, RichTextBox1.Text, sHurLink, vbTextCompare)
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "<img src=", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, " alt", vbTextCompare)
   End If
   sRegion = Mid(RichTextBox1.Text, iIndexSt + 10, (iIndexEnd - 1) - (iIndexSt + 10))
   If InStr(1, sRegion, " ", vbTextCompare) <> 0 Then
      sRegion = Mid(sRegion, 1, InStr(1, sRegion, " ", vbTextCompare) - 2)
   End If
   picTureName = App.Path + "\Icons\" & Mid(sRegion, InStrRev(sRegion, "/") + 1)
   SavePngFille sRegion, Mid(sRegion, InStrRev(sRegion, "/") + 1), imgLrgMap
   Load frmCountry
   If Animation Then
      GetAnimation sHurLink, sStartPos
   End If
End Sub

Private Sub GetHurricaneMap(sHurLink As String)
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim sPageName As String
   Dim sStartPos As String
   Dim sRegion As String
  
   On Error Resume Next
   GetWebpage "http://www.intellicast.com" & sHurLink
   sStartPos = "Hurricane Maps"
   'DoEvents
   PlayAnimation = False
   If InStr(1, RichTextBox1.Text, "Play Animation", vbTextCompare) <> 0 Then
      PlayAnimation = True
   End If
   iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
   If iIndexSt = 0 Then Exit Sub
   'Large Map
   iIndex = InStr(iIndexSt, RichTextBox1.Text, "Content Container", vbTextCompare)
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "src=", vbTextCompare)
   iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, " alt", vbTextCompare)
   sRegion = Mid(RichTextBox1.Text, iIndexSt + 5, (iIndexEnd - 1) - (iIndexSt + 5))
   If InStr(1, sRegion, " ", vbTextCompare) <> 0 Then
      sRegion = Mid(sRegion, 1, InStr(1, sRegion, " ", vbTextCompare) - 2)
   End If
   picTureName = App.Path + "\Icons\" & Mid(sRegion, InStrRev(sRegion, "/") + 1)
   SavePngFille sRegion, Mid(sRegion, InStrRev(sRegion, "/") + 1), imgLrgMap
   Load frmCountry
   If Animation Then
      GetAnimation sHurLink, sStartPos
   End If
End Sub

Private Sub GetHurricaneTrack(sHurLink As String)
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim sPageName As String
   Dim sStartPos As String
   Dim sRegion As String
   Dim sHeading As String
   Dim sHeading1 As String
   Dim sHeading2 As String
   Dim Limits As Integer
   Dim sStormName As String
   Dim cnt As Integer
   Dim bfrsRow As Boolean
   
   On Error Resume Next
   GetWebpage "http://www.intellicast.com" & sHurLink
   sStartPos = "Active Storm Track" '"Hurricane Season"
   
   iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
   If iIndexSt = 0 Then GoTo stHurseason
   
   'get large map 1
   iIndex = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
   iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
   slargeMapLink1 = "http://www.intellicast.com" & Mid(RichTextBox1.Text, iIndex + 6, (iIndexSt - 1) - (iIndex + 6))
   
   iIndex = InStr(iIndexSt, RichTextBox1.Text, "<img src=", vbTextCompare)
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "style=", vbTextCompare)
   sRegion = Mid(RichTextBox1.Text, iIndex + 10, (iIndexSt - 2) - (iIndex + 10))
   If InStr(1, sRegion, "_200w/CustomGraphic/", vbTextCompare) = 0 Then GoTo stHurseason
   SavePngFille sRegion, Mid(sRegion, InStrRev(sRegion, "/") + 1), frmAlert.picHur1
   
   iIndex = InStr(iIndexSt, RichTextBox1.Text, "alt=", vbTextCompare)
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "/>", vbTextCompare)
   frmAlert.lblHur1.Caption = Mid(RichTextBox1.Text, iIndex + 5, (iIndexSt - 2) - (iIndex + 5))
   
   frmAlert.Picture1.Visible = True
   frmAlert.picHur1.Visible = True
   
   'get large map 1
   iIndex = InStr(iIndexSt, RichTextBox1.Text, "<img src", vbTextCompare)
   iIndexSt = InStrRev(RichTextBox1.Text, "href=", iIndex, vbTextCompare)
   slargeMapLink2 = "http://www.intellicast.com" & Mid(RichTextBox1.Text, iIndexSt + 6, (iIndex - 2) - (iIndexSt + 6))
   
   iIndex = InStr(iIndexSt, RichTextBox1.Text, "<img src=", vbTextCompare)
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "style=", vbTextCompare)
   sRegion = Mid(RichTextBox1.Text, iIndex + 10, (iIndexEnd - 2) - (iIndex + 10))
   If InStr(1, sRegion, "_200w/CustomGraphic/", vbTextCompare) = 0 Then GoTo stHurseason
   SavePngFille sRegion, Mid(sRegion, InStrRev(sRegion, "/") + 1), frmAlert.picHur2
   
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "alt=", vbTextCompare)
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "/>", vbTextCompare)
   frmAlert.lblHur2.Caption = Mid(RichTextBox1.Text, iIndex + 5, (iIndexSt - 2) - (iIndex + 5))
   If InStr(1, Mid(RichTextBox1.Text, iIndex + 5, (iIndexSt - 2) - (iIndex + 5)), "Add", vbTextCompare) = 0 Then
    frmAlert.picHur2.Visible = True
    frmAlert.Picture2.Visible = True
   End If
stHurseason:
   iIndexSt = InStr(1, RichTextBox1.Text, "Hurricane Season", vbTextCompare)
   If iIndexSt = 0 Then Exit Sub
   'Heading
   iIndex = InStr(iIndexSt, RichTextBox1.Text, "Content Container", vbTextCompare)
   iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
   iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "<br /><br />", vbTextCompare)
   sRegion = "Hurricane Season" & vbCrLf & Mid(RichTextBox1.Text, iIndexSt + 5, (iIndexEnd - 1) - (iIndexSt + 5))
   'Heading
   iIndex = InStr(iIndexSt, RichTextBox1.Text, "<strong>", vbTextCompare)
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "</strong><br />", vbTextCompare)
   sHeading = Mid(RichTextBox1.Text, iIndex + 8, (iIndexSt) - (iIndex + 8))
   'Information
   iIndex = InStr(iIndexSt, RichTextBox1.Text, "</table>", vbTextCompare)
   iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</div>", vbTextCompare)
   iIndex3 = InStr(iIndexEnd, RichTextBox1.Text, "<br /><br />", vbTextCompare)
   sHeading = sHeading & vbCrLf & Replace(Mid(RichTextBox1.Text, iIndexEnd + 10, (iIndex3) - (iIndexEnd + 10)), "<br />", vbCrLf)
   'Names
   iIndex = InStr(iIndex3, RichTextBox1.Text, "<strong>", vbTextCompare)
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "</strong><br />", vbTextCompare)
   sHeading2 = Mid(RichTextBox1.Text, iIndex + 8, (iIndexSt) - (iIndex + 8))
   iIndex = InStr(iIndexSt, RichTextBox1.Text, "<div>", vbTextCompare)
   sHeading2 = sHeading2 & vbCrLf & Mid(RichTextBox1.Text, iIndexSt + 19, (iIndex) - (iIndexSt + 19))
   iIndex2 = InStr(iIndex, RichTextBox1.Text, "<ul", vbTextCompare)
   Do
      iIndexEnd = InStr(iIndex2, RichTextBox1.Text, "<li>", vbTextCompare)
      iIndex3 = InStr(iIndexEnd, RichTextBox1.Text, "</li>", vbTextCompare)
      If cnt = 0 And bfrsRow = False Then
         frmAlert.lsvStormName.ListItems.Add , , Chr(42) & " " & Mid(RichTextBox1.Text, iIndexEnd + 4, (iIndex3) - (iIndexEnd + 4))
      End If
      
      If bfrsRow Then
         cnt = cnt + 1
         frmAlert.lsvStormName.ListItems(cnt).ListSubItems.Add , , Chr(42) & " " & Mid(RichTextBox1.Text, iIndexEnd + 4, (iIndex3) - (iIndexEnd + 4))
      End If
      If InStr(1, Mid(RichTextBox1.Text, iIndex3, 20), "</ul>", vbTextCompare) <> 0 Then
         cnt = 0
         bfrsRow = True
         frmAlert.lsvStormName.ListItems(cnt).ListSubItems.Add , , Chr(42) & " " & Mid(RichTextBox1.Text, iIndexEnd + 4, (iIndex3) - (iIndexEnd + 4))
      End If
      If InStr(1, Mid(RichTextBox1.Text, iIndex3, 40), " </div>", vbTextCompare) <> 0 Then
         Limits = 1
      End If
      iIndex2 = iIndex3
   Loop Until Limits = 1
   frmAlert.txtAlert.Visible = True
   frmAlert.txtAlert.Text = sRegion & vbCrLf & sHeading & vbCrLf & sHeading2 & vbCrLf & sStormName
   iMinCount = 0
   frmAlert.Caption = "Active Track"
   frmAlert.Show vbModal
End Sub

Private Sub GetWeatherAlerts(sHurLink As String)
  Dim iIndex, iIndex2, iIndex3 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sfrsRun As Boolean
  Dim intCount As Integer
  Dim iLinecount As Integer
  Dim Limits As Integer
  Dim cnt As Integer
  Dim bfrsRow As Boolean
  Dim sNewString As String
   
  On Error Resume Next
  GetWebpage "http://www.intellicast.com" & sHurLink
  sStartPos = "Weather Alerts:"
  
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then Exit Sub
  Do
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "style=", vbTextCompare)
    iIndex3 = InStr(iIndexEnd, RichTextBox1.Text, ">", vbTextCompare)
    If InStr(1, Mid(RichTextBox1.Text, iIndexEnd + 7, (iIndex3 - 2) - (iIndexEnd + 7)), "color:#900", vbTextCompare) <> 0 Then
      bfrsRow = True
    Else
      bfrsRow = False
    End If
    iIndex = InStr(iIndex3, RichTextBox1.Text, "</a>", vbTextCompare)
    If cnt = 0 Then
      iLinecount = iLinecount + 1
      If bfrsRow Then
        frmAlert.lstWeatherAlert.ListItems.Add , , Chr(42) & " " & Mid(RichTextBox1.Text, iIndex3 + 1, (iIndex) - (iIndex3 + 1))
      Else
        frmAlert.lstWeatherAlert.ListItems.Add , , Mid(RichTextBox1.Text, iIndex3 + 1, (iIndex) - (iIndex3 + 1))
      End If
      cnt = cnt + iLinecount
    End If
     
    If cnt <> 0 And sfrsRun Then
      If bfrsRow Then
        frmAlert.lstWeatherAlert.ListItems(cnt).ListSubItems.Add , , Chr(42) & " " & Mid(RichTextBox1.Text, iIndex3 + 1, (iIndex) - (iIndex3 + 1))
      Else
        frmAlert.lstWeatherAlert.ListItems(cnt).ListSubItems.Add , , Mid(RichTextBox1.Text, iIndex3 + 1, (iIndex) - (iIndex3 + 1))
      End If
    End If
    sfrsRun = True
    intCount = intCount + 1
    If intCount > 2 Then
      intCount = 0
      cnt = 0
      sfrsRun = False
    End If
    If InStr(1, Mid(RichTextBox1.Text, iIndex, 40), "</table></div>", vbTextCompare) <> 0 Then
      Limits = 1
    End If
    iIndexSt = iIndex
  Loop Until Limits = 1
  
  frmAlert.lstWeatherAlert.ListItems.Add , , ""
  iLinecount = iLinecount + 1
  frmAlert.lstWeatherAlert.ListItems.Add , , ""
  iLinecount = iLinecount + 1
  
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "Title", vbTextCompare)
  iIndex3 = InStr(iIndexEnd, RichTextBox1.Text, ">", vbTextCompare)
  iIndexSt = InStr(iIndex3, RichTextBox1.Text, "</", vbTextCompare)
  frmAlert.lstWeatherAlert.ListItems(iLinecount).ListSubItems.Add , , Mid(RichTextBox1.Text, iIndex3 + 1, (iIndexSt) - (iIndex3 + 1))
  frmAlert.lstWeatherAlert.ListItems.Add , , ""
  iLinecount = iLinecount + 1
  Do
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "strong", vbTextCompare)
    iIndex3 = InStr(iIndexEnd, RichTextBox1.Text, ">", vbTextCompare)
    iIndexSt = InStr(iIndex3, RichTextBox1.Text, "</", vbTextCompare)
    frmAlert.lstWeatherAlert.ListItems.Add , , Mid(RichTextBox1.Text, iIndex3 + 1, (iIndexSt) - (iIndex3 + 1))
    iLinecount = iLinecount + 1
    
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
    iIndex3 = InStr(iIndexEnd, RichTextBox1.Text, "</div>", vbTextCompare)
    
    sNewString = GetStateAlertType(Mid(RichTextBox1.Text, iIndexEnd, (iIndex3) - (iIndexEnd)), iLinecount)
    
    If Len(sNewString) > 23 Then
      frmAlert.lstWeatherAlert.ListItems(iLinecount).ListSubItems.Add , , Mid(sNewString, 1, 23)
      frmAlert.lstWeatherAlert.ListItems(iLinecount).ListSubItems.Add , , Mid(sNewString, 24, 22)
      If Len(Mid(sNewString, 45)) <> 0 Then
        frmAlert.lstWeatherAlert.ListItems.Add , , Space(Len(frmAlert.lstWeatherAlert.ListItems(iLinecount).Text)) & "-"
        iLinecount = iLinecount + 1
        frmAlert.lstWeatherAlert.ListItems(iLinecount).ListSubItems.Add , , Mid(sNewString, 46)
      End If
    Else
      frmAlert.lstWeatherAlert.ListItems(iLinecount).ListSubItems.Add , , sNewString
    End If
    iIndexSt = iIndex3
    If InStr(1, Mid(RichTextBox1.Text, iIndex3, 100), "</td>", vbTextCompare) <> 0 Then
      Exit Do
    End If
  Loop
  frmAlert.lstWeatherAlert.Visible = True
  frmAlert.txtAlert.Visible = False
  frmAlert.Caption = "Weather Alerts: United State"
  frmAlert.Show vbModal
End Sub

Private Sub GetSevereWeatherMap(sHurLink As String, sLinkStart As String)
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim sPageName As String
   Dim sStartPos As String
   Dim sRegion As String
   On Error Resume Next
   GetWebpage "http://www.intellicast.com" & sHurLink ' & "?animate=true"
   PlayAnimation = False
   If InStr(1, RichTextBox1.Text, "Play Animation", vbTextCompare) <> 0 Then
      PlayAnimation = True
   End If
   sStartPos = sLinkStart
   'DoEvents
   iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
   If iIndexSt = 0 Then Exit Sub
   'Large Map
   iIndex3 = InStr(iIndexSt, RichTextBox1.Text, "Content Container", vbTextCompare)
   iIndex = InStr(iIndex3, RichTextBox1.Text, "<img id=", vbTextCompare)
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "src=", vbTextCompare)
   iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, " alt", vbTextCompare)
   sRegion = Mid(RichTextBox1.Text, iIndexSt + 5, (iIndexEnd - 1) - (iIndexSt + 5))
   If InStr(1, sRegion, " ", vbTextCompare) <> 0 Then
      sRegion = Mid(sRegion, 1, InStr(1, sRegion, " ", vbTextCompare) - 2)
   End If
   picTureName = App.Path + "\Icons\" & Mid(sRegion, InStrRev(sRegion, "/") + 1)
   SavePngFille sRegion, Mid(sRegion, InStrRev(sRegion, "/") + 1), imgLrgMap
   If PlayRegAnimation Then
      sFrmName = sLinkStart & SatName
   Else
      sFrmName = sLinkStart
   End If
   frmCountry.Show
   If Animation Then
      GetAnimation sHurLink, sStartPos
   End If
End Sub

Private Sub GetAnimation(sHurLink As String, sLinkStart As String)
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim sPageName As String
   Dim sStartPos As String
   Dim sRegion As String
  
   On Error Resume Next
   If PlayRegAnimation Then
      GetWebpage "http://www.intellicast.com" & sHurLink & "&animate=true"
      If InStr(1, sLinkStart, "Satellite", vbTextCompare) <> 0 Then
         sStartPos = ">Infrared Satellite<"
         sFrmName = "Viewing " & sLinkStart & " Infrared Satellite"
      Else
         sStartPos = ">Current Radar<"
         sFrmName = "Viewing " & sLinkStart & " Radar"
      End If
      sStatusText = sFrmName
   Else
      GetWebpage "http://www.intellicast.com" & sHurLink & "?animate=true"
      sStartPos = sLinkStart
      sFrmName = sLinkStart
      sStatusText = sFrmName
   End If
   PlayAnimation = False
   'DoEvents
   iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
   If iIndexSt = 0 Then Exit Sub
   iIndex3 = InStr(iIndexSt, RichTextBox1.Text, "Content Container", vbTextCompare)
   iIndex = InStr(iIndex3, RichTextBox1.Text, "<img id=", vbTextCompare)
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "src=", vbTextCompare)
   iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, " alt", vbTextCompare)
   AnimationLink = Mid(RichTextBox1.Text, iIndexSt + 5, (iIndexEnd - 1) - (iIndexSt + 5))
   If InStr(1, sRegion, " ", vbTextCompare) <> 0 Then
      AnimationLink = Mid(sRegion, 1, InStr(1, sRegion, " ", vbTextCompare) - 2)
   End If
   frmAnimate.Show vbModal
End Sub

Private Sub getSatRegions()
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim cnt As Integer
   
   GetWebpage "http://www.intellicast.com/Global/Satellite/Infrared.aspx"
   
   'DoEvents
   cnt = 0
   iIndexEnd = InStr(1, RichTextBox1.Text, ">Infrared Satellite</div>", vbTextCompare)
   Do While cnt < 12
      iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "value=", vbTextCompare)
      iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
      mnuGbSat(cnt).Tag = Mid(RichTextBox1.Text, iIndex2 + 7, (iIndex - 1) - (iIndex2 + 7))
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "</option>", vbTextCompare)
      mnuGbSat(cnt).Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
      cnt = cnt + 1
      iIndexEnd = iIndexSt
   Loop
End Sub

Private Sub GetSatWaterVaper()
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim cnt As Integer
   
   GetWebpage "http://www.intellicast.com/National/Satellite/WaterVapor.aspx"
   
   'DoEvents
   cnt = 0
   iIndexEnd = InStr(1, RichTextBox1.Text, ">Satellite Maps<", vbTextCompare)
   Do While cnt < 38
      iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "value=", vbTextCompare)
      iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
      mnuWV(cnt).Tag = Mid(RichTextBox1.Text, iIndex2 + 7, (iIndex - 1) - (iIndex2 + 7))
      mnuRadCur(cnt).Tag = Mid(RichTextBox1.Text, iIndex2 + 7, (iIndex - 1) - (iIndex2 + 7))
      mnuCurLp(cnt).Tag = Mid(RichTextBox1.Text, iIndex2 + 7, (iIndex - 1) - (iIndex2 + 7))
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "</option>", vbTextCompare)
      mnuWV(cnt).Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
      mnuRadCur(cnt).Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
      mnuCurLp(cnt).Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
      cnt = cnt + 1
      iIndexEnd = iIndexSt
   Loop
End Sub

Private Sub GetCurrentSatellite()
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim cnt As Integer
   
   GetWebpage "http://www.intellicast.com/Global/Satellite/Current.aspx"
   
   'DoEvents
   cnt = 0
   iIndexEnd = InStr(1, RichTextBox1.Text, "Region:", vbTextCompare)
   Do While cnt < 46
      iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "value=", vbTextCompare)
      iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
      mnuCurSat(cnt).Tag = Mid(RichTextBox1.Text, iIndex2 + 7, (iIndex - 1) - (iIndex2 + 7))
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "</option>", vbTextCompare)
      mnuCurSat(cnt).Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
      cnt = cnt + 1
      iIndexEnd = iIndexSt
   Loop
End Sub

Private Sub getVisSatellite()
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim cnt As Integer
   
   GetWebpage "http://www.intellicast.com/National/Satellite/Visible.aspx"
   
   'DoEvents
   cnt = 0
   iIndexEnd = InStr(1, RichTextBox1.Text, "Region:", vbTextCompare)
   Do While cnt < 134
      iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "value=", vbTextCompare)
      iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
      mnuVisSat(cnt).Tag = Mid(RichTextBox1.Text, iIndex2 + 7, (iIndex - 1) - (iIndex2 + 7))
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "</option>", vbTextCompare)
      mnuVisSat(cnt).Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
      cnt = cnt + 1
      iIndexEnd = iIndexSt
   Loop
End Sub

Private Sub GetLatitude(sStringToFind As String, sCountryName As String)
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim iIndex As Long
   Dim sPageName As String
   Dim sStartPos As String
   Dim sLongitude As String
   Dim sLatitude As String
   Dim sCotryName As String
   Dim sLatName As String
   Dim sLonName As String
   
   On Error Resume Next
   If IsNumeric(sStringToFind) Then
      sPageName = "http://www.travelmath.com/zip-code/" & sStringToFind
   ElseIf sCountryName = "United States" Then
      sStringToFind = Replace(lblCity.Caption, " ", "+") '& "," & "+" & Replace(sCountryName, " ", "+")
      sPageName = "http://www.travelmath.com/city/" & sStringToFind
   Else
      If InStr(1, sStringToFind, "+", vbTextCompare) = 0 Then
         sStringToFind = Replace(sStringToFind, " ", "+") & "," & "+" & Replace(sCountryName, " ", "+")
      End If
      sPageName = "http://www.travelmath.com/city/" & sStringToFind
   End If
   
   GetWebpage sPageName
   sStartPos = "location0"
   'DoEvents
  
   iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
   If iIndex = 0 Then
      MsgBox "Unable to Show " & Replace(sStringToFind, "+", " ") & " GPS Location", vbInformation, "Weather Of The World Program"
      Exit Sub
   End If
   'City
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "<h4>", vbTextCompare)
   iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</h4>", vbTextCompare)
   sStatArea = Mid(RichTextBox1.Text, iIndexSt + 4, (iIndexEnd) - (iIndexSt + 4))
   'Region
   iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<strong>", vbTextCompare)
   iIndex = InStr(iIndexSt, RichTextBox1.Text, "<br />", vbTextCompare)
   sStatRegion = Replace(Mid(RichTextBox1.Text, iIndexSt + 8, (iIndex) - (iIndexSt + 8)), "</strong>", " ")
   If InStr(1, sStatRegion, "http:", vbTextCompare) <> 0 Then
      iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<strong>", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</strong>", vbTextCompare)
      sLatName = Mid(RichTextBox1.Text, iIndexSt + 8, (iIndexEnd) - (iIndexSt + 8))
      iIndexSt = InStr(iIndexEnd + 11, RichTextBox1.Text, ">", vbTextCompare)
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</a><br />", vbTextCompare)
      sStatRegion = sLatName & " " & Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1))
      'sStatCountry = "Kentronics Inc."
   End If
      'Country
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</strong>", vbTextCompare)
      sCotryName = Mid(RichTextBox1.Text, iIndexSt + 8, (iIndexEnd) - (iIndexSt + 8))
      iIndexSt = InStr(iIndexEnd + 11, RichTextBox1.Text, ">", vbTextCompare)
      iIndex = InStr(iIndexSt, RichTextBox1.Text, "</a><br />", vbTextCompare)
      sStatCountry = sCotryName & " " & Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1)), "</strong>", " ")
   'End If
   If IsNumeric(sStringToFind) = False Then
      'Latitude
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</strong>", vbTextCompare)
      sLatName = Mid(RichTextBox1.Text, iIndexSt + 8, (iIndexEnd) - (iIndexSt + 8))
      iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<br />", vbTextCompare)
      sStatState = sLatName & " " & Mid(RichTextBox1.Text, iIndexEnd + 9, (iIndexSt) - (iIndexEnd + 9))
      If InStr(1, sStatState, "http:", vbTextCompare) <> 0 Then
         iIndexSt = InStr(iIndexEnd + 11, RichTextBox1.Text, ">", vbTextCompare)
         iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</a><br />", vbTextCompare)
         sStatState = sLatName & " " & Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1))
      End If
      'Longitude
      iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<strong>", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</strong>", vbTextCompare)
      sLonName = Mid(RichTextBox1.Text, iIndexSt + 8, (iIndexEnd) - (iIndexSt + 8))
      iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "</strong>", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</div>", vbTextCompare)
      sStatCounty = sLonName & " " & Mid(RichTextBox1.Text, iIndexSt + 9, (iIndexEnd) - (iIndexSt + 9))
      If InStr(1, sStatCounty, "Latitude", vbTextCompare) <> 0 Then
         sStatCounty = "Kentronics Inc."
      Else
         sStatCounty = sLonName & " " & Mid(RichTextBox1.Text, iIndexSt + 9, (iIndexEnd) - (iIndexSt + 9))
      End If
      If InStr(1, sStatCounty, "http:", vbTextCompare) <> 0 Then
         iIndexSt = InStr(iIndexEnd + 11, RichTextBox1.Text, ">", vbTextCompare)
         iIndex = InStr(iIndexSt, RichTextBox1.Text, "</a><br />", vbTextCompare)
         sStatCounty = sLonName & " " & Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1))
      End If
   Else
      'County
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</strong>", vbTextCompare)
      sLatName = Mid(RichTextBox1.Text, iIndexSt + 8, (iIndexEnd) - (iIndexSt + 8))
      iIndexSt = InStr(iIndexEnd + 11, RichTextBox1.Text, ">", vbTextCompare)
      iIndex = InStr(iIndexSt, RichTextBox1.Text, "</a><br />", vbTextCompare)
      sStatState = sLatName & " " & Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1)), "</strong>", " ")
      'State
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "<strong>", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</strong>", vbTextCompare)
      sLonName = Mid(RichTextBox1.Text, iIndexSt + 8, (iIndexEnd) - (iIndexSt + 8))
      iIndexSt = InStr(iIndexEnd + 11, RichTextBox1.Text, ">", vbTextCompare)
      iIndex = InStr(iIndexSt, RichTextBox1.Text, "</a><br />", vbTextCompare)
      sStatCounty = sLonName & " " & Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1)), "</strong>", " ")
   End If
   'Latitude
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "Latitude:", vbTextCompare)
   iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</strong>", vbTextCompare)
   iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<br />", vbTextCompare)
   sLatitude = Mid(RichTextBox1.Text, iIndexEnd + 9, (iIndexSt) - (iIndexEnd + 9))
   'Longitude
   iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</strong>", vbTextCompare)
   iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "</div>", vbTextCompare)
   sLongitude = Mid(RichTextBox1.Text, iIndexEnd + 9, (iIndexSt) - (iIndexEnd + 9))
   
   AnimationLink = "http://www.mappingsupport.com/p/gmap4.php?ll=" & sLatitude & "," & sLongitude & "&z=11&t=m&icon=pgs"
   If IsNumeric(sStringToFind) Then
      sFrmName = lblCity.Caption & " GPS Location"
   Else
      sFrmName = Replace(sStringToFind, "+", " ") & " GPS Location"
   End If
   frmAnimate.Show vbModal
End Sub

Private Sub GetRadForcast()
   Dim iIndex, iIndex2, iIndex3 As Long
   Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim cnt As Integer
   
   GetWebpage "http://www.intellicast.com/National/Radar/Forecast.aspx"
   
   'DoEvents
   cnt = 0
   iIndexEnd = InStr(1, RichTextBox1.Text, "Region:", vbTextCompare)
   Do While cnt < 134
      iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "value=", vbTextCompare)
      iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
      mnuRadFor(cnt).Tag = Mid(RichTextBox1.Text, iIndex2 + 7, (iIndex - 1) - (iIndex2 + 7))
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "</option>", vbTextCompare)
      mnuRadFor(cnt).Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
      cnt = cnt + 1
      iIndexEnd = iIndexSt
   Loop
End Sub

Private Sub GetCityTag()
   'Dim oFndNode As Node
   'Set oFndNode = TreeFindNode(TView, sCityName, True, 1)
  sCityCode = QueryValue(HKEY_CURRENT_USER, CityCodeValue, "City_Tag_Name")
  sCityCode = StripTerminator(sCityCode)
End Sub

Private Function GetCityCode(sZip As String) As String
   Dim iIndexSt As Long
   Dim iIndexEnd As Long

   GetWebpage "http://www.intellicast.com/Local/Default.aspx?query=" & sZip
   'DoEvents
   'City Name
   iIndexSt = InStr(1, RichTextBox1.Text, "Primary Header FloatLeft", vbTextCompare)
   If iIndexSt = 0 Then
      GetCityCode = ""
      Exit Function
   End If
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "style=", vbTextCompare)
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, ">", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ",", vbTextCompare)
  sSelCityName = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1))
  'City Code
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "Current Conditions", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "location=", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
  GetCityCode = Mid(RichTextBox1.Text, iIndexSt + 9, (iIndexEnd - 1) - (iIndexSt + 9))
End Function

Public Sub stopAnimate()
  If Timer2.Enabled = True Then
    Timer2.Enabled = False
    Image1.Visible = False
  End If
End Sub

Private Sub GetCountryFacts(CtryUrl As String)
  Dim iIndex, iIndex2, iIndex3 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sLimit As Integer
  Dim sCountryStat As String
  Dim Sovereign As String
  Dim sNames As String
  Dim sMoreFacts As String
  Dim sFactsBody As String
  Dim sMoreInfo As String
  Dim sExtraBody As String
  
  sPageName = CtryUrl '"http://www.infoplease.com/ipa/A0107386.html"
  GetWebpage sPageName
  sStartPos = "Maptable end"
  'DoEvents
  txtCountryStat.Text = ""
  iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndex = 0 Then Exit Sub
  sNames = Space(500) & Mid(lblCity.Caption, InStr(1, lblCity, ",", vbTextCompare) + 1) & vbCrLf
  Do
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "<b pageno=""1"">", vbTextCompare)
    iIndexEnd = InStr(iIndexSt + 8, RichTextBox1.Text, "</b>", vbTextCompare)
    sCountryStat = Mid(RichTextBox1.Text, iIndexSt + 14, (iIndexEnd) - (iIndexSt + 14))
    
    If InStr(1, Mid(RichTextBox1.Text, iIndexEnd + 5, 20), "<a href=", vbTextCompare) <> 0 Then
      iIndexSt = InStr(iIndexEnd + 8, RichTextBox1.Text, ">", vbTextCompare)
      iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "</p><p ", vbTextCompare)
      Sovereign = Replace(Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex2) - (iIndexSt + 1)), "</a>", ""), "   ", "")
    Else
    If InStr(1, sCountryStat, "Land area:", vbTextCompare) <> 0 Then
      iIndex2 = InStr(iIndexEnd + 5, RichTextBox1.Text, "<b pageno=""1"">", vbTextCompare)
      Sovereign = Mid(RichTextBox1.Text, iIndexEnd + 5, (iIndex2) - (iIndexEnd + 5))
      iIndexSt = InStr(iIndex2 + 15, RichTextBox1.Text, "<b pageno=""1"">", vbTextCompare)
      sNames = sNames & sCountryStat & " " & Sovereign
      'total area
      iIndex3 = InStr(iIndexEnd + 8, RichTextBox1.Text, ">", vbTextCompare)
      iIndexEnd = InStr(iIndex3, RichTextBox1.Text, "</b> ", vbTextCompare)
      sCountryStat = Mid(RichTextBox1.Text, iIndex3 + 1, (iIndexEnd) - (iIndex3 + 1))
      iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "</p><p ", vbTextCompare)
      Sovereign = Mid(RichTextBox1.Text, iIndexEnd + 5, (iIndex2) - (iIndexEnd + 5))
    Else
      iIndex2 = InStr(iIndexEnd + 5, RichTextBox1.Text, "</p><p ", vbTextCompare)
      Sovereign = Mid(RichTextBox1.Text, iIndexEnd + 5, (iIndex2) - (iIndexEnd + 5))
    End If
    End If
    sNames = sNames & sCountryStat & " " & Sovereign & vbCrLf
    sNames = Replace(Replace(sNames, "</a>", ""), "eacute;", "e")
    sNames = Replace(sNames, "&pound;", Chr(163))
    sNames = Replace(sNames, "/a>", "")
    sNames = Replace(sNames, "&ndash;", Chr(45))
    sNames = Replace(sNames, ";", vbCrLf)
    iIndex = InStr(iIndex2 + 6, RichTextBox1.Text, "<b pageno=""1"">", vbTextCompare)
    If InStr(1, Mid(RichTextBox1.Text, iIndex2 + 5, 20), "align=", vbTextCompare) <> 0 Then
      sLimit = 1
    End If
    iIndexEnd = iIndex
  Loop Until sLimit = 1
  iIndexEnd = iIndex2
  'More Facts & Figures
  iIndexSt = InStr(iIndex2 + 8, RichTextBox1.Text, "?pageno=", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "</b>", vbTextCompare) '
  sMoreFacts = Mid(RichTextBox1.Text, iIndexSt + 10, (iIndex) - (iIndexSt + 10))
  sMoreFacts = Replace(sMoreFacts, ">", "")
  Do
    iIndex3 = InStr(iIndex, RichTextBox1.Text, "<h1", vbTextCompare)
    iIndexSt = InStr(iIndex3, RichTextBox1.Text, "class=""level3"">", vbTextCompare)
    If iIndexSt = 0 Then
      iIndex3 = InStr(iIndex, RichTextBox1.Text, "<p>", vbTextCompare)
      iIndexSt = InStr(iIndex3, RichTextBox1.Text, "</p>", vbTextCompare)
      sExtraBody = Mid(RichTextBox1.Text, iIndex3 + 3, (iIndexSt) - (iIndex3 + 3))
      sMoreInfo = sMoreInfo & sExtraBody
      sMoreInfo = Replace(sMoreInfo, "  ", " ")
      GoTo endLoop
    End If
    iIndex = InStr(iIndexSt, RichTextBox1.Text, "</h1><p>", vbTextCompare)
    sFactsBody = Mid(RichTextBox1.Text, iIndexSt + 15, (iIndex) - (iIndexSt + 15))
    sMoreInfo = sMoreInfo & vbCrLf & sFactsBody & vbCrLf
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "</p>", vbTextCompare)
    sFactsBody = Mid(RichTextBox1.Text, iIndex + 8, (iIndexSt) - (iIndex + 8))
    sFactsBody = Replace(Replace(sFactsBody, "</a>", ""), "</span>", "")
    If InStr(1, sFactsBody, "<a href=", vbTextCompare) <> 0 Then
      Do
        sFactsBody = Mid(sFactsBody, 1, InStr(1, sFactsBody, "<a href=", vbTextCompare) - 1) & Mid(sFactsBody, InStr(1, sFactsBody, ">", vbTextCompare) + 1)
      Loop Until InStr(1, sFactsBody, "<a href=", vbTextCompare) = 0
    End If
    sFactsBody = Replace(sFactsBody, "<span class=""small"" pageno=""1"">", "")
    If InStr(1, sFactsBody, "<span class=", vbTextCompare) <> 0 Then
      Do
        sFactsBody = Mid(sFactsBody, 1, InStr(1, sFactsBody, "<span class=", vbTextCompare) - 1) & Mid(sFactsBody, InStr(1, sFactsBody, ">", vbTextCompare) + 1)
      Loop Until InStr(1, sFactsBody, "<span class=", vbTextCompare) = 0
    End If
    sMoreInfo = sMoreInfo & sFactsBody & vbCrLf
    sLimit = 0
endLoop:
    If InStr(1, Mid(RichTextBox1.Text, iIndexSt + 5, 20), "pageno=", vbTextCompare) <> 0 Or InStr(1, Mid(RichTextBox1.Text, iIndexSt + 5, 20), "align=", vbTextCompare) <> 0 Then
      sLimit = 1
    End If
    iIndex = iIndexSt
  Loop Until sLimit = 1
  txtCountryStat.Text = Replace(Replace(sNames, ";", vbCrLf), "   ", "") & vbCrLf & Space(50) & sMoreFacts & vbCrLf & sMoreInfo
  txtCountryStat.Text = Replace(txtCountryStat.Text, "</span>", "")
  txtCountryStat.Text = Replace(txtCountryStat.Text, "<span class=""small"" pageno=""1"">", "")
  txtCountryStat.Text = Replace(txtCountryStat.Text, "<b pageno=""1"">", "")
  txtCountryStat.Text = Replace(txtCountryStat.Text, "<i pageno=""1"">", "")
  txtCountryStat.Text = Replace(txtCountryStat.Text, "pageno=""1"">1</sup>", "")
  txtCountryStat.Text = Replace(txtCountryStat.Text, "</i>", "")
  txtCountryStat.Text = txtCountryStat.Text & vbCrLf & GetMoreFacts(sPageName)
  If InStr(1, txtCountryStat.Text, "<i pageno=", vbTextCompare) <> 0 Then
    txtCountryStat.Text = Mid(txtCountryStat.Text, 1, InStrRev(txtCountryStat.Text, "<i pageno=") - 2)
  End If
  frmAlert.txtAlert.Visible = True
  frmAlert.txtAlert.Text = txtCountryStat.Text
  frmAlert.Caption = Mid(lblCity.Caption, InStr(1, lblCity, ",", vbTextCompare) + 1) & " Facts & Figures"
  frmAlert.txtAlert.FontSize = 10
  frmAlert.Show vbModal
End Sub

Private Function GetMoreFacts(sUrlpage As String) As String
  Dim iIndex, iIndex2, iIndex3 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sFactsBody As String
  Dim sHeading As String
  Dim cityArray() As String
  Dim cnt As Integer
  Dim sLimit As Integer
  
  On Error GoTo errorHandler
  
  sPageName = sUrlpage
  GetWebpage sPageName
  sStartPos = "Main Page<"
  'DoEvents
  txtCountryStat.Text = ""
  iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndex = 0 Then Exit Function
  Do
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "<a href=", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
    ReDim Preserve cityArray(cnt)
    cityArray(cnt) = Mid(RichTextBox1.Text, iIndexSt + 9, (iIndexEnd - 1) - (iIndexSt + 9))
    iIndexSt = InStr(iIndexEnd + 2, RichTextBox1.Text, "</li>", vbTextCompare)
    iIndex = iIndexSt
    cnt = cnt + 1
  Loop Until InStr(1, Mid(RichTextBox1.Text, iIndexSt, 35), "</table>", vbTextCompare) <> 0
  For sLimit = 0 To UBound(cityArray, 1)
    GetWebpage sPageName & cityArray(sLimit)
    sStartPos = "pagebreak"
    iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
    If iIndex = 0 Then Exit Function
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "<b>", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</b>", vbTextCompare)
    
    sHeading = Mid(RichTextBox1.Text, iIndexSt + 3, (iIndexEnd) - (iIndexSt + 3))
    iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<p>", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "<table align=", vbTextCompare)
    
    sFactsBody = Mid(RichTextBox1.Text, iIndexSt + 3, (iIndexEnd) - (iIndexSt + 3))
    If InStr(1, sFactsBody, "<i", vbTextCompare) <> 0 Then
      sFactsBody = Mid(sFactsBody, 1, InStrRev(sFactsBody, "<i") - 1)
    End If
    sFactsBody = Replace(sFactsBody, "</p><p>", vbCrLf)
    sFactsBody = Replace(sFactsBody, "  ", " ")
    If InStr(1, sFactsBody, "<i pageno=", vbTextCompare) <> 0 Then
      sFactsBody = Mid(sFactsBody, 1, InStrRev(sFactsBody, "<i pageno=") - 1)
    End If
    If sLimit <> 0 Then
      txtCountryStat.Text = txtCountryStat.Text & vbCrLf & vbCrLf & sHeading & vbCrLf & vbCrLf & sFactsBody
    Else
      txtCountryStat.Text = txtCountryStat.Text & vbCrLf & sHeading & vbCrLf & vbCrLf & sFactsBody
    End If
  Next
  GetMoreFacts = Replace(txtCountryStat.Text, "</p>", "")
  Exit Function
errorHandler:
  MsgBox Mid(RichTextBox1.Text, iIndexSt, iIndexEnd - (iIndexSt))
End Function

Private Sub GetPopulation()
  Dim iIndex, iIndex2 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sCnt As Integer
  Dim X As Integer
  Dim sCountryName As String
  
  sPageName = "http://www.infoplease.com/ipa/A0004379.html"
  GetWebpage sPageName
  sStartPos = "BodyText"
  frmAlert.lstPopulation.ColumnHeaders(1).Width = 3800
  frmAlert.lstPopulation.ColumnHeaders(2).Width = 2300
  frmAlert.lstPopulation.ColumnHeaders(3).Width = 2300
  frmAlert.lstPopulation.ColumnHeaders.Remove 4
  frmAlert.lstPopulation.ColumnHeaders.Item(1).Text = "Country"
  frmAlert.lstPopulation.ColumnHeaders.Item(2).Text = "Population"
  frmAlert.lstPopulation.ColumnHeaders.Item(3).Text = "Area (in sq mi)"
 
  iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndex = 0 Then Exit Sub
  ProgressBar1.Scrolling = ccScrollingSmooth
  ProgressBar1.Max = 199
  ProgressBar1.Visible = True
  frmAlert.MousePointer = 11
  'sCnt = 1
  Do
   If X Mod 3 = 0 Then
      'Country
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "align=", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
      iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "</", vbTextCompare)
      sCountryName = Trim(Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex2) - (iIndexEnd + 1)))
      sCountryName = Replace(Replace(sCountryName, "&eacute;", Chr(233)), "&atilde;", Chr(226))
      sfndResult = FindStringinListControl(frmAlert.cmbcntyName, Trim(sCountryName))
      If sfndResult <> -1 Then
        frmAlert.lstPopulation.ListItems.Add , , Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex2) - (iIndexEnd + 1)), , sfndResult + 1
      Else
        frmAlert.lstPopulation.ListItems.Add , , Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex2) - (iIndexEnd + 1)) ', , iCnt
      End If
      sCnt = sCnt + 1
    Else
      iIndexSt = InStr(iIndex2, RichTextBox1.Text, "align=", vbTextCompare)
      If iIndexSt = 0 Then Exit Do
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
      iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "</", vbTextCompare)
      sCityName = Replace(Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex2) - (iIndexEnd + 1)), "&aacute;", Chr(225))
      sCityName = Replace(Replace(sCityName, "&eacute;", Chr(233)), "&atilde;", Chr(226))
      frmAlert.lstPopulation.ListItems(sCnt).ListSubItems.Add , , sCityName
      iIndex = iIndex2
    End If
    X = X + 1
    ProgressBar1.Value = sCnt
  Loop Until InStr(1, Mid(RichTextBox1.Text, iIndex2, 50), "</table>", vbTextCompare) <> 0
    
  frmAlert.lstWeatherAlert.Visible = False
  frmAlert.txtAlert.Visible = False
  frmAlert.lstPopulation.Visible = True
  frmAlert.lstPopulation.Height = frmAlert.lstPopulation.Height - 100
  frmAlert.Caption = "Area and Population of Countries"
  frmAlert.MousePointer = 0
  ProgressBar1.Max = 1
  ProgressBar1.Visible = False
  frmAlert.Show vbModal
End Sub

Private Sub GetPopDensity()
  Dim iIndex, iIndex2 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sCnt As Integer
  Dim X As Integer
  Dim sCountryName As String
  
  sPageName = "http://www.infoplease.com/ipa/A0934666.html"
  GetWebpage sPageName
  sStartPos = "BodyText"
  'DoEvents
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "><td", vbTextCompare)
  If iIndex2 = 0 Then Exit Sub
  ProgressBar1.Scrolling = ccScrollingStandard
  ProgressBar1.Max = 229
  ProgressBar1.Visible = True
  sCnt = 1
  Do
    'Country
    iIndex = InStr(iIndex2, RichTextBox1.Text, "valign=", vbTextCompare)
    If InStr(1, Mid(RichTextBox1.Text, iIndex, 20), "><a", vbTextCompare) <> 0 Then
      iIndex2 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
      iIndexSt = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</a", vbTextCompare)
    Else
      iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</td", vbTextCompare)
    End If
    sCountryName = Trim(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1)))
    sCountryName = Replace(Replace(sCountryName, "&eacute;", Chr(233)), "&atilde;", Chr(226))
    sfndResult = FindStringinListControl(frmAlert.cmbcntyName, Trim(sCountryName))
    If sfndResult <> -1 Then
      frmAlert.lstPopulation.ListItems.Add , , sCountryName, , sfndResult + 1
    Else
      frmAlert.lstPopulation.ListItems.Add , , sCountryName ', , iCnt
    End If
    
    'frmAlert.lstPopulation.ListItems.Add , , sCountryName
    'City
    For X = 0 To 2
      iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "valign=", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
      iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "</td", vbTextCompare)
      sCityName = Replace(Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex2) - (iIndexEnd + 1)), "&aacute;", Chr(225))
      sCityName = Replace(Replace(sCityName, "&eacute;", Chr(233)), "&atilde;", Chr(226))
      frmAlert.lstPopulation.ListItems(sCnt).ListSubItems.Add , , sCityName
      iIndexEnd = iIndex2
    Next
    DoEvents
    ProgressBar1.Value = sCnt
    sCnt = sCnt + 1
  Loop Until InStr(1, Mid(RichTextBox1.Text, iIndex2, 25), "></table", vbTextCompare) <> 0
  frmAlert.lstWeatherAlert.Visible = False
  frmAlert.txtAlert.Visible = False
  frmAlert.lstPopulation.Height = frmAlert.lstPopulation.Height - 100
  frmAlert.lstPopulation.Visible = True
  frmAlert.lstPopulation.ColumnHeaders(1).Width = 3500
  frmAlert.lstPopulation.ColumnHeaders(2).Width = 1600
  frmAlert.lstPopulation.ColumnHeaders(3).Width = 1750
  frmAlert.lstPopulation.ColumnHeaders(4).Width = 1650
  frmAlert.lstPopulation.ColumnHeaders(2).Text = "Population"
  frmAlert.lstPopulation.ColumnHeaders(3).Text = "Land Area Sq/Mi"
  frmAlert.lstPopulation.ColumnHeaders(4).Text = "Density Sq/Mi"
  frmAlert.lstWeatherAlert.HideColumnHeaders = False
  frmAlert.Caption = "Population Density per Square Mile of Countries"
  ProgressBar1.Max = 1
  ProgressBar1.Visible = False
  frmAlert.Show vbModal
End Sub

Private Sub Get50MostPop(WrlUrl As String, year As String)
  Dim iIndex, iIndex2 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sCnt As Integer
  Dim X As Integer
  Dim sCountryName As String
  
  sPageName = WrlUrl
  GetWebpage sPageName
  sStartPos = "BodyText"
  'DoEvents
 
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then Exit Sub
  For X = 1 To 3
    iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "valign=", vbTextCompare)
    iIndexSt = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
    iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "</th>", vbTextCompare)
    frmAlert.lstPopulation.ColumnHeaders(X).Text = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex2) - (iIndexSt + 1))
  Next
  frmAlert.lstPopulation.ColumnHeaders(1).Width = 1400
  frmAlert.lstPopulation.ColumnHeaders(2).Width = 4100
  frmAlert.lstPopulation.ColumnHeaders(3).Width = 3000
  frmAlert.lstPopulation.ColumnHeaders.Remove 4
  
  ProgressBar1.Scrolling = ccScrollingStandard
  ProgressBar1.Max = 51
  ProgressBar1.Visible = True
  sCnt = 1
  If year = 2008 Then
  ProgressBar1.Max = 50
    GoTo s2008
  End If
  'Country
  iIndex = InStr(iIndex2, RichTextBox1.Text, "valign=", vbTextCompare)
  If InStr(1, Mid(RichTextBox1.Text, iIndex, 20), "><a", vbTextCompare) <> 0 Then
    iIndex2 = InStr(iIndex, RichTextBox1.Text, "href=", vbTextCompare)
    iIndexSt = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</a", vbTextCompare)
  Else
    iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</td", vbTextCompare)
  End If
  sCountryName = Trim(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1)))
  sCountryName = Replace(Replace(sCountryName, "&eacute;", Chr(233)), "&atilde;", Chr(226))
  frmAlert.lstPopulation.ListItems.Add , , Replace(sCountryName, "&nbsp;", "-")
  If sCountryName = "&nbsp;" Then
    iIndexSt = InStr(iIndexEnd + 6, RichTextBox1.Text, ">", vbTextCompare)
    iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "</td", vbTextCompare)
    frmAlert.lstPopulation.ListItems(sCnt).ListSubItems.Add , , Trim(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex2) - (iIndexSt + 1)))
    iIndexSt = InStr(iIndex2 + 6, RichTextBox1.Text, ">", vbTextCompare)
    iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "</td", vbTextCompare)
    frmAlert.lstPopulation.ListItems(sCnt).ListSubItems.Add , , Trim(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex2) - (iIndexSt + 1)))
    sCnt = sCnt + 1
  End If
s2008:
  'City
  Do
    'Rank
    iIndex = InStr(iIndex2, RichTextBox1.Text, "valign=", vbTextCompare)
    iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "</td", vbTextCompare)
    frmAlert.lstPopulation.ListItems.Add , , Trim(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex2) - (iIndexSt + 1)))
    'Country
    iIndex = InStr(iIndex2, RichTextBox1.Text, "href=", vbTextCompare)
    iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</a>", vbTextCompare)
    frmAlert.lstPopulation.ListItems(sCnt).ListSubItems.Add , , Replace(Trim(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1))), "<br />&nbsp;", " ")
    'Population
    iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "valign=", vbTextCompare)
    If iIndexSt = 0 Then
      Exit Do
    End If
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
    iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "</td", vbTextCompare)
    frmAlert.lstPopulation.ListItems(sCnt).ListSubItems.Add , , Trim(Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex2) - (iIndexEnd + 1)))
    
    DoEvents
    ProgressBar1.Value = sCnt
    sCnt = sCnt + 1
  Loop Until InStr(1, Mid(RichTextBox1.Text, iIndex2, 35), "</table", vbTextCompare) <> 0
    
  frmAlert.lstWeatherAlert.Visible = False
  frmAlert.txtAlert.Visible = False
  frmAlert.lstPopulation.Height = frmAlert.lstPopulation.Height - 100
  frmAlert.lstPopulation.Visible = True
  frmAlert.lstWeatherAlert.HideColumnHeaders = False
  frmAlert.Caption = "World's 50 Most Populous Countries: " & year
  ProgressBar1.Max = 1
  ProgressBar1.Visible = False
  frmAlert.Show vbModal
End Sub

Private Sub GetEconomicStats(WrlUrl As String, year As String)
  Dim iIndex, iIndex2 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sCnt As Integer
  Dim X As Integer
  Dim sCountryName As String
  
  sPageName = WrlUrl
  GetWebpage sPageName
  sStartPos = "BodyText"
  'DoEvents
 
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then Exit Sub
  frmAlert.lstPopulation.ColumnHeaders(1).Width = 3300
  frmAlert.lstPopulation.ColumnHeaders(2).Width = 1700
  frmAlert.lstPopulation.ColumnHeaders(3).Width = 1200
  frmAlert.lstPopulation.ColumnHeaders(4).Width = 1100
  frmAlert.lstPopulation.ColumnHeaders.Add , , , 1100, 0
  For X = 1 To 5
    If X = 1 Then
      iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "valign=", vbTextCompare)
      
      If year = "2005" Then
        iIndexSt = InStr(iIndex2, RichTextBox1.Text, "><b", vbTextCompare)
        iIndex2 = InStr(iIndexSt + 2, RichTextBox1.Text, ">", vbTextCompare)
        iIndexSt = InStr(iIndex2, RichTextBox1.Text, "</b", vbTextCompare)
        frmAlert.lstPopulation.ColumnHeaders(X).Text = Mid(RichTextBox1.Text, iIndex2 + 1, (iIndexSt) - (iIndex2 + 1))
      Else
        iIndexSt = InStr(iIndex2, RichTextBox1.Text, "<b>", vbTextCompare)
        iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "</b>", vbTextCompare)
        frmAlert.lstPopulation.ColumnHeaders(X).Text = Mid(RichTextBox1.Text, iIndexSt + 3, (iIndex2) - (iIndexSt + 3))
      End If
    Else
      iIndex = InStr(iIndex2, RichTextBox1.Text, "valign=", vbTextCompare)
      iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
      If year = "2005" Then
        iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "</th", vbTextCompare)
      Else
        iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "</th>", vbTextCompare)
      End If
      'frmAlert.lstPopulation.ColumnHeaders(X).Alignment = lvwColumnCenter
      frmAlert.lstPopulation.ColumnHeaders(X).Text = Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex2) - (iIndexSt + 1)), "<br />", " ")
    End If
  Next
  
  ProgressBar1.Scrolling = ccScrollingStandard
  ProgressBar1.Max = 195
  ProgressBar1.Visible = True
  sCnt = 1
  
  Do
    'Country
    iIndex = InStr(iIndex2, RichTextBox1.Text, "href=", vbTextCompare)
    iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</a", vbTextCompare)
    sCountryName = Trim(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1)))
    sCountryName = Replace(Replace(sCountryName, "&eacute;", Chr(233)), "&atilde;", Chr(226))
    frmAlert.lstPopulation.ListItems.Add , , sCountryName
  
    For X = 0 To 3
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, "valign=", vbTextCompare)
      iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
      If InStr(1, Mid(RichTextBox1.Text, iIndexSt, 12), "sup", vbTextCompare) <> 0 Then
        iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "<", vbTextCompare)
        sCityName = Trim(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex2) - (iIndexSt + 1))) & Chr(185)
        iIndex = InStr(iIndex2, RichTextBox1.Text, "</td", vbTextCompare)
        iIndexEnd = iIndex
      Else
        iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "</td", vbTextCompare)
        sCityName = Trim(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex2) - (iIndexSt + 1)))
        iIndexEnd = iIndex2
      End If
      frmAlert.lstPopulation.ListItems(sCnt).ListSubItems.Add , , Replace(Replace(sCityName, "&ndash;", "-"), "&#8211;", "-")
    Next
    
    DoEvents
    ProgressBar1.Value = sCnt
    sCnt = sCnt + 1
  Loop Until sCountryName = "Zimbabwe"
  frmAlert.lstWeatherAlert.Visible = False
  frmAlert.txtAlert.Visible = False
  frmAlert.lstPopulation.Height = frmAlert.lstPopulation.Height - 100
  frmAlert.lstPopulation.Visible = True
  frmAlert.lstWeatherAlert.HideColumnHeaders = False
  frmAlert.Caption = "Economic Statistics by Country: " & year
  ProgressBar1.Max = 1
  ProgressBar1.Visible = False
  frmAlert.Show vbModal
End Sub

Private Sub GetCommNation(WrlUrl As String)
  Dim iIndex, iIndex2 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sCnt As Integer
  Dim X As Integer
  Dim sCountryName As String
  
  sPageName = WrlUrl
  GetWebpage sPageName
  sStartPos = "<ul class="
  'DoEvents
 
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then Exit Sub
  frmAlert.lstPopulation.ColumnHeaders.Remove 3
  frmAlert.lstPopulation.ColumnHeaders.Remove 2
  frmAlert.lstPopulation.ColumnHeaders(2).Text = "Country"
  frmAlert.lstPopulation.GridLines = False
  frmAlert.lstPopulation.ColumnHeaders(1).Width = 4350
  frmAlert.lstPopulation.ColumnHeaders(2).Width = 4350
  ProgressBar1.Scrolling = ccScrollingSmooth
  ProgressBar1.Max = 51
  ProgressBar1.Visible = True
  sCnt = 1
  
  Do
    'Country
    iIndex = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
    If iIndex = 0 Then Exit Do
    iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</a>", vbTextCompare)
    sCountryName = Trim(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1)))
    sCountryName = Replace(Replace(sCountryName, "   ", " "), "  ", " ")
    sfndResult = FindStringinListControl(frmAlert.cmbcntyName, Trim(sCountryName))
    If sfndResult <> -1 Then
      frmAlert.lstPopulation.ListItems.Add , , sCountryName, , sfndResult + 1
    Else
      frmAlert.lstPopulation.ListItems.Add , , sCountryName ', , iCnt
    End If
    If sCountryName = "Zambia" Then Exit Do
      X = X + 1
    iIndex = InStr(iIndexEnd, RichTextBox1.Text, "href=", vbTextCompare)
    iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</a>", vbTextCompare)
    sCountryName = Trim(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1)))
    sCountryName = Replace(Replace(sCountryName, "    ", " "), "  ", "")
    sfndResult = FindStringinListControl(frmAlert.cmbcntyName, Trim(sCountryName))
    If sfndResult <> -1 Then
      frmAlert.lstPopulation.ListItems(X).ListSubItems.Add , , sCountryName, sfndResult + 1
    Else
      frmAlert.lstPopulation.ListItems(X).ListSubItems.Add , , sCountryName
    End If
    DoEvents
    iIndexSt = iIndexEnd
    ProgressBar1.Value = sCnt
    sCnt = sCnt + 1
  Loop Until sCountryName = "Zambia"
  frmAlert.lstWeatherAlert.Visible = False
  frmAlert.txtAlert.Visible = False
  frmAlert.lstPopulation.Height = frmAlert.lstPopulation.Height - 100
  frmAlert.lstPopulation.Visible = True
  frmAlert.lstWeatherAlert.HideColumnHeaders = False
  frmAlert.Caption = "Members of the Commonwealth of Nations"
  ProgressBar1.Max = 1
  ProgressBar1.Visible = False
  frmAlert.Show vbModal
End Sub

Private Sub LoadCountryFlags()
  Dim nFileNum As Integer
  Dim sString As String
  Dim cntCount As Integer

  nFileNum = FreeFile
  Open App.Path + "\Country Flag Names.Dat" For Binary Access Read As #nFileNum
  Do While Not EOF(nFileNum)
    'read the length of the string
    Get #nFileNum, , nLen
    'initialize the string with the correct number of spaces
    sString = Space$(nLen)
    Get #nFileNum, , sString
    sString = DecryptText((sString), sPassword, True)
    'Line Input #nFileNum, sString
    If Len(Trim$(sString)) > 1 Then
      cmbCode.AddItem Trim(sString)
      ReDim Preserve CountriesArray(cntCount)
      CountriesArray(cntCount) = Replace(Replace(Trim(sString), " & ", " and "), "Saint ", "St. ")
      cntCount = cntCount + 1
    End If
  Loop
  Close #nFileNum
  cmbCode.ListIndex = 0
End Sub

Public Sub UpdateMenuValues(menuIndex As Integer, MenuDel As Boolean)
  Dim KeyCollection As Collection
  Dim Object As Variant
  Dim KeyName As String
  Dim cnt As Integer
  Dim oldKeyCount As Integer
  Dim I As Integer
  
  Set KeyCollection = EnumRegistryValues(HKEY_CURRENT_USER, FilelistKey)
  oldKeyCount = KeyCollection.Count
  If KeyCollection.Count < 1 Then
    mnuRemoveBookMark.Enabled = False
    Exit Sub
  Else
    mnuRemoveBookMark.Enabled = True
  End If
  If Not MenuDel Then
    If KeyCollection.Count <> 0 Then
      For Each Object In KeyCollection
        cnt = Mid(Object(0), InStrRev(Object(0), "-") + 1)
        I = I + 1
        Select Case Mid(Object(0), 1, InStr(1, Object(0), "-") - 1)
          Case "City_Name"
            mnuFavorite(cnt).Caption = Object(1)
            mnuRemove(cnt).Caption = Object(1)
          Case "City_Tag_Name"
            mnuFavorite(cnt).Tag = Object(1)
            mnuRemove(cnt).Tag = Object(1)
        End Select
        If Len(mnuFavorite(cnt).Caption) <> 0 Then
          mnuFavorite(cnt).Visible = True
          mnuRemove(cnt).Visible = True
          mnuFavorite(cnt).Enabled = True
        Else
          mnuFavorite(cnt).Visible = False
          mnuRemove(cnt).Visible = False
        End If
        If I >= 10 Then
          mnuAdd.Enabled = False
        End If
      Next
    End If
  Else
    If KeyCollection.Count <> 0 Then
      For Each Object In KeyCollection
        cnt = Mid(Object(0), InStrRev(Object(0), "-") + 1)
        KeyName = Object(0)
        If menuIndex = cnt Then
          DeleteRegisterValue HKEY_CURRENT_USER, FilelistKey, KeyName
          mnuAdd.Enabled = True
          oldKeyCount = oldKeyCount - 2
          If oldKeyCount >= 2 Then
            mnuFavorite(cnt).Caption = ""
            mnuRemove(cnt).Caption = ""
            mnuFavorite(cnt).Enabled = False
            mnuFavorite(cnt).Visible = False
            mnuRemove(cnt).Visible = False
          ElseIf oldKeyCount <= -2 Then
            mnuRemoveBookMark.Enabled = False
            mnuFavorite(cnt).Enabled = False
          Else
            mnuFavorite(cnt).Caption = ""
            mnuRemove(cnt).Caption = ""
          End If
        End If
      Next
    End If
  End If
  Set KeyCollection = Nothing
End Sub

Public Function DeleteRegisterValue(lPredefinedKey As Long, sKeyName As String, sValueName As String) As Long
  Dim lRetVal As Long         'result of the API functions
  Dim hKey As Long         'handle of opened key
  Dim vValue As Variant      'setting of queried value
  lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
  lRetVal = RegDeleteValue(hKey, sValueName)
  RegCloseKey (hKey)
End Function

Public Function StripTerminator(ByVal strString As String) As String
  ' strip off trailing NULL's from API calls
  Dim intZeroPos      As Integer

  intZeroPos = InStr(strString, vbNullChar)
    
  If intZeroPos > 1 Then
    StripTerminator = Trim$(Left$(strString, intZeroPos - 1))
  ElseIf intZeroPos = 1 Then
    StripTerminator = vbNullString
  Else
    StripTerminator = strString
  End If
End Function

Private Sub reMoveIcons()
  Set fso = CreateObject("Scripting.FileSystemObject")
  fso.DeleteFile App.Path & "\Icons\*.*", True
  Set fso = Nothing
End Sub

Private Sub GetCountryTimeDate(sStringToFind As String, sCountryName As String)
  Dim iIndexEnd As Long
   Dim iIndexSt As Long
   Dim iIndex As Long
   Dim sPageName As String
   Dim sStartPos As String
      
   'On Error Resume Next
   If IsNumeric(sStringToFind) Then
      sPageName = "http://www.travelmath.com/zip-code/" & sStringToFind
   ElseIf sCountryName = "United States" Then
      sStringToFind = Replace(lblCity.Caption, " ", "+")
      sPageName = "http://www.travelmath.com/city/" & sStringToFind
   Else
      If InStr(1, sStringToFind, "+", vbTextCompare) = 0 Then
         sStringToFind = Replace(sStringToFind, " ", "+") & "," & "+" & Replace(sCountryName, " ", "+")
      End If
      sPageName = "http://www.travelmath.com/city/" & sStringToFind
   End If
   
   GetWebpage sPageName
   sStartPos = "Time zone:"
   'DoEvents
  
   iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
   If iIndex = 0 Then
      MsgBox "Unable to Show " & Replace(sStringToFind, "+", " ") & " Time & Date", vbInformation, "Weather Of The World Program"
      Exit Sub
   End If
   iIndexSt = InStr(iIndex, RichTextBox1.Text, "UTC/GMT", vbTextCompare)
   iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "<p>", vbTextCompare)
   iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</p>", vbTextCompare)
   sStatArea = Mid(RichTextBox1.Text, iIndexEnd + 3, (iIndex) - (iIndexEnd + 3))
   MsgBox Mid(sStatArea, 1, InStr(1, sStatArea, "is ", vbTextCompare) + 2) & vbCrLf & _
          Mid(sStatArea, InStrRev(sStatArea, ":") - 2), vbInformation, "Weather Of The World Program"
End Sub

Private Sub GetNatHoliday()
  Dim iIndex, iIndex2 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim Limits As Integer
  Dim sPageName As String
  Dim sStartPos As String
  Dim iCnt As Integer
  Dim X As Integer
  Dim sCountryName As String
  
  sPageName = "http://www.infoplease.com/ipa/A0907876.html"
  GetWebpage sPageName
  sStartPos = "BodyText"
  'DoEvents
 
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then Exit Sub
  frmAlert.lstPopulation.ColumnHeaders(1).Width = 3000
  frmAlert.lstPopulation.ColumnHeaders(2).Width = 2300
  frmAlert.lstPopulation.ColumnHeaders(3).Width = 3250
  frmAlert.lstPopulation.ColumnHeaders.Remove 4
  frmAlert.lstPopulation.ColumnHeaders.Item(1).Text = "Country"
  frmAlert.lstPopulation.ColumnHeaders.Item(2).Text = "Date"
  frmAlert.lstPopulation.ColumnHeaders.Item(3).Text = "Holiday"
  ProgressBar1.Scrolling = ccScrollingStandard
  ProgressBar1.Max = 657
  ProgressBar1.Visible = True
  
  Do
    If X Mod 3 = 0 Then
      iIndex = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
      iIndex2 = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
      iIndexSt = InStr(iIndex2, RichTextBox1.Text, "</a", vbTextCompare)
      sCountryName = Mid(RichTextBox1.Text, iIndex2 + 1, (iIndexSt) - (iIndex2 + 1))
      sfndResult = FindStringinListControl(frmAlert.cmbcntyName, Trim(sCountryName))
      If sfndResult <> -1 Then
        frmAlert.lstPopulation.ListItems.Add , , sCountryName, , sfndResult + 1
      Else
        frmAlert.lstPopulation.ListItems.Add , , sCountryName
      End If
      iCnt = iCnt + 1
    Else
      iIndex = InStr(iIndexSt, RichTextBox1.Text, """top"">", vbTextCompare)
      iIndex2 = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
      iIndexSt = InStr(iIndex2, RichTextBox1.Text, "</td>", vbTextCompare)
      sCountryName = Replace((Replace(Mid(RichTextBox1.Text, iIndex2 + 1, (iIndexSt) - (iIndex2 + 1)), Chr(10), " ")), "   ", "")
      If InStr(1, sCountryName, "<sup", vbTextCompare) <> 0 Then
        sCountryName = Mid(sCountryName, 1, InStr(1, sCountryName, "<sup", vbTextCompare) - 1)
      End If
      frmAlert.lstPopulation.ListItems(iCnt).ListSubItems.Add , , sCountryName, , sCountryName
    End If
    
    If InStr(1, Mid(RichTextBox1.Text, iIndexSt, 40), "</table>", vbTextCompare) <> 0 Then
      Limits = 1
    End If
    X = X + 1
    ProgressBar1.Value = X
  Loop Until Limits = 1
  
  frmAlert.lstWeatherAlert.Visible = False
  frmAlert.txtAlert.Visible = False
  frmAlert.lstPopulation.Height = frmAlert.lstPopulation.Height - 100
  frmAlert.lstPopulation.Visible = True
  frmAlert.Caption = "National Holidays Around the World"
  ProgressBar1.Max = 1
  ProgressBar1.Visible = False
  frmAlert.Show vbModal
End Sub

Private Sub GetWorldCapital(sUrl As String)
  Dim iIndex, iIndex2 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim Limits As Integer
  Dim sPageName As String
  Dim sStartPos As String
  Dim iCnt As Integer
  Dim X As Integer
  Dim sCountryName As String
  
  sPageName = sUrl
  GetWebpage sPageName
  sStartPos = "BodyText"
  'DoEvents
 
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then
    Exit Sub
  End If
  frmAlert.lstPopulation.ColumnHeaders(1).Width = 3500
  frmAlert.lstPopulation.ColumnHeaders(2).Width = 5000
  frmAlert.lstPopulation.ColumnHeaders.Remove 4
  frmAlert.lstPopulation.ColumnHeaders.Remove 3
  frmAlert.lstPopulation.ColumnHeaders.Item(1).Text = "Country"
  frmAlert.lstPopulation.ColumnHeaders.Item(2).Text = "City, Population"
  ProgressBar1.Scrolling = ccScrollingStandard
  ProgressBar1.Max = 394
  ProgressBar1.Visible = True
  frmWeatherMain.MousePointer = 11
  Do
    If X Mod 2 = 0 Then
      iIndex = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
      iIndex2 = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
      iIndexSt = InStr(iIndex2, RichTextBox1.Text, "</a", vbTextCompare)
      sCountryName = Mid(RichTextBox1.Text, iIndex2 + 1, (iIndexSt) - (iIndex2 + 1))
      sfndResult = FindStringinListControl(frmAlert.cmbcntyName, Trim(sCountryName))
      If sfndResult <> -1 Then
        frmAlert.lstPopulation.ListItems.Add , , sCountryName, , sfndResult + 1
      Else
        frmAlert.lstPopulation.ListItems.Add , , sCountryName ', , iCnt
      End If
      iCnt = iCnt + 1
    Else
      iIndex = InStr(iIndexSt, RichTextBox1.Text, "<td>", vbTextCompare)
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "</td>", vbTextCompare)
      sCountryName = Replace((Replace(Mid(RichTextBox1.Text, iIndex + 4, (iIndexSt) - (iIndex + 4)), Chr(10), " ")), "   ", "")
      If InStr(1, sCountryName, "<sup", vbTextCompare) <> 0 Then
        sCountryName = Mid(sCountryName, 1, InStr(1, sCountryName, "<sup", vbTextCompare) - 1)
      End If
      sCountryName = Replace(sCountryName, "&aacute;", "a")
      frmAlert.lstPopulation.ListItems(iCnt).ListSubItems.Add , , Replace(Replace(Replace(sCountryName, "<b>", ""), "</b>", ""), "&eacute;", "e"), , sCountryName
    End If
    
    If InStr(1, Mid(RichTextBox1.Text, iIndexSt, 40), "</table>", vbTextCompare) <> 0 Then
      Limits = 1
    End If
    X = X + 1
    ProgressBar1.Value = X
  Loop Until Limits = 1
  
  frmAlert.lstWeatherAlert.Visible = False
  frmAlert.txtAlert.Visible = False
  frmAlert.lstPopulation.Height = frmAlert.lstPopulation.Height - 100
  frmAlert.lstPopulation.Visible = True
  frmAlert.Caption = "Capital Of the World"
  ProgressBar1.Max = 1
  ProgressBar1.Visible = False
  frmWeatherMain.MousePointer = 0
  frmAlert.Show vbModal
End Sub

Private Sub GetCountrylHol(sContryLink As String, sHolYear As String)
  Dim iIndex, iIndex2 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim Limits As Integer
  Dim sPageName As String
  Dim sStartPos As String
  Dim iCnt As Integer
  Dim X As Integer
  Dim sCountryName As String
  
   sPageName = Replace(sContryLink, "2011", sHolYear)
  GetWebpage sPageName
  sStartPos = "Big Square"
  'DoEvents
  
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then Exit Sub
  frmAlert.lstPopulation.ColumnHeaders(1).Width = 3000
  frmAlert.lstPopulation.ColumnHeaders(2).Width = 5600
  frmAlert.lstPopulation.ColumnHeaders.Remove 4
  frmAlert.lstPopulation.ColumnHeaders.Remove 3
  frmAlert.lstPopulation.ColumnHeaders.Item(1).Text = "Date"
  frmAlert.lstPopulation.ColumnHeaders.Item(2).Text = "Holidays Name"
  
  Do
    If X Mod 2 = 0 Then
      iIndex = InStr(iIndexSt, RichTextBox1.Text, "class=cal", vbTextCompare)
      iIndex2 = InStr(iIndex + 13, RichTextBox1.Text, ">", vbTextCompare)
      iIndexSt = InStr(iIndex2, RichTextBox1.Text, "</", vbTextCompare)
      sCountryName = Mid(RichTextBox1.Text, iIndex2 + 1, (iIndexSt) - (iIndex2 + 1))
      frmAlert.lstPopulation.ListItems.Add , , sCountryName
      iCnt = iCnt + 1
    Else
      If InStr(1, Mid(RichTextBox1.Text, iIndexSt, 10), "<td>", vbTextCompare) <> 0 Then
        iIndex = InStr(iIndexSt + 6, RichTextBox1.Text, ">", vbTextCompare)
      Else
        iIndex = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
      End If
      iIndex2 = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
      iIndexSt = InStr(iIndex2, RichTextBox1.Text, "</", vbTextCompare)
      sCountryName = Mid(RichTextBox1.Text, iIndex2 + 1, (iIndexSt) - (iIndex2 + 1))
      frmAlert.lstPopulation.ListItems(iCnt).ListSubItems.Add , , sCountryName
    End If
    X = X + 1
    If InStr(1, Mid(RichTextBox1.Text, iIndexSt, 40), "</table>", vbTextCompare) <> 0 Then
      Limits = 1
    End If
  Loop Until Limits = 1
  
  frmAlert.lstWeatherAlert.Visible = False
  frmAlert.txtAlert.Visible = False
  frmAlert.lstPopulation.Height = frmAlert.lstPopulation.Height - 100
  frmAlert.lstPopulation.Visible = True
  frmAlert.Caption = "National Holidays In " & sHolYear & " For " & sSelCountryName '& " In " &
  frmAlert.Show vbModal
End Sub

Private Sub LoadCountryHol()
  Dim IndxCnt As Integer
  Dim nFileNum As Integer
  Dim sString As String
  Dim myArray() As String
  
  cmbCode.Clear
  nFileNum = FreeFile
  Open App.Path & "\Countries National Holidays.Dat" For Binary Access Read As #nFileNum
  'On Error Resume Next
  Do While Not EOF(nFileNum)
    'read the length of the string
    Get #nFileNum, , nLen
    'initialize the string with the correct number of spaces
    sString = Space$(nLen)
    Get #nFileNum, , sString
    sString = DecryptText((sString), sPassword, True)
    If Len(Trim$(sString)) > 1 Then
      myArray = Split(sString, ",")
      cmbCode.AddItem myArray(0) ' Trim(sString)
      ReDim Preserve LinkArray(IndxCnt)
      LinkArray(IndxCnt) = myArray(1)
      IndxCnt = IndxCnt + 1
    End If
  Loop
End Sub

Private Sub GetCountriesNatlHol()
  Dim iIndex, iIndex2 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim Limits As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim iCnt As Integer
  Dim X As Integer, iTotal As Integer
  Dim sCountryName As String
  Dim NameArray() As String
  Dim nameToStore As String
  Dim sCityName As String
  
  NameArray() = Split(HolDateSelect, "/")
  If NameArray(2) > year(Now) Or NameArray(2) < 2008 Then
    If NameArray(2) > year(Now) Then
      MsgBox year(Now) & " Is Maximum Date allowed", vbInformation, "Weather Of The Wearld"
    Else
      MsgBox "Year 2008 Is Minimum Date allowed", vbInformation, "Weather Of The Wearld"
    End If
    Exit Sub
  End If
  sPageName = "http://holidayyear.com/today.php?year=" & NameArray(2) & "&date=" & NameArray(1) & "&mon=" & NameArray(0)
  GetWebpage sPageName
  sStartPos = "Big Square"
  'DoEvents
  
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then
    MsgBox "Unable To Show Holiday For Date " & Format(HolDateSelect, "Long Date"), vbInformation, "Weather Of The World"
    Exit Sub
  End If
  frmAlert.lstPopulation.ColumnHeaders(1).Width = 4650
  frmAlert.lstPopulation.ColumnHeaders(2).Width = 3800
  frmAlert.lstPopulation.ColumnHeaders.Remove 4
  frmAlert.lstPopulation.ColumnHeaders.Remove 3
  frmAlert.lstPopulation.ColumnHeaders.Item(1).Text = "Holidays Name"
  frmAlert.lstPopulation.ColumnHeaders.Item(2).Text = "Country"
  RichTextBox1.Text = Replace(RichTextBox1.Text, Chr(10), "")
  Limits = Limits + 1
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "class=cal", vbTextCompare)
  iIndex2 = InStr(iIndex + 13, RichTextBox1.Text, ">", vbTextCompare)
  iIndexSt = InStr(iIndex2, RichTextBox1.Text, "</table", vbTextCompare)
  sCountryName = Mid(RichTextBox1.Text, iIndex2, (iIndexSt) - (iIndex2))
  If InStr(1, sCountryName, "href=", vbTextCompare) <> 0 Then
    sCountryName = HttpLinkRemove(sCountryName)
  End If
FrsItem:
  ReDim NameArray(0)
  iIndex = InStr(Limits, sCountryName, "</", vbTextCompare)
  If iIndex = 0 Then Exit Sub
  iIndex2 = InStrRev(sCountryName, ">", iIndex, vbTextCompare)
  sCityName = Mid(sCountryName, iIndex2 + 1, (iIndex) - (iIndex2 + 1))
  frmAlert.lstPopulation.ListItems.Add , , sCityName
  iTotal = iTotal + 1
  iCnt = iCnt + 1
  If InStr(1, Mid(sCountryName, iIndex, 60), "<select", vbTextCompare) <> 0 Then
    If InStr(iIndex, sCountryName, "<option>", vbTextCompare) <> 0 Then
      iIndexSt = InStr(Limits, sCountryName, "<option>", vbTextCompare)
      iIndexEnd = InStr(iIndexSt + 9, sCountryName, "</option>", vbTextCompare)
      nameToStore = Mid(sCountryName, iIndexSt + 8, (iIndexEnd) - (iIndexSt + 8))
      NameArray() = Split(nameToStore, ",")
    End If
    
    For X = 0 To UBound(NameArray, 1)
      frmAlert.lstPopulation.ListItems(iCnt).ListSubItems.Add , , NameArray(X)
      If X < UBound(NameArray, 1) Then
        frmAlert.lstPopulation.ListItems.Add , , Space((Len(sCityName) / 2) + 5) & "-" 'sCityName
        iCnt = iCnt + 1
      End If
    Next
  Else
    iIndexEnd = InStr(iIndex + 4, sCountryName, ">", vbTextCompare)
    iIndex2 = InStr(iIndexEnd, sCountryName, "</", vbTextCompare)
    sCityName = Mid(sCountryName, iIndexEnd + 1, (iIndex2) - (iIndexEnd + 1))
    frmAlert.lstPopulation.ListItems(iCnt).ListSubItems.Add , , Replace(sCityName, "<td>", "")
  End If
  If InStr(iIndexEnd, sCountryName, "class=cal", vbTextCompare) <> 0 Then
    iIndex = InStr(iIndexEnd, sCountryName, "class=cal", vbTextCompare)
    If InStr(1, Mid(sCountryName, iIndex, 20), "class=cal", vbTextCompare) <> 0 Then
      iIndexSt = InStr(iIndex + 1, sCountryName, ">", vbTextCompare)
      iIndex2 = InStr(iIndexSt, sCountryName, "</", vbTextCompare)
    ElseIf InStr(1, Mid(sCountryName, iIndex, 20), "<", vbTextCompare) = 0 Then
      iIndexSt = InStr(iIndex2, sCountryName, ">", vbTextCompare)
      iIndex2 = InStr(iIndexSt, sCountryName, "</", vbTextCompare)
      If InStr(1, Mid(sCountryName, iIndexSt + 1, (iIndex2) - (iIndexSt + 1)), "class=", vbTextCompare) = 0 Then
        frmAlert.lstPopulation.ListItems.Add , , sCityName
        iCnt = iCnt + 1
        iTotal = iTotal + 1
      Else
        frmAlert.lstPopulation.ListItems.Add , , sCityName
        iCnt = iCnt + 1
        iTotal = iTotal + 1
      End If
      iIndexSt = InStr(iIndex2 + 3, sCountryName, ">", vbTextCompare)
      iIndex2 = InStr(iIndexSt, sCountryName, "</", vbTextCompare)
      frmAlert.lstPopulation.ListItems(iCnt).ListSubItems.Add , , Mid(sCountryName, iIndexSt + 1, (iIndex2) - (iIndexSt + 1))
      iIndexSt = InStr(iIndex2, sCountryName, "<a", vbTextCompare)
    Else
      iIndexSt = InStr(iIndexEnd, sCountryName, "<a", vbTextCompare)
    End If
    Limits = iIndexSt
    GoTo FrsItem
  End If
  frmAlert.lstWeatherAlert.Visible = False
  frmAlert.txtAlert.Visible = False
  frmAlert.lstPopulation.Height = frmAlert.lstPopulation.Height - 100
  frmAlert.lstPopulation.Visible = True
  frmAlert.lblCountry.Visible = True
  frmAlert.lblCount.Visible = True
  frmAlert.lblCount.Caption = iTotal & " " & IIf(iTotal > 1, "Holidays", "Holiday")
  frmAlert.lblCountry.Caption = iCnt & " " & IIf(iCnt > 1, "Countries", "Country")
  frmAlert.Caption = "World Holidays on " & Format(HolDateSelect, "Long Date")
  frmAlert.Show vbModal
End Sub

Private Function HttpLinkRemove(StringToParse As String) As String
  Dim iStartIndex As Long
  Dim iEndIndex As Long
  Dim iNewIndes As Long
  Dim X As Integer
  Dim sCityNames As String
  Dim newString As String
  Dim NameArray() As String
  On Error GoTo errorHandler
  
  NameArray() = Split(StringToParse, "</")
  sCityNames = StringToParse
  newString = StringToParse
  For X = 0 To UBound(NameArray, 1)
    If InStr(1, StringToParse, "href=", vbTextCompare) <> 0 Then
      iNewIndes = InStr(1, StringToParse, " class=link1 href=", vbTextCompare)
      iStartIndex = InStr(iNewIndes, StringToParse, ">", vbTextCompare)
      sCityNames = Mid(sCityNames, 1, iNewIndes) & Mid(sCityNames, iStartIndex)
    End If
    StringToParse = sCityNames
  Next
  sCityNames = Replace(sCityNames, "</tr><tr", "")
  sCityNames = Replace(sCityNames, "<td>", "")
  sCityNames = Replace(sCityNames, "</option><option>", ",")
  
  If InStr(1, sCityNames, "select", vbTextCompare) <> 0 Then
    sCityNames = Replace(sCityNames, "</option><option>", ",")
  End If
  HttpLinkRemove = sCityNames
  Exit Function
errorHandler:
  MsgBox "No Holiday To Show For Date " & Format(HolDateSelect, "Long Date"), vbInformation, "Weather Of The World"
  Unload frmAlert
End Function

Private Sub GetCountryAnthem(sCntryCode As String, sCntryName As String)
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sCountryAbr As String
  
  'On Error GoTo errorHandler
  MousePointer = 11
  sPageName = "http://www.studentsoftheworld.info/country_information.php?Pays=" & sCntryCode
  GetWebpage sPageName
  sStartPos = "National Anthem"
  iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndex = 0 Then
    MsgBox "Unable to show " & sCntryName & " Anthem", vbInformation, "Weather Of The World"
    MousePointer = 0
    Exit Sub
  End If
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "<TEXTAREA", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</TEXTAREA>", vbTextCompare)
  
  sCountryAbr = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
  frmAlert.rchTxtAnthem.Text = sCntryName & " National Anthem" & vbCrLf & vbCrLf & sCountryAbr
  frmAlert.rchTxtAnthem.Font.size = 11
  frmAlert.rchTxtAnthem.Visible = True
  frmAlert.Caption = sCntryName & " National Anthem"
  MousePointer = 0
  frmAlert.Show vbModal
End Sub


Private Sub LoadNatAnthem()
  Dim IndxCnt As Integer
  Dim nFileNum As Integer
  Dim sString As String
  Dim myArray() As String
  
  cmbAnthem.Clear
  nFileNum = FreeFile
  Open App.Path & "\Countries National Anthem.Dat" For Binary Access Read As #nFileNum
  'On Error Resume Next
  Do While Not EOF(nFileNum)
    'read the length of the string
    Get #nFileNum, , nLen
    'initialize the string with the correct number of spaces
    sString = Space$(nLen)
    Get #nFileNum, , sString
    sString = DecryptText((sString), sPassword, True)
    If Len(Trim$(sString)) > 1 Then
      myArray = Split(sString, ",")
      cmbAnthem.AddItem Trim(myArray(1))
      ReDim Preserve AnthemArray(IndxCnt)
      AnthemArray(IndxCnt) = myArray(0)
      IndxCnt = IndxCnt + 1
    End If
  Loop
End Sub

Private Sub GetCountryPhoneCode(sCountryName As String)
  Dim imageUrl As String
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sCountryAbr As String
  Dim PhoneCode As String
  Dim sInfo As String
  Dim sInfo1 As String
  Dim sInfo2 As String
  Dim sInfo3 As String
  Dim sInfoIDD As String
  Dim sInfoNDD As String
  Dim X As Integer, Limits As Integer
  Dim iNameCnt As Integer
  Dim ctycnt As Integer
  
  'On Error GoTo errorHandler
  sPageName = "http://countrycode.org/" & sCountryName
  GetWebpage sPageName
  sStartPos = "main_table_blue"
  iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndex = 0 Then
    MsgBox "Unable to show " & sSelCountryName & " Country Code", vbInformation, "Weather Of The World"
    Unload frmPoneCode
    Exit Sub
  End If
  'test for country code
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "<h1", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
  
  If Val(Right(Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1)), 2)) <> 0 Then
    iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<td width=", vbTextCompare)
    iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
    sCountryAbr = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
    
    iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "class=", vbTextCompare)
    iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
    PhoneCode = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
  End If
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<img src=", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, " alt=", vbTextCompare)
  imageUrl = "http://countrycode.org/" & Mid(RichTextBox1.Text, iIndexSt + 10, (iIndex - 1) - (iIndexSt + 10))
  SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), frmPoneCode.ImgCntFlag
  
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<p>", vbTextCompare)
  
  iIndexSt = InStr(iIndexEnd + 5, RichTextBox1.Text, "<p>", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
  sInfo = Mid(RichTextBox1.Text, iIndexSt + 3, (iIndex) - (iIndexSt + 3))
  
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<p>", vbTextCompare)
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "</", vbTextCompare)
  sInfo1 = Mid(RichTextBox1.Text, iIndexEnd + 3, (iIndexSt) - (iIndexEnd + 3))
  frmPoneCode.lblInfo.Caption = sInfo & vbCrLf & sInfo1
  For X = 0 To 1
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "class=", vbTextCompare)
    iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, ">", vbTextCompare)
    iIndex = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
    If X = 0 Then
      sInfo = Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1)), "<b>", "")
    Else
      sInfo2 = Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1)), "<b>", "")
    End If
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "<td class=", vbTextCompare)
    iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
    If X = 0 Then
      sInfo1 = Replace(Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1)), "&nbsp;", "")
    Else
      sInfo3 = Replace(Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1)), "&nbsp;", "")
    End If
    
    iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<br>", vbTextCompare)
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
    If X = 0 Then
      sInfoIDD = Replace(Mid(RichTextBox1.Text, iIndex + 10, (iIndexSt - 5) - (iIndex + 10)), Chr(13), "")
    Else
      sInfoNDD = Mid(RichTextBox1.Text, iIndex + 5, (iIndexSt - 1) - (iIndex + 5))
    End If
    iIndexSt = InStr(iIndexSt, RichTextBox1.Text, "Value", vbTextCompare)
  Next
  'Display city code
  iIndexSt = InStr(iIndexSt, RichTextBox1.Text, "common_table", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "width=", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "</", vbTextCompare)
  frmPoneCode.lblNoCity.Caption = Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1))
  frmPoneCode.lblCityCount.Caption = Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1))
  Do
    If ctycnt Mod 4 = 0 Then
      iIndex = InStr(iIndexSt, RichTextBox1.Text, "<td align=", vbTextCompare)
      If iIndex = 0 Then
        frmPoneCode.lblNoCity.Caption = "No " & frmPoneCode.lblNoCity.Caption
        GoTo xfault
      End If
      iIndexEnd = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
      iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "</", vbTextCompare)
      If InStr(1, Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1)), " not ", vbTextCompare) <> 0 Then
        frmPoneCode.lblNoCity.Caption = Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1))
xfault:
        frmPoneCode.lstPhoneCode.Visible = False
        Exit Do
      Else
        frmPoneCode.lstPhoneCode.Visible = True
      End If
      frmPoneCode.lstPhoneCode.ListItems.Add , , Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1))
      iNameCnt = iNameCnt + 1
    Else
      If ctycnt Mod 2 = 1 Then
        iIndex = InStr(iIndexSt, RichTextBox1.Text, "<b", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
        frmPoneCode.lstPhoneCode.ListItems(iNameCnt).ListSubItems.Add , , Replace(Replace(Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1)), "&nbsp;&nbsp;", ""), "<br>", " ")
      Else
        iIndex = InStr(iIndexSt, RichTextBox1.Text, "<td align=", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "</", vbTextCompare)
        frmPoneCode.lstPhoneCode.ListItems(iNameCnt).ListSubItems.Add , , Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1))
      End If
    End If
    'test for end table
    If InStr(1, Mid(RichTextBox1.Text, iIndexSt, 60), "</table>", vbTextCompare) <> 0 Then
      Limits = 1
    End If
    ctycnt = ctycnt + 1
  Loop Until Limits = 1
  
  frmPoneCode.lblIDDInfo.Caption = Replace(sInfo & " " & sInfo1 & vbCrLf & sInfoIDD & sInfo2 & " " & sInfo3 & vbCrLf & sInfoNDD, "<br>", " ")
  frmPoneCode.lblcontryName.Caption = sSelCountryName & " " & sCountryAbr & " " & PhoneCode
  frmPoneCode.lblCityCount.Caption = Mid(frmPoneCode.lblCityCount.Caption, 1, InStr(1, frmPoneCode.lblCityCount.Caption, "City Codes") - 1) & "Has " & ctycnt \ 2 & " City Code(s)"
  frmPoneCode.Caption = sSelCountryName & " international dialing Code"
  If Len(frmPoneCode.lblcontryName.Caption) >= 48 Then
    frmPoneCode.lblcontryName.FontSize = 13
  Else
    frmPoneCode.lblcontryName.FontSize = 14
  End If
  frmPoneCode.Show vbModal
End Sub

Private Sub LaodPhoneCode()
  Dim IndxCnt As Integer
  Dim nFileNum As Integer
  Dim sString As String
  Dim myArray() As String
  
  cmbPhCode.Clear
  nFileNum = FreeFile
  Open App.Path & "\Countries Phone Code.Dat" For Binary Access Read As #nFileNum
  'On Error Resume Next
  Do While Not EOF(nFileNum)
    'read the length of the string
    Get #nFileNum, , nLen
    'initialize the string with the correct number of spaces
    sString = Space$(nLen)
    Get #nFileNum, , sString
    sString = DecryptText((sString), sPassword, True)
    If Len(Trim$(sString)) > 1 Then
      myArray = Split(sString, ",")
      cmbPhCode.AddItem Trim(myArray(1))
      ReDim Preserve PhoneArray(IndxCnt)
      PhoneArray(IndxCnt) = myArray(0)
      IndxCnt = IndxCnt + 1
    End If
  Loop
End Sub

Public Sub getCountryStatic(sCountryName As String)
  Dim imageUrl As String
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sInfo As String
  Dim sInfo1 As String
  Dim sInfo2 As String
  Dim sInfo3 As String
  Dim X As Integer, Limits As Integer
  Dim iNameCnt As Integer, sNameLen As Integer
  Dim ctycnt As Integer
  Dim sParseText As String
  Dim sAreaText As String
  Dim sIsoCode2 As String, sIsoCode3 As String
  
  'On Error GoTo errorHandler
  sPageName = "http://countrycode.org/" & sCountryName
  GetWebpage sPageName
  
  frmPoneCode.lstPhoneCode.ColumnHeaders(1).Width = 3450
  frmPoneCode.lstPhoneCode.ColumnHeaders(2).Width = 5110
  frmPoneCode.lstPhoneCode.HideColumnHeaders = True
  frmPoneCode.lstPhoneCode.ColumnHeaders.Remove 4
  frmPoneCode.lstPhoneCode.ColumnHeaders.Remove 3
  frmPoneCode.lstPhoneCode.GridLines = False
  frmPoneCode.lblInfo.FontBold = True
  frmPoneCode.lblInfo.ForeColor = vbBlue
  frmPoneCode.lblIDDInfo.ForeColor = vbBlue
  frmPoneCode.lblIDDInfo.FontBold = True
  frmPoneCode.lblCityCount.Visible = False
  frmPoneCode.lblNoCity.FontUnderline = True
  frmPoneCode.Frame1.Caption = sCountryName & " Statistics"
  
  sStartPos = "main_table_blue"
  iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndex = 0 Then
    MsgBox "Unable to show " & sSelCountryName & " Country Code", vbInformation, "Weather Of The World"
    Unload frmPoneCode
    Exit Sub
  End If
  
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "ISO Country Code, 2 Digit:", vbTextCompare)
  If iIndexSt = 0 Then
    frmPoneCode.lblCityCount.Visible = False
    GoTo noISO
  Else
    frmPoneCode.lblCityCount.Visible = True
  End If
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "class=", vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
  sIsoCode2 = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1))
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "class=", vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
  sIsoCode3 = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1))
noISO:
  'Get country flag
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "<img src=", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, " alt=", vbTextCompare)
  imageUrl = "http://countrycode.org/" & Mid(RichTextBox1.Text, iIndexSt + 10, (iIndex - 1) - (iIndexSt + 10))
  SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), frmPoneCode.ImgCntFlag
  
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "colspan=", vbTextCompare)
  iIndexSt = InStr(iIndexEnd + 5, RichTextBox1.Text, ">", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
  frmPoneCode.lblcontryName.Caption = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1))
  frmPoneCode.lblNoCity = "More " & Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1))
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "bottom_tt", vbTextCompare)
  iIndex = InStr(iIndexEnd + 20, RichTextBox1.Text, "bottom_tt", vbTextCompare)
  
  For X = 0 To 1
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "label", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
    iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</", vbTextCompare)
    sInfo = Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex) - (iIndexEnd + 1))
    If X = 0 Then
      frmPoneCode.lstPhoneCode.ListItems.Add , , sInfo
      iNameCnt = iNameCnt + 1
    End If
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "Value", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
    iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<", vbTextCompare)
    sInfo1 = Replace(Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex) - (iIndexEnd + 1)), "&nbsp;", "")
    If X = 0 Then
      frmPoneCode.lstPhoneCode.ListItems(iNameCnt).ListSubItems.Add , , Trim(sInfo1)
    End If
  Next
  'Get Electrical Outlet picture
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "<table", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "label", vbTextCompare)
  GetElectricPic Mid(RichTextBox1.Text, iIndexSt, iIndex - iIndexSt)
  
  'Get phone jack picture
  iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
  frmPoneCode.lblIDDInfo.Caption = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1))
  
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "<table", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "label", vbTextCompare)
  If iIndex <> 0 Then
    GetPhonePic Mid(RichTextBox1.Text, iIndexSt, iIndex - iIndexSt)
  Else
    iIndex = InStr(1, RichTextBox1.Text, "Outlet", vbTextCompare)
  End If
  Do
    If ctycnt Mod 2 = 0 Then
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "label", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</", vbTextCompare)
      frmPoneCode.lstPhoneCode.ListItems.Add , , Trim(Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex) - (iIndexEnd + 1)))
      iNameCnt = iNameCnt + 1
    Else
      iIndexSt = InStr(iIndex, RichTextBox1.Text, "Value", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
      If InStr(1, frmPoneCode.lstPhoneCode.ListItems(iNameCnt).Text, "Area", vbTextCompare) <> 0 Then
        iIndex = InStr(iIndexEnd, RichTextBox1.Text, "&nbsp;", vbTextCompare)
      Else
        iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</", vbTextCompare)
      End If
      sParseText = Replace(Replace(Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndex) - (iIndexEnd + 1)), "&nbsp;", ""), " <span class=""rank_value"">", " ")
      If InStr(1, frmPoneCode.lstPhoneCode.ListItems(iNameCnt).Text, "Languages", vbTextCompare) <> 0 And Len(sParseText) > 50 Then
        sNameLen = Len(frmPoneCode.lstPhoneCode.ListItems(iNameCnt).Text)
        frmPoneCode.lstPhoneCode.ListItems(iNameCnt).ListSubItems.Add , , Trim(Mid(sParseText, 1, InStr(40, sParseText, " ")))
        Do
          sParseText = Mid(sParseText, InStr(40, sParseText, " ") + 1)
          If Len(sParseText) > 50 And InStr(41, sParseText, " ") <> 0 Then
            frmPoneCode.lstPhoneCode.ListItems.Add , , Space(sNameLen) & "-"
            iNameCnt = iNameCnt + 1
            frmPoneCode.lstPhoneCode.ListItems(iNameCnt).ListSubItems.Add , , Trim(Mid(sParseText, 1, InStr(40, sParseText, " ")))
          Else
            frmPoneCode.lstPhoneCode.ListItems.Add , , Space(sNameLen) & "-"
            iNameCnt = iNameCnt + 1
            frmPoneCode.lstPhoneCode.ListItems(iNameCnt).ListSubItems.Add , , Trim(sParseText)
            Exit Do
          End If
        Loop Until InStr(1, sParseText, " ", vbTextCompare) = 0
      ElseIf InStr(1, frmPoneCode.lstPhoneCode.ListItems(iNameCnt).Text, "Area", vbTextCompare) <> 0 Then
        sAreaText = Replace(Replace(Mid(sParseText, InStr(1, sParseText, "<br>") + 5), Chr(9), ""), "   ", "")
        frmPoneCode.lstPhoneCode.ListItems(iNameCnt).ListSubItems.Add , , Trim(Mid(sParseText, 1, InStr(1, sParseText, "</") - 1) & "  " & sAreaText)
      Else
        frmPoneCode.lstPhoneCode.ListItems(iNameCnt).ListSubItems.Add , , Trim(sParseText)
      End If
    End If
    If InStr(1, Mid(RichTextBox1.Text, iIndex, 60), "</table>", vbTextCompare) <> 0 Then
      Limits = 1
    End If
    ctycnt = ctycnt + 1
  Loop Until Limits = 1
  
  frmPoneCode.lblCityCount.Caption = "ISO Country Code " & sIsoCode2 & "/" & sIsoCode3
  If iNameCnt < 17 Then
    frmPoneCode.lblRanking.Visible = False
    frmPoneCode.lstPhoneCode.ListItems.Add , , "--------------------------------------------------------"
    frmPoneCode.lstPhoneCode.ListItems(iNameCnt + 1).ListSubItems.Add , , "--------------------------------------------------------------------------"
    frmPoneCode.lstPhoneCode.ListItems.Add , , Space(40) & "** In Brackets Are"
    frmPoneCode.lstPhoneCode.ListItems(iNameCnt + 2).ListSubItems.Add , , "World Ranking **"
  Else
    frmPoneCode.lblRanking.Caption = "** In Brackets Are World Ranking **"
    frmPoneCode.lblRanking.Visible = True
  End If
  frmPoneCode.lblInfo.Caption = sInfo & " " & sInfo1
  frmPoneCode.lstPhoneCode.Visible = True
  If Len(frmPoneCode.lblcontryName.Caption) >= 48 Then
    frmPoneCode.lblcontryName.FontSize = 13
  Else
    frmPoneCode.lblcontryName.FontSize = 14
  End If
  MousePointer = 0
  frmPoneCode.Show vbModal
End Sub

Private Sub GetElectricPic(sElecPicLink As String)
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim imageUrl As String, sInfo As String
  Dim X As Integer
  
  iIndex = 1
  Do
    iIndexSt = InStr(iIndex, sElecPicLink, "<img src=", vbTextCompare)
    iIndex = InStr(iIndexSt, sElecPicLink, " width=", vbTextCompare)
    imageUrl = "http://countrycode.org/" & Mid(sElecPicLink, iIndexSt + 10, (iIndex - 1) - (iIndexSt + 10))
    SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), frmPoneCode.imgElStat(X)
  
    iIndexSt = InStr(iIndex, sElecPicLink, "style=", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, sElecPicLink, ">", vbTextCompare)
    iIndexSt = InStr(iIndexEnd, sElecPicLink, "</", vbTextCompare)
    sInfo = Mid(sElecPicLink, iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1))
    frmPoneCode.lblElec(X).Caption = Replace(sInfo, "<br><nobr>", "")
    'test for more picture
    iIndex = InStr(iIndexSt, sElecPicLink, "<table", vbTextCompare)
    
    X = X + 1
    If X > 3 Or iIndex = 0 Then
      Exit Do
    End If
  Loop Until InStr(1, sElecPicLink, "</table", vbTextCompare) = 0
End Sub

Private Sub GetPhonePic(sPhonePicLink As String)
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim imageUrl As String, sInfo As String
  Dim X As Integer
  
  iIndex = 1
  Do
    iIndexSt = InStr(iIndex, sPhonePicLink, "<img src=", vbTextCompare)
    iIndex = InStr(iIndexSt, sPhonePicLink, " width=", vbTextCompare)
    imageUrl = "http://countrycode.org/" & Mid(sPhonePicLink, iIndexSt + 10, (iIndex - 1) - (iIndexSt + 10))
    SavePngFille imageUrl, Mid(imageUrl, InStrRev(imageUrl, "/") + 1), frmPoneCode.imgPHStat(X)
  
    iIndexSt = InStr(iIndex, sPhonePicLink, "style=", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, sPhonePicLink, ">", vbTextCompare)
    iIndexSt = InStr(iIndexEnd, sPhonePicLink, "&nbsp;", vbTextCompare)
    sInfo = Mid(sPhonePicLink, iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1))
    frmPoneCode.lblphone(X).Caption = Replace(sInfo, "<br><nobr>", "")
    'test for more picture
    iIndex = InStr(iIndexSt, sPhonePicLink, "<table", vbTextCompare)
    X = X + 1
    If X > 3 Or iIndex = 0 Then
      Exit Do
    End If
  Loop Until InStr(1, sPhonePicLink, "</table", vbTextCompare) = 0
End Sub

Private Sub GetRaceofCountry()
  Dim iIndex, iIndex2 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sCnt As Integer
  Dim X As Integer, sNameLen As Integer
  Dim sCountryName As String, sRaceName As String
  Dim sParseText As String
  
  MousePointer = 11
  sPageName = "http://www.infoplease.com/ipa/A0855617.html"
  GetWebpage sPageName
  sStartPos = "BodyText"
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then Exit Sub
  
  frmPoneCode.lstRace.ColumnHeaders(1).Width = 3000
  frmPoneCode.lstRace.ColumnHeaders(2).Width = 5500
  frmPoneCode.Frame1.Caption = "Ethnicity and Race by Countries"
  frmPoneCode.lstRace.GridLines = True
  frmPoneCode.lstRace.FullRowSelect = True
  ProgressBar1.Scrolling = ccScrollingSmooth
  ProgressBar1.Max = 194
  ProgressBar1.Visible = True
  sCnt = 1
  
  Do
    'Country
    iIndex = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
    If iIndex = 0 Then
      MousePointer = 0
      Exit Do
    End If
    iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
    sCountryName = Replace(Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1)), "&iacute;", "i"), "&ocirc;", "o")
    sCountryName = Replace(Replace(sCountryName, "&eacute;", "e"), "&atilde;", "a")
    frmPoneCode.lstRace.ListItems.Add , , sCountryName
    X = X + 1
    iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<td", vbTextCompare)
    iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</td", vbTextCompare)
    sRaceName = Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1)), Chr(10), "")
    sParseText = Replace(sRaceName, "<supclass=""fnr"">1</sup>", "")
    
    If Len(sRaceName) > 50 Then
        sNameLen = Len(frmPoneCode.lstRace.ListItems(X).Text)
        frmPoneCode.lstRace.ListItems(X).ListSubItems.Add , , Trim(Mid(sParseText, 1, InStr(45, sParseText, " ")))
        Do
          sParseText = Mid(sParseText, InStr(45, sParseText, " ") + 1)
          If Len(sParseText) > 50 And InStr(46, sParseText, " ") <> 0 Then
            frmPoneCode.lstRace.ListItems.Add , , Space(sNameLen) & "-"
            X = X + 1
            frmPoneCode.lstRace.ListItems(X).ListSubItems.Add , , Trim(Mid(sParseText, 1, InStr(45, sParseText, " ")))
          Else
            frmPoneCode.lstRace.ListItems.Add , , Space(sNameLen) & "-"
            X = X + 1
            frmPoneCode.lstRace.ListItems(X).ListSubItems.Add , , Trim(sParseText)
            Exit Do
          End If
        Loop Until InStr(1, sParseText, " ", vbTextCompare) = 0
      Else
        frmPoneCode.lstRace.ListItems(X).ListSubItems.Add , , Trim(sParseText)
      End If
      
    If sCountryName = "Zimbabwe" Then Exit Do
    iIndexSt = iIndexEnd
    ProgressBar1.Value = sCnt
    sCnt = sCnt + 1
  Loop Until sCountryName = "Zimbabwe"
  frmPoneCode.lstRace.Visible = True
  ProgressBar1.Max = 1
  ProgressBar1.Visible = False
  MousePointer = 0
  frmPoneCode.Show vbModal
End Sub

Private Sub DisableMenu(bMenuAble As Boolean)
  mnuCountryStat.Enabled = bMenuAble
   mnuEdit.Enabled = bMenuAble
   mnuFile.Enabled = bMenuAble
   mnuBook.Enabled = bMenuAble
   cmdFar.Enabled = bMenuAble
   cmdCel.Enabled = bMenuAble
   mnuStorm.Enabled = bMenuAble
   mnuShowMap.Enabled = bMenuAble
   mnuSatellite.Enabled = bMenuAble
   mnuRadar.Enabled = bMenuAble
   mnuWeather.Enabled = bMenuAble
   mnuPopStatistics.Enabled = bMenuAble
   mnuWorld.Enabled = bMenuAble
   cmdNext.Enabled = bMenuAble
   cmdPrevious.Enabled = bMenuAble
   If bMenuAble Then
    If iredoIndex <> 0 Then
       bPreState = True
    End If
    cmdPrevious.Enabled = bPreState
    cmdNext.Enabled = bNextState
    If IsCelsius Then
       cmdFar.Enabled = True
       cmdCel.Enabled = False
       mnuFar.Checked = False
       mnuCel.Checked = True
    Else
       cmdFar.Enabled = False
       cmdCel.Enabled = True
       mnuFar.Checked = True
       mnuCel.Checked = False
    End If
  End If
End Sub

Public Function GetStateAlertType(StringToParse As String, iCount As Integer) As String
  Dim iStartIndex As Long
  Dim iEndIndex As Long
  Dim iNewIndes As Long
  Dim X As Integer
  Dim sCityNames As String
  Dim newString As String
  Dim NameArray() As String
  
  ReDim NameArray(0)
  iEndIndex = 1
  NameArray() = Split(StringToParse, "href=")
  
  newString = StringToParse
  For X = 0 To UBound(NameArray, 1) - 1
    iNewIndes = InStr(iEndIndex, StringToParse, "href=", vbTextCompare)
    iStartIndex = InStr(iNewIndes, StringToParse, ">", vbTextCompare)
    iEndIndex = InStr(iStartIndex, StringToParse, "</", vbTextCompare)
    sCityNames = Mid(StringToParse, iStartIndex + 1, (iEndIndex) - (iStartIndex + 1))
    If X = 0 Then
      newString = sCityNames
    Else
      newString = newString & " " & sCityNames
    End If
  Next
  GetStateAlertType = newString
End Function

Private Sub GetCurrentTrack(uRlLink As String, mIndex As Integer)
  Dim iIndex As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim nFileNum As Integer
  Dim myFile As String
  Dim myData() As Byte
  Dim sFieName As String
  Dim sPageName As String
  
  On Error Resume Next
  sPageName = "http://www.intellicast.com" & uRlLink
  GetWebpage sPageName
  
  nFileNum = FreeFile
  iIndexSt = InStr(1, RichTextBox1.Text, "Content Container", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "src=", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, " ", vbTextCompare)
  myFile = Mid(RichTextBox1.Text, iIndex + 5, (iIndexEnd - 1) - (iIndex + 5))
  Select Case mIndex
    Case 0
      myData() = Inet1.OpenURL(myFile, icByteArray)
      sFieName = "Large-" & Mid(myFile, InStrRev(myFile, "/") + 1)
          
      Open App.Path + "\Icons\" & sFieName For Binary Access Write As #nFileNum
        Put #nFileNum, , myData()
      Close #nFileNum
      picTureName = App.Path + "\Icons\" & sFieName
      Load frmCountry
    Case 1, 2
      AnimationLink = myFile
      frmAnimate.Show vbModal
  End Select
End Sub

Private Sub GetStateAlerts(sHurLink As String, sStateName As String)
  Dim iIndex, iIndex2 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim Limits As Integer
  Dim cnt As Integer
  Dim Alertcnt As Integer
  Dim StringToParse As String
  Dim X As Integer
  Dim NameArray() As String
  
  'On Error Resume Next
  sPageName = "http://www.intellicast.com" & sHurLink
  GetWebpage sPageName
  sStartPos = "Weather Alerts:"
  
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then Exit Sub
  
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "href=", vbTextCompare)
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "</div", vbTextCompare)
  StringToParse = Mid(RichTextBox1.Text, iIndexEnd, (iIndex) - (iIndexEnd))
  
  NameArray() = Split(StringToParse, "class=")
  For X = 1 To UBound(NameArray, 1)
    If InStr(1, NameArray(X), "Alert", vbTextCompare) <> 0 Then
      iIndex = InStr(1, NameArray(X), "strong", vbTextCompare)
      iIndexEnd = InStr(iIndex, NameArray(X), "</", vbTextCompare)
      frmAlert.lstWeatherAlert.ListItems.Add , , ""
      Alertcnt = Alertcnt + 1
      cnt = 0
      frmAlert.lstWeatherAlert.ListItems(Alertcnt).ListSubItems.Add , , Mid(NameArray(X), iIndex + 7, (iIndexEnd) - (iIndex + 7))
      frmAlert.lstWeatherAlert.ListItems.Add , , ""
      Alertcnt = Alertcnt + 1
    End If
    Do
      Limits = 0
      If cnt Mod 3 = 0 Then
        iIndexSt = InStr(iIndexEnd, NameArray(X), "href=", vbTextCompare)
        iIndex = InStr(iIndexSt, NameArray(X), ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, NameArray(X), "</", vbTextCompare)
        frmAlert.lstWeatherAlert.ListItems.Add , , Mid(NameArray(X), iIndex + 1, (iIndexEnd) - (iIndex + 1))
        Alertcnt = Alertcnt + 1
      Else
        iIndexSt = InStr(iIndexEnd, NameArray(X), "href=", vbTextCompare)
        iIndex = InStr(iIndexSt, NameArray(X), ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, NameArray(X), "</", vbTextCompare)
        frmAlert.lstWeatherAlert.ListItems(Alertcnt).ListSubItems.Add , , Mid(NameArray(X), iIndex + 1, (iIndexEnd) - (iIndex + 1))
      End If
      cnt = cnt + 1
      If InStr(1, Mid(NameArray(X), iIndexEnd, 100), "href=", vbTextCompare) = 0 Then
        Limits = 1
      End If
    Loop Until Limits = 1
    frmAlert.lstWeatherAlert.ListItems.Add , , ""
    Alertcnt = Alertcnt + 1
  Next

  frmAlert.lstWeatherAlert.Visible = True
  frmAlert.txtAlert.Visible = False
  frmAlert.Caption = "Weather Alerts: " & sStateName
  frmAlert.Show vbModal
End Sub

Private Sub GetRadSummary()
  Dim iIndex, iIndex2, iIndex3 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim cnt As Integer
  
  GetWebpage "http://www.intellicast.com/National/Radar/Summary.aspx"
  
  'DoEvents
  cnt = 0
  iIndexEnd = InStr(1, RichTextBox1.Text, "Region:", vbTextCompare)
  Do While cnt < 38
    iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "value=", vbTextCompare)
    iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
    muuRadarSummary(cnt).Tag = Mid(RichTextBox1.Text, iIndex2 + 7, (iIndex - 1) - (iIndex2 + 7))
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "</option>", vbTextCompare)
    muuRadarSummary(cnt).Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
    muuRadarSummary(cnt).Visible = True
    cnt = cnt + 1
    iIndexEnd = iIndexSt
  Loop
End Sub

Private Sub DisplayRadarMap()
  Dim iIndex As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim nFileNum As Integer
  Dim myFile As String
  Dim myData() As Byte
  Dim sFieName As String
  
  On Error Resume Next
  MousePointer = 11
  nFileNum = FreeFile
  iIndexSt = InStr(1, RichTextBox1.Text, "Content Container", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "src=", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, " ", vbTextCompare)
  myFile = Mid(RichTextBox1.Text, iIndex + 5, (iIndexEnd - 1) - (iIndex + 5))
  myData() = Inet1.OpenURL(myFile, icByteArray)
  sFieName = "Large-" & Mid(myFile, InStrRev(myFile, "/") + 1)
      
  Open App.Path + "\Icons\" & sFieName For Binary Access Write As #nFileNum
    Put #nFileNum, , myData()
  Close #nFileNum
  picTureName = App.Path + "\Icons\" & sFieName
  MousePointer = 0
  Load frmCountry
End Sub

Private Sub GetBulletins()
  Dim iIndex, iIndex2, iIndex3 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim cnt As Integer
  Dim sStormName As String
  
  GetWebpage "http://www.intellicast.com/Storm/Hurricane/Track.aspx"
  
  cnt = 0
  iIndexEnd = InStr(1, RichTextBox1.Text, "Active Storm Track", vbTextCompare)
  Do
    iIndex2 = InStr(iIndexEnd, RichTextBox1.Text, "Alert", vbTextCompare)
    If iIndex2 = 0 Then Exit Do
    iIndex = InStr(iIndex2, RichTextBox1.Text, ">", vbTextCompare)
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
    sStormName = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1))
    
    iIndex2 = InStr(iIndexSt, RichTextBox1.Text, "Bulletins:", vbTextCompare)
    iIndex = InStr(iIndex2, RichTextBox1.Text, "href=", vbTextCompare)
    iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    mnuStormAdvisory(cnt).Tag = Mid(RichTextBox1.Text, iIndex + 6, (iIndexSt - 1) - (iIndex + 6))
    
    iIndex2 = InStr(iIndexSt, RichTextBox1.Text, " ", vbTextCompare)
    iIndexEnd = InStr(iIndex2, RichTextBox1.Text, "</", vbTextCompare)
    mnuStormAdvisory(cnt).Caption = sStormName & " " & Mid(RichTextBox1.Text, iIndex2 + 1, (iIndexEnd - 1) - (iIndex2 + 1))
    If InStr(iIndexSt, RichTextBox1.Text, "Bulletins:", vbTextCompare) = 0 Then
      Exit Do
    End If
    cnt = cnt + 1
    mnuStormAdvisory(cnt).Visible = True
  Loop
End Sub

Private Sub GetWeatherAdvisory(sHurLink As String)
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sBulletinName As String
  Dim sBulletinText As String
  
  'On Error GoTo errorHandler
  MousePointer = 11
  sPageName = "http://www.intellicast.com" & sHurLink
  GetWebpage sPageName
  sStartPos = "BulletinName"
  iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndex = 0 Then
    MsgBox "Unable to show " & sBulletinName, vbInformation, "Weather Of The World"
    Exit Sub
  End If
  iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
  sBulletinName = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1))
  
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "BulletinText", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</span", vbTextCompare)
  sBulletinText = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
  
  sBulletinText = Replace(Replace(sBulletinText, "   ", ""), "...", " ")

  frmAlert.rchTxtAnthem.Text = sBulletinName & vbCrLf & vbCrLf & StrConv(Replace(sBulletinText, "<br/>", vbCrLf), vbProperCase)
  frmAlert.rchTxtAnthem.Font.size = 11
  frmAlert.rchTxtAnthem.Visible = True
  frmAlert.Caption = sBulletinName
  MousePointer = 0
  frmAlert.Show vbModal
End Sub

Private Sub GerWeatherBulletins()
  Dim iIndex, iIndex2, iIndex3 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim iLinecount As Integer
  Dim Limits As Integer
  Dim Alertcnt As Integer
  Dim sStateLink As String
   
  On Error Resume Next
  GetWebpage "http://www.intellicast.com/Storm/Severe/Bulletins.aspx"
  sStartPos = "Weather Alerts:"
  
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then Exit Sub
  mnuAlertState.Enabled = True
  
  For Alertcnt = 0 To 61
    mnuStateAlert(Alertcnt).Visible = False
    mnuStateAlert(Alertcnt).Caption = ""
  Next
  Alertcnt = 0
  Do
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "style=", vbTextCompare)
    iIndex3 = InStr(iIndexEnd, RichTextBox1.Text, ">", vbTextCompare)
    iIndex = InStr(iIndex3, RichTextBox1.Text, "</a>", vbTextCompare)
    
    If InStr(1, Mid(RichTextBox1.Text, iIndexEnd + 7, (iIndex3 - 2) - (iIndexEnd + 7)), "color:#900", vbTextCompare) <> 0 Then
      sStateLink = Mid(RichTextBox1.Text, InStrRev(RichTextBox1.Text, "href=", iIndex3) + 6, InStrRev(RichTextBox1.Text, " ", iIndex3) - InStrRev(RichTextBox1.Text, "href=", iIndex3) - 7)
      mnuStateAlert(Alertcnt).Caption = Mid(RichTextBox1.Text, iIndex3 + 1, (iIndex) - (iIndex3 + 1))
      mnuStateAlert(Alertcnt).Tag = sStateLink
      mnuStateAlert(Alertcnt).Visible = True
      Alertcnt = Alertcnt + 1
    End If
     
    If InStr(1, Mid(RichTextBox1.Text, iIndex, 40), "</table></div>", vbTextCompare) <> 0 Then
      Limits = 1
    End If
    iIndexSt = iIndex
  Loop Until Limits = 1
End Sub
