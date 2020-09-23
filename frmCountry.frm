VERSION 5.00
Begin VB.Form frmCountry 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "Play Animation"
      Height          =   375
      Left            =   1320
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   1680
   End
   Begin VB.Image picSource 
      Height          =   1095
      Left            =   960
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgLgCountry 
      Height          =   1080
      Left            =   45
      MousePointer    =   99  'Custom
      ToolTipText     =   "Right Click To Close"
      Top             =   45
      Visible         =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "frmCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PicRatio As Single
Dim BoxWidth As Integer
Dim BoxHeight As Integer
Dim zooming As Boolean
Dim MinLeft As Integer
Dim MinTop As Integer
Dim MaxLeft As Integer
Dim MaxTop As Integer

Private Sub cmdAnimate_Click()
   If cmdAnimate.Caption = "Exit" Then
      Unload Me
   Else
      Animation = True
      Timer1.Enabled = True
   End If
End Sub

Private Sub Form_Load()
  Set cmdAnimate.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
  frmCountry.Caption = sFrmName ' & " Of " & scntName
  SizePic picTureName
  frmCountry.Show vbModal
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmCountry = Nothing
End Sub

Private Sub imgLgCountry_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Timer1.Enabled = True
  End If
End Sub

Private Sub SizePic(picName As String)
  'load first pic in in picSource, get ratio,
  'size imgLgCountry, send size, empty box1
  'On Error Resume Next
  picSource.Picture = LoadPicture(picName)
  PicRatio = picSource.Width / picSource.Height
  
  If PicRatio > 1.33 Then 'pic is landscape
    BoxWidth = Screen.Width / 2.3
    BoxHeight = (Screen.Width / PicRatio) / 2.3
  End If

  If PicRatio < 1.33 Then
    BoxHeight = Screen.Height / 2.3 'pic is portrait
    BoxWidth = (Screen.Height * PicRatio) / 2.3
  End If

  If PicRatio = 1.33 Then 'pic is square
    BoxHeight = Screen.Height / 2.3
    BoxWidth = Screen.Width / 2.3
  End If
  
  Call ShowPic(BoxWidth, BoxHeight, picName)
  imgLgCountry.Visible = True
End Sub

Public Sub ShowPic(BoxWidth As Integer, BoxHeight As Integer, picName As String)
  'empty box2, size box2,load imgLgCountry from  box1
  imgLgCountry.Visible = False
  imgLgCountry.Height = BoxHeight
  imgLgCountry.Width = BoxWidth
  imgLgCountry.Picture = LoadPicture(picName)
  picSource.Picture = LoadPicture()
  imgLgCountry.Top = 0
  imgLgCountry.Left = 5
  frmCountry.Height = imgLgCountry.Height + 360
  frmCountry.Width = imgLgCountry.Width + 110
  frmCountry.Top = frmWeatherMain.Top + (frmWeatherMain.Height / 2) - (frmCountry.Height / 2)
  If frmCountry.Top < 0 Then
    frmCountry.Top = frmWeatherMain.Top
  End If
  frmCountry.Left = frmWeatherMain.Left + (frmWeatherMain.Width / 2) - (frmCountry.Width / 2)
  If PlayAnimation Then
    cmdAnimate.Caption = "Play Animation"
    cmdAnimate.Visible = True
  End If
  cmdAnimate.Top = frmCountry.Height - 850
  cmdAnimate.Left = (frmCountry.Width / 2) - (cmdAnimate.Width / 2)
End Sub

Public Sub ZoomPicture(BoxWidth As Integer, BoxHeight As Integer)
   Dim X, Y As Single
   On Error Resume Next
  
   X = BoxWidth
   Y = BoxHeight
   X = X / 1.0851
   Y = Y / 1.0851
   BoxWidth = X
   BoxHeight = Y
    
   Call ShowZoom(BoxWidth, BoxHeight, picTureName)
   'Center picture
   frmCountry.Left = frmWeatherMain.Left + (frmWeatherMain.Width / 2) - (frmCountry.Width / 2)
   frmCountry.Top = frmWeatherMain.Top + (frmWeatherMain.Height / 2) - (frmCountry.Height / 2)
   imgLgCountry.Visible = True
   If Y < 425 Or X < 200 Then
      Timer1.Enabled = False
      Unload Me
   End If
End Sub

Private Sub Timer1_Timer()
   Call ZoomPicture(imgLgCountry.Width, imgLgCountry.Height)
End Sub

Private Sub ShowZoom(BoxWidth As Integer, BoxHeight As Integer, picName As String)
   imgLgCountry.Height = BoxHeight
   imgLgCountry.Width = BoxWidth
   frmCountry.Height = imgLgCountry.Height + 300
   frmCountry.Width = imgLgCountry.Width + 25
End Sub
