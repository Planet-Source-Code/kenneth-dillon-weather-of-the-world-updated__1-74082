Attribute VB_Name = "modGeneral"
Option Explicit
Public sStatusText As String
Public slargeMapLink1 As String
Public slargeMapLink2 As String
Public sfndResult As Integer
Public CountriesArray() As String
Public HolDateSelect As String
Public isTallest As Boolean
Public Nozip As Boolean
Public intMH As Integer 'MaxHeight of imagebox
Public intMW As Integer 'MaxWidth of image box
Public OCX() As Byte
Public bGPS As Boolean
Public sStatState As String
Public sStatArea As String
Public sStatCountry As String
Public sStatRegion As String
Public sStatCounty As String
Public PlayRegAnimation As Boolean
Public PlayAnimation As Boolean
Public AnimationLink As String
Public Animation As Boolean
Public sMapPicture As String
Public sFlagPicture As String
Public picTureName As String
Public scntName As String
Public iMinCount As Integer
Public sFrmName As String
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const CB_FINDSTRINGEXACT = &H158
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Boolean
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_THICKFRAME = &H40000
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long
        'Constants
'Const LB_FINDSTRINGEXACT = &H1A2    'To locate exact match

'Declares
Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Public Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As String) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Sub WAIT(ByVal lMilliSec As Long)
    WaitForSingleObject GetCurrentProcess, lMilliSec
End Sub

Public Function FindStringinListControl(ListControl As Object, _
  ByVal SearchText As String) As Long

  '**************************************
  'Input:
  'ListControl: List or ComboBox Object
  'SearchText: String to Search For

  'Returns: ListIndex of Item if found
  'or -1 if not found
  '***************************************
  
  Dim lHwnd As Long
  Dim lMsg As Long

  'On Error Resume Next
  lHwnd = ListControl.hwnd

  If TypeOf ListControl Is ListBox Then
    lMsg = LB_FINDSTRINGEXACT
  ElseIf TypeOf ListControl Is ComboBox Then
    lMsg = CB_FINDSTRINGEXACT
  Else
    FindStringinListControl = -1
    Exit Function
  End If
  FindStringinListControl = SendMessageAsString(lHwnd, lMsg, -1, SearchText)
End Function

Public Function FileExists(FileName As String) As Boolean
  FileExists = (Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) <> "")
End Function

Public Function SystemDirectory() As String
  Dim RSTR As String
  Dim RLEN As Long

  RSTR = String(255, 0)
  RLEN = GetSystemDirectory(RSTR, Len(RSTR))
  If RLEN < Len(RSTR) Then
    RSTR = Left(RSTR, RLEN)
    If Right(RSTR, 1) = "\" Then
      SystemDirectory = Left(RSTR, Len(RSTR) - 1)
    Else
      SystemDirectory = RSTR
    End If
  Else
    SystemDirectory = ""
  End If
End Function

Public Sub SetPictureBox(NameOfForm As Form, picTureName As String, imgIndex As Integer)
  intMH = NameOfForm.picMain.ScaleHeight '- 20 'in pixels
  intMW = NameOfForm.picMain.ScaleWidth '- 20 'in pixels
  'send it to the LoadAnImage Sub
  Call LoadAnImage(NameOfForm, picTureName, imgIndex, NameOfForm.picMain, NameOfForm.picHidden, NameOfForm.imgPicture, intMH, intMW)
End Sub

Public Sub LoadAnImage(picFormName As Form, picName As String, PicIndex As Integer, picBoxMain As PictureBox, picBoxHidden As PictureBox, imageBoxDisplay As Image, ByVal MaxHeight As Integer, ByVal maxWidth As Integer)
  On Error GoTo LoadAnImageErr
  Dim HighRatio As Single
  Dim WideRatio As Single
  Dim intMaxPicHeight As Integer
  Dim intMaxPicWidth As Integer
  Dim intActualHeight As Integer
  Dim intActualWidth As Integer
  
  intActualHeight = 0
  intActualWidth = 0
  intMaxPicHeight = 0
  intMaxPicWidth = 0
  HighRatio = 0
  WideRatio = 0
  'Get the picture name to load into the hidden picturebox
  'set up the max dimension variables
  intMaxPicHeight = MaxHeight
  intMaxPicWidth = maxWidth
 
  'Load the picture into the hidden picturebox
  'with its Autosize set to True.
  If Len(Trim(picName)) <> 0 Then
    picBoxHidden.Picture = LoadPicture(picName)
    picBoxMain.Visible = False
  Else
    Set picBoxHidden.Picture = picFormName.ImageList2.ListImages(PicIndex).Picture
  End If
  'Get the pic size in pixels
  intActualHeight = CInt(picBoxHidden.ScaleHeight)
  intActualWidth = CInt(picBoxHidden.ScaleWidth)
  'Form a ratio of original height to width - eg 800 x 600 pixels image
  WideRatio = picBoxHidden.ScaleHeight / picBoxHidden.ScaleWidth '600/800
  HighRatio = picBoxHidden.ScaleWidth / picBoxHidden.ScaleHeight '800/600
  'Make the image box invisible until the image is loaded
  imageBoxDisplay.Visible = False
  'Check for Portrait or Landscape image
  If intActualHeight >= intActualWidth Then
    'must be higher than wide - ie portrait
    'Check for smaller image than max allows
    If intActualHeight <= intMaxPicHeight Then
      imageBoxDisplay.Height = intActualHeight
      imageBoxDisplay.Width = intActualWidth
    Else
      imageBoxDisplay.Height = intMaxPicHeight
      imageBoxDisplay.Width = intMaxPicHeight * HighRatio
    End If
  Else
    'must be wider than high - ie landscape
    If intActualWidth <= intMaxPicWidth Then
      imageBoxDisplay.Width = intActualWidth
      imageBoxDisplay.Height = intActualHeight
    Else
      imageBoxDisplay.Width = intMaxPicWidth
      imageBoxDisplay.Height = intMaxPicWidth * WideRatio
      'again make sure the height is not more than the max allows
      If imageBoxDisplay.Height > intMaxPicHeight Then
        'Resize it
        imageBoxDisplay.Height = intMaxPicHeight
        imageBoxDisplay.Width = intMaxPicHeight * HighRatio
      End If
    End If
  End If
  'Center the image within its container picturebox.
  'Load the graphic into the image control.
  imageBoxDisplay.Picture = picBoxHidden.Picture
  If Len(Trim(picName)) <> 0 Then
    imageBoxDisplay.Left = (picFormName.fmMap.Width / 2) - (imageBoxDisplay.Width / 2)
    imageBoxDisplay.Top = (picFormName.fmMap.Height / 2) - (imageBoxDisplay.Height / 2) + 100
  End If
  'Show the image.
  imageBoxDisplay.Visible = True
ExitLoadAnImage:
  Exit Sub
LoadAnImageErr:
  MsgBox Err.Description
  Resume ExitLoadAnImage
End Sub

