VERSION 5.00
Begin VB.UserControl RegionButton 
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   Picture         =   "RegionBtn.ctx":0000
   ScaleHeight     =   1335
   ScaleWidth      =   1335
   ToolboxBitmap   =   "RegionBtn.ctx":06A9
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   1560
      LinkTimeout     =   100
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Timer tmrAnimator 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   150
      Top             =   1740
   End
   Begin VB.PictureBox picFormSkin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "RegionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const RGN_OR = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

' High level sound support API
   Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias _
      "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As _
      Long) As Long
   Const SND_SYNC = &H0
   Const SND_ASYNC = &H1
   Const SND_NODEFAULT = &H2
   Const SND_LOOP = &H8
   Const SND_NOSTOP = &H10

Dim windowRegion As Long
Dim FromRegionData() As Byte
Dim ToRegionData() As Byte
Dim framePic As Integer
Dim picWidth As Long, picHeight  As Long
Dim numImages As Long

Dim CurrentScaleWidth As Long
Dim CurrentScaleHeigth As Long
Dim TwipSacaleWeith As Long
Dim TwipSacaleHeight As Long

Dim GetRegKeyM As String
Dim pic As StdPicture
Dim XLoop As Long
Dim YLoop As Long
Dim RowFram As Long
Dim ColFram As Long
Dim CurntScalceX As Long
Dim CurntScalceY As Long
Dim AniInfoFile As String
Dim StartFrame As Long
Dim EndFrame As Long
Dim AniLoopAble As Boolean
Dim AnimationSpeed As Long
Dim AnimationVoice As String
Dim i As Long
Dim J As Long
Public AnimationEndOfPlay As Boolean
Private RegisterCode As Long
Event Click() 'MappingInfo=picFormSkin,picFormSkin,-1,Click
Event DblClick() 'MappingInfo=picFormSkin,picFormSkin,-1,DblClick
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picFormSkin,picFormSkin,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picFormSkin,picFormSkin,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Private Sub SetSizeLoc()
CurrentScaleWidth = Picture1.ScaleX(Picture1.Width / RowFram, 1, 3)
CurrentScaleHeigth = Picture1.ScaleY(Picture1.Height / ColFram, 1, 3)
TwipSacaleWeith = Picture1.Width / RowFram
TwipSacaleHeight = Picture1.Height / ColFram
' set the window metrics equal to the picturebox metrics
UserControl.Width = TwipSacaleWeith
UserControl.Height = TwipSacaleHeight
picWidth = CurrentScaleWidth
picHeight = CurrentScaleHeigth
picFormSkin.Width = TwipSacaleWeith
picFormSkin.Height = TwipSacaleHeight
picFormSkin.Move 0, 0
End Sub

Private Sub Setting()
  Dim arraySize As Long
  ' set the array dimensions
  arraySize = picWidth * picHeight
  ReDim FromRegionData(1 To numImages, 0 To arraySize)
  For i = 1 To numImages
     CreateRegionDataFileFromPic i, picFormSkin
  Next i
  XLoop = 0
  YLoop = 0
  framePic = 1
 
End Sub

' creates a window region data file which contains transparent and opaque pixel data
Private Sub CreateRegionDataFileFromPic(ByVal index As Integer, ByRef picSkin As PictureBox)
  On Error GoTo errHandler
  Dim X As Long, Y As Long, TransparentColor As Long
  Dim val As Byte
  picFormSkin.PaintPicture Picture1.Picture, 0, 0, TwipSacaleWeith, TwipSacaleHeight, XLoop, YLoop, TwipSacaleWeith, _
  TwipSacaleHeight
  
  If XLoop + TwipSacaleWeith < Picture1.Width Then
    XLoop = XLoop + TwipSacaleWeith
    Else
    XLoop = 0
  If YLoop + TwipSacaleHeight < Picture1.Height And ColFram <> 1 Then
    YLoop = YLoop + TwipSacaleHeight
    Else
    XLoop = 0
    YLoop = 0
  End If
  End If

  ' get transparent color
  TransparentColor = GetPixel(picFormSkin.hDC, 0, 0)
  ' scan pic in picturebox and create raw window region data file
  For Y = 0 To picHeight - 1
    For X = 0 To picWidth - 1
      If GetPixel(picFormSkin.hDC, X, Y) <> TransparentColor Then
        val = 1
      Else
        val = 0
      End If
         FromRegionData(index, X + (Y * picWidth)) = val
    Next X
  Next Y
 Exit Sub
errHandler:
  If Err.Number = 53 Then Resume Next
End Sub

' creates a buffered copy of the window region from a raw window region data file
Private Sub LoadRegionDataFromFile(ByVal index As Integer)
  On Error GoTo errHandler
  
  Dim X As Long, Y As Long
     
  For Y = 0 To picHeight - 1
    For X = 0 To picWidth - 1
 ToRegionData(index, X + (Y * picWidth)) = FromRegionData(index, X + (Y * picWidth))
    Next X
  Next Y
  
  Exit Sub
  
errHandler:
  MsgBox "Fatal Error: Region Data File not found!", vbCritical Or vbOKOnly, "Application Error"
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleX
Public Function SetSizeX(ByVal Width As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Variant
    SetSizeX = UserControl.ScaleX(Width, FromScale, ToScale)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleY
Public Function SetSizeY(ByVal Height As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
    SetSizeY = UserControl.ScaleY(Height, FromScale, ToScale)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'Private Sub UserControl_Initialize()
'GetRegisterKey "9.97592854978955E+17"
'AnimationSetPicture App.Path + "\run.BMP", App.Path + "\run.ani"
'End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub picFormSkin_Click()
    RaiseEvent Click
End Sub

Private Sub picFormSkin_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picFormSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picFormSkin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Function AnimationSetPicture(PicNamePath As String, AnimateSequencerFilePath As String) As Boolean
Dim PicFilePath As String
Dim AniSecFilePath As String, FilePath As String
Dim Row As Long, Col As Long
  If GetRegKeyM <> CStr(997592854) Then
    frmAbout.lblDescription.Caption = "Please input Register code."
    frmAbout.Show vbModal, UserControl
    Exit Function
  End If
AniSecFilePath = Dir(AnimateSequencerFilePath, vbHidden Or vbReadOnly Or vbSystem)
FilePath = Dir(PicNamePath, vbHidden Or vbReadOnly Or vbSystem)
If FilePath = "" Or AniSecFilePath = "" Then
        MsgBox "The file is not existing.", , "Caution"
Else
       AniInfoFile = AnimateSequencerFilePath
       Row = CLng(GetIniData(AnimateSequencerFilePath, "PictureInfo", "Row"))
       Col = CLng(GetIniData(AnimateSequencerFilePath, "PictureInfo", "Col"))
       Picture1.Picture = LoadPicture(PicNamePath, 0, 0, 0, 0)
       XLoop = 0
       YLoop = 0
    If Row <> 0 And Col <> 0 Then
        RowFram = Row
        ColFram = Col
        numImages = RowFram * ColFram
        SetSizeLoc 'Set Picture reference Scale
        Setting
        AnimationSetPicture = True
    Else
        MsgBox "Row or Col must be bigger than 0.", , "Caution"
    End If
End If
End Function
Private Sub tmrAnimator_Timer()
   windowRegion = CreateFormRegion(framePic)
  SetWindowRgn hWnd, windowRegion, True
  framePic = framePic + 1
If framePic > EndFrame Then
    tmrAnimator.Enabled = False
    AnimationEndOfPlay = True
    framePic = StartFrame
    XLoop = 0
    YLoop = 0
    XLoop = TwipSacaleWeith * (i - 1)
    YLoop = TwipSacaleHeight * (J - 1)
End If
End Sub

'Private Sub picFormSkin_KeyDown(KeyCode As Integer, Shift As Integer)
 '  windowRegion = CreateFormRegion(framePic)
  ' SetWindowRgn hWnd, windowRegion, True
  'framePic = framePic + 1
'If framePic > EndFrame Then
'    tmrAnimator.Enabled = False
 '   AnimationEndOfPlay = True
 '   framePic = StartFrame
 '   XLoop = 0
 '   YLoop = 0
 '   XLoop = TwipSacaleWeith * (i - 1)
 '   YLoop = TwipSacaleHeight * (J - 1)
'End If
'End Sub

Public Function AnimationPlay(AnimateKeyFrame As String, AnimateMediaPath As String) As Boolean
Dim CountingUp As Long
If AniInfoFile = "" Then
    MsgBox "Plaese use ""AnimationSetPicture"" function befor this function.", , "Caution"
    Exit Function
End If
If AnimateKeyFrame <> "" And AnimateKeyFrame <> " " Then
    If FindSection(AniInfoFile, AnimateKeyFrame) = True Then
    'Setup Variable
        StartFrame = 0
        EndFrame = 0
        AnimationSpeed = 0
        AnimationVoice = "False"
    'Get animation info
        StartFrame = CLng(GetIniData(AniInfoFile, AnimateKeyFrame, "StartFrame"))
        EndFrame = CLng(GetIniData(AniInfoFile, AnimateKeyFrame, "EndFrame"))
        AnimationSpeed = CLng(GetIniData(AniInfoFile, AnimateKeyFrame, "AnimationSpeed"))
        AnimationVoice = GetIniData(AniInfoFile, AnimateKeyFrame, "AnimationVoice")
    Else
        MsgBox "Frame is'nt exist or has wrong name.", , "Caution"
        Exit Function
    End If
 
 For J = 1 To ColFram
 For i = 1 To RowFram
       CountingUp = CountingUp + 1
    If StartFrame = CountingUp Then
       Exit For
    End If
 Next i
    If StartFrame = CountingUp Then
        Exit For
        End If
 Next J
'Set All Variable For Doing Animation loop.
    XLoop = 0
    YLoop = 0
    framePic = StartFrame
    XLoop = TwipSacaleWeith * (i - 1)
    YLoop = TwipSacaleHeight * (J - 1)
        If XLoop + TwipSacaleWeith > Picture1.Width Then
            XLoop = 0
        If YLoop + TwipSacaleHeight > Picture1.Height And ColFram <> 1 Then
            XLoop = 0
            YLoop = 0
        End If
        End If

        If AnimationVoice <> "False" And AnimationVoice <> "" And AnimationVoice <> " " Then
        AnimateMediaPath = AnimateMediaPath & "\"
            sndPlaySound AnimateMediaPath & AnimationVoice, SND_ASYNC Or SND_NODEFAULT
        End If
    'Start Animation loop.
    tmrAnimator.Interval = AnimationSpeed
    AnimationPlay = True
    AnimationEndOfPlay = False
    tmrAnimator.Enabled = True
    
End If
End Function

' creates an irregular window region based from the buffered region data
Private Function CreateFormRegion(ByVal index As Integer) As Long
  Dim X As Long, Y As Long, StartX As Long
  Dim FullRegion As Long, CurrentRgn As Long, TransparentColor As Long
  Dim NoRegionYet As Boolean
  
NoRegionYet = True
'Copy Picture Frame
    picFormSkin.PaintPicture Picture1.Picture, 0, 0, TwipSacaleWeith, TwipSacaleHeight, XLoop, YLoop, TwipSacaleWeith, _
       TwipSacaleHeight
    If XLoop + TwipSacaleWeith + 1 < Picture1.Width Then
        XLoop = XLoop + TwipSacaleWeith
    Else
        XLoop = 0
    If YLoop + TwipSacaleHeight + 1 < Picture1.Height And ColFram <> 1 Then
        YLoop = YLoop + TwipSacaleHeight
    Else
        XLoop = 0
        YLoop = 0
    End If
    End If
'MsgBox framePic
  ' scan image for transparent pixels
  For Y = 0 To picHeight - 1
    For X = 0 To picWidth - 1
      StartX = X
      While FromRegionData(index, X + (Y * picWidth)) = 1 And X < picWidth
        X = X + 1
      Wend
      If StartX <> X Then
        CurrentRgn = CreateRectRgn(StartX, Y, X, Y + 1)
        If NoRegionYet Then
          FullRegion = CurrentRgn
          NoRegionYet = False
        Else
          CombineRgn FullRegion, FullRegion, CurrentRgn, RGN_OR
          DeleteObject CurrentRgn
        End If
      End If
    Next X
  Next Y
    CreateFormRegion = FullRegion
picFormSkin.Visible = True
End Function

Public Sub About()
frmAbout.Show vbModal, UserControl
End Sub

Public Function GetRegisterKey(Number As Long) As Boolean
GetRegKeyM = CStr(Number)
End Function


