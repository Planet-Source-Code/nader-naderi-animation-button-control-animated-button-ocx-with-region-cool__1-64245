VERSION 5.00
Object = "*\ARegionBtn OCX.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5790
   Icon            =   "Test2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin project1.RegionButton RegionButton1 
      Height          =   975
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Spin"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4260
      TabIndex        =   6
      Top             =   795
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Blub"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4245
      TabIndex        =   5
      Top             =   1395
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load Spin"
      Height          =   495
      Left            =   2925
      TabIndex        =   4
      Top             =   795
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load Blob"
      Height          =   495
      Left            =   2925
      TabIndex        =   3
      Top             =   1395
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2415
      Top             =   2040
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load Run"
      Height          =   510
      Left            =   2925
      TabIndex        =   2
      Top             =   2025
      Width           =   1170
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Run"
      Enabled         =   0   'False
      Height          =   480
      Left            =   4245
      TabIndex        =   1
      Top             =   2070
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Stop Loop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2955
      TabIndex        =   0
      Top             =   135
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click on OCX control"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Toggle As Boolean
Private Sub Command1_Click()
 RegionButton1.Visible = True
 RegionButton1.AnimationPlay "Twist", App.Path
End Sub
Private Sub Command2_Click()
RegionButton1.Visible = True
RegionButton1.AnimationPlay "Twist", App.Path
End Sub

Private Sub Command3_Click()
RegionButton1.AnimationSetPicture App.Path + "\DONUT.BMP", App.Path + "\donut.inf"
ActiveBtn
Command1.Enabled = True
Command6.Enabled = False
Command2.Enabled = False
RegionButton1.Visible = True
RegionButton1.AnimationPlay "Twist", App.Path
End Sub

Private Sub Command4_Click()
RegionButton1.AnimationSetPicture App.Path + "\Blub.BMP", App.Path + "\blub.inf"
ActiveBtn
Command1.Enabled = False
Command6.Enabled = False
Command2.Enabled = True
RegionButton1.Visible = True
RegionButton1.AnimationPlay "Twist", App.Path
End Sub

Private Sub Command5_Click()
RegionButton1.AnimationSetPicture App.Path + "\run.BMP", App.Path + "\run.inf"
ActiveBtn
Command1.Enabled = False
Command6.Enabled = True
Command2.Enabled = False
RegionButton1.Visible = True
RegionButton1.AnimationPlay "Twist", App.Path
End Sub

Private Sub Command6_Click()
RegionButton1.AnimationPlay "Twist", App.Path
End Sub

Private Sub Command7_Click()
If Toggle = False Then
    RegionButton1.AnimationEndOfPlay = True
    Toggle = True
    Command7.Caption = "Start Loop"
    Timer1.Enabled = True
    Else
    RegionButton1.AnimationEndOfPlay = False
    Toggle = False
    Command7.Caption = "Stop Loop"
    Timer1.Enabled = False
End If
End Sub

Private Sub Form_Load()
RegionButton1.GetRegisterKey 997592854
'RegionButton1.AnimationSetPicture App.Path + "\Blub.BMP", App.Path + "\blub.inf"
'RegionButton1.AnimationSetPicture App.Path + "\run.BMP", App.Path + "\run.inf"
RegionButton1.AnimationSetPicture App.Path + "\DONUT.BMP", App.Path + "\donut.inf"
RegionButton1.Visible = True
RegionButton1.AnimationPlay "Twist", App.Path
End Sub

Private Sub RegionButton1_Click()
RegionButton1.Visible = True
RegionButton1.AnimationPlay "Twist", App.Path
End Sub

Private Sub Timer1_Timer()
If RegionButton1.AnimationEndOfPlay = True Then
RegionButton1.AnimationEndOfPlay = False
RegionButton1.AnimationPlay "Twist", App.Path
If RegionButton1.Visible = False Then RegionButton1.Visible = True
End If
End Sub


Private Sub ActiveBtn()
Command7.Enabled = True
End Sub
