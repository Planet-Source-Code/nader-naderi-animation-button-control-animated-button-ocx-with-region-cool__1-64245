VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   2835
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4470
      TabIndex        =   1
      Top             =   2070
      Width           =   1260
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   1410
      Left            =   270
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   948.15
      ScaleMode       =   0  'User
      ScaleWidth      =   948.15
      TabIndex        =   0
      Top             =   270
      Width           =   1410
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Written By Nader Naderi 2002 - 2003"
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   990
      TabIndex        =   5
      Top             =   2070
      Width           =   3255
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0"
      Height          =   225
      Left            =   1800
      TabIndex        =   4
      Top             =   810
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   270
      X2              =   5819
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Label lblTitle 
      Caption         =   "Region Button Active X"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1800
      TabIndex        =   3
      Top             =   270
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      Caption         =   "This Program protected by copyright laws."
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1800
      TabIndex        =   2
      Top             =   1155
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   270
      X2              =   5834
      Y1              =   1830
      Y2              =   1830
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub cmdOK_Click()
Unload Me
End Sub

