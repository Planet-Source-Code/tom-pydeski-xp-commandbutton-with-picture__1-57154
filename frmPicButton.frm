VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PicButton"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   2310
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1680
      Picture         =   "frmPicButton.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "       &Execute"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This API function allows us to change the parent of any component that has a hWnd
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Sub Form_Load()
With picLogo
    SetParent .hWnd, cmdExecute.hWnd
    .Top = 100
    .Left = 100
    .Visible = True
    Refresh
    DoEvents
End With
End Sub
