VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program47"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   2940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "BMP Wallpaper"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shut Down"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long

Private Sub Command1_Click()
ExitWindowsEx 1, 1
End Sub

Private Sub Command2_Click()
PaintDesktop Form1.hdc
End Sub
