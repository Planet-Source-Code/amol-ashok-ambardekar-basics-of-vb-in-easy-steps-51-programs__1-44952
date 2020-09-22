VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program7"
   ClientHeight    =   3255
   ClientLeft      =   3810
   ClientTop       =   2805
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4890
   Begin VB.CommandButton Command4 
      Caption         =   "Minimized"
      Height          =   975
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   975
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Maximized"
      Height          =   975
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Normal"
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.WindowState = vbNormal
End Sub

Private Sub Command2_Click()
Form1.WindowState = vbMaximized
End Sub

Private Sub Command3_Click()
Unload Me
'End or Unload Me <<Use one of them to end the program
End Sub

Private Sub Command4_Click()
Form1.WindowState = vbMinimized
End Sub
