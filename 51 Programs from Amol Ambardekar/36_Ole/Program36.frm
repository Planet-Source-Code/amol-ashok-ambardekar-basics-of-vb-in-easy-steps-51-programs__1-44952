VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program36"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.OLE OLE1 
      Class           =   "mplayer"
      DisplayType     =   1  'Icon
      Height          =   1095
      Left            =   480
      OleObjectBlob   =   "Program36.frx":0000
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
OLE1.Action = 7
End Sub

Private Sub Command2_Click()
OLE1.Action = 9
End Sub
