VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program4"
   ClientHeight    =   3345
   ClientLeft      =   3060
   ClientTop       =   1170
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   6405
   Begin VB.CommandButton Command3 
      Caption         =   "Disable"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enable"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hit Me"
      DisabledPicture =   "Program4.frx":0000
      DownPicture     =   "Program4.frx":0442
      Height          =   1095
      Left            =   2640
      Picture         =   "Program4.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "I am On"
End Sub

Private Sub Command2_Click()
Command1.Enabled = True
End Sub

Private Sub Command3_Click()
Command1.Enabled = False
End Sub
