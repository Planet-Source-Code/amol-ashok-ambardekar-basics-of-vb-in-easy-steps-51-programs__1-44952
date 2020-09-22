VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program3"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Hide"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command Button Under Test"
      Height          =   1095
      Left            =   2520
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
MsgBox "I am Clicked"
End Sub

Private Sub Command2_Click()
Command1.Visible = True
End Sub

Private Sub Command3_Click()
Command1.Visible = False
End Sub
