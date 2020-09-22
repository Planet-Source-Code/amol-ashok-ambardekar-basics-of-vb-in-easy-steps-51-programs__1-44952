VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program13"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Number Addition"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "String Addition"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Text            =   "45"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "12"
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text3.Text = Text1.Text + Text2.Text
End Sub

Private Sub Command2_Click()
Text3.Text = Val(Text1.Text) + Val(Text2.Text)
End Sub
