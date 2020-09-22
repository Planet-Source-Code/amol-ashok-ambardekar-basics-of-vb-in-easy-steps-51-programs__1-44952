VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program1"
   ClientHeight    =   3105
   ClientLeft      =   3060
   ClientTop       =   2805
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   5475
   Begin VB.CommandButton Command1 
      Caption         =   "Hit Me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "Welcome to Visual Basic Programmig"
End Sub
