VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program51"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   3210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Show Report"
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataReport1.Show
End Sub
