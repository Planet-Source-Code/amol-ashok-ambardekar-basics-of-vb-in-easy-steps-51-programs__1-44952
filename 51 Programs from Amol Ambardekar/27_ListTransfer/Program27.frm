VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program27"
   ClientHeight    =   3105
   ClientLeft      =   2790
   ClientTop       =   2985
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   6675
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.ListBox List2 
      Height          =   2790
      ItemData        =   "Program27.frx":0000
      Left            =   3840
      List            =   "Program27.frx":0002
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   2790
      ItemData        =   "Program27.frx":0004
      Left            =   120
      List            =   "Program27.frx":0006
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For i = List1.ListCount - 1 To 0 Step -1
If List1.Selected(i) = True Then
List2.AddItem List1.List(i)
List1.RemoveItem i
End If
Next i
End Sub

Private Sub Form_Activate()
For i = 1 To 25
List1.AddItem i
Next i
End Sub

