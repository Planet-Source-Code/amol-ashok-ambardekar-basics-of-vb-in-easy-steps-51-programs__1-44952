VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Program39"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   2010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Documents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   480
      MouseIcon       =   "Program39.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2880
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Players"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   480
      MouseIcon       =   "Program39.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2040
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Utilities"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   480
      MouseIcon       =   "Program39.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Programs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   480
      MouseIcon       =   "Program39.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   360
      Width           =   1020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer 'To store the index of the control visited recently
Dim j As Integer 'To avoid flickering when mouse moves on the same control
Private Sub Form_Activate()
j = 4 'the max index of control + 1
i = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
j = 4
Label1(i).FontItalic = False
Label1(i).FontUnderline = False
Label1(i).ForeColor = vbBlack
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If j <> Index Then
Label1(i).FontItalic = False
Label1(i).FontUnderline = False
Label1(i).ForeColor = vbBlack
Label1(Index).FontItalic = True
Label1(Index).FontUnderline = True
Label1(Index).ForeColor = vbBlue
i = Index
j = Index
End If
End Sub
