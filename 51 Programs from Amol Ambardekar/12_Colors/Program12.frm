VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program12"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Palette"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "System"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "RGB"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "QBColors"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VBColors"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.BackColor = vbRed
Text1.Text = "RED"
Text2.BackColor = vbGreen
Text2.Text = "GREEN"
Text3.BackColor = vbBlue
Text3.Text = "BLUE"
End Sub

Private Sub Command2_Click()
Text1.BackColor = QBColor(4)
Text1.Text = "QBColor(4)"
Text2.BackColor = QBColor(2)
Text2.Text = "QBColor(2)"
Text3.BackColor = QBColor(1)
Text3.Text = "QBColor(1)"
End Sub

Private Sub Command3_Click()
Text1.BackColor = RGB(255, 0, 0)
Text1.Text = "RGB(255, 0, 0)"
Text2.BackColor = RGB(0, 255, 0)
Text2.Text = "RGB(0, 255, 0)"
Text3.BackColor = RGB(0, 0, 255)
Text3.Text = "RGB(0, 0, 255)"
End Sub

Private Sub Command5_Click()
Text1.BackColor = &HFF&
Text1.Text = "&HFF&"
Text2.BackColor = &HFF00&
Text2.Text = "&HFF00&"
Text3.BackColor = &HFF0000
Text3.Text = "&HFF0000&"
End Sub

Private Sub Command4_Click()
Text1.BackColor = vbActiveTitleBar
Text1.Text = "ActiveTitleBar"
Text2.BackColor = vbDesktop
Text2.Text = "Desktop"
Text3.BackColor = vbInactiveBorder
Text3.Text = "InactiveBorder"
End Sub
