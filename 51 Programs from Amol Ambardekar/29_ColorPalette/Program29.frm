VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program29"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   855
      Left            =   2040
      TabIndex        =   9
      Top             =   2160
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   375
      LargeChange     =   5
      Left            =   1680
      Max             =   255
      TabIndex        =   8
      Top             =   1440
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      LargeChange     =   5
      Left            =   1680
      Max             =   255
      TabIndex        =   7
      Top             =   840
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   5
      Left            =   1680
      Max             =   255
      TabIndex        =   6
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Left            =   4920
      TabIndex        =   5
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Left            =   4920
      TabIndex        =   4
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Left            =   4920
      TabIndex        =   3
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "BLUE"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "GREEN"
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
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "RED"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UpdateColor is a custom private sub written by the programmer
'This is analogous to procedure in pascal and functions returning void in C/C++
'Example of how program can be divided on the basis of small units like functions and procedures(SubRoutines in VB)
Private Sub UpdateColor()
Text1.BackColor = RGB(HScroll1, HScroll2, HScroll3)
End Sub
Private Sub HScroll1_Change()
Label4.Caption = HScroll1
UpdateColor
End Sub

Private Sub HScroll1_Scroll()
Label4.Caption = HScroll1
UpdateColor
End Sub

Private Sub HScroll2_Change()
Label5.Caption = HScroll2
UpdateColor
End Sub

Private Sub HScroll2_Scroll()
Label5.Caption = HScroll2
UpdateColor
End Sub

Private Sub HScroll3_Change()
Label6.Caption = HScroll3
UpdateColor
End Sub

Private Sub HScroll3_Scroll()
Label6.Caption = HScroll3
UpdateColor
End Sub

