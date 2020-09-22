VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program30"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   3375
      LargeChange     =   150
      Left            =   0
      Max             =   3250
      SmallChange     =   15
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   150
      Left            =   240
      Max             =   4500
      SmallChange     =   15
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AMOL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   630
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
'Setting for Initial status
HScroll1.Max = HScroll1.Width - VScroll1.Width
VScroll1.Max = VScroll1.Height - HScroll1.Height
HScroll1.Value = Label1.Left + HScroll1.Height
VScroll1.Value = Label1.Top + VScroll1.Width
End Sub

Private Sub HScroll1_Change()
Label1.Left = HScroll1 + HScroll1.Height
End Sub

Private Sub HScroll1_Scroll()
Label1.Left = HScroll1 + HScroll1.Height
End Sub

Private Sub VScroll1_Change()
Label1.Top = VScroll1 + VScroll1.Width
End Sub

Private Sub VScroll1_Scroll()
Label1.Top = VScroll1 + VScroll1.Width
End Sub
