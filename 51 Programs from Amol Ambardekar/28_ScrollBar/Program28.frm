VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program28"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   4
      Left            =   240
      Max             =   72
      Min             =   8
      SmallChange     =   2
      TabIndex        =   0
      Top             =   2280
      Value           =   8
      Width           =   5415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amol"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   390
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HScroll1_Change()
Label1.FontSize = HScroll1
End Sub

Private Sub HScroll1_Scroll()
Label1.FontSize = HScroll1.Value 'As Hscroll1.value=Hscoll1
End Sub
