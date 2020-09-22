VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program32"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   2160
      Max             =   1000
      Min             =   10
      SmallChange     =   10
      TabIndex        =   1
      Top             =   960
      Value           =   10
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AMOL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Label1.Left = -Label1.Width
End Sub

Private Sub HScroll1_Change()
Timer1.Interval = HScroll1
End Sub

Private Sub HScroll1_Scroll()
Timer1.Interval = HScroll1
End Sub

Private Sub Timer1_Timer()
Label1.Left = Label1.Left + 60
If Label1.Left > Form1.Width Then Label1.Left = -Label1.Width
End Sub
