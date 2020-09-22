VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program35"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   1330
      X2              =   2160
      Y1              =   1470
      Y2              =   960
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "9"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "6"
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
      Left            =   1290
      TabIndex        =   2
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "3"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "12"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   270
   End
   Begin VB.Shape Shape1 
      Height          =   1995
      Left            =   360
      Shape           =   3  'Circle
      Top             =   480
      Width           =   1995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Double
Private Sub Form_Load()
i = 3.14
End Sub

Private Sub Timer1_Timer()
Line1.Visible = True
i = i - 0.104666666666667
Line1.X2 = 1000 * Sin(i) + Line1.X1
Line1.Y2 = 1000 * Cos(i) + Line1.Y1
End Sub
