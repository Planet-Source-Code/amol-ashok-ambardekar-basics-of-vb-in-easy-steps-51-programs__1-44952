VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program17"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "/100"
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
      TabIndex        =   15
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "/100"
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
      TabIndex        =   14
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "/100"
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
      TabIndex        =   13
      Top             =   240
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Grade"
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
      Left            =   2760
      TabIndex        =   11
      Top             =   3240
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Percentage"
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
      Left            =   2760
      TabIndex        =   9
      Top             =   2640
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total"
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
      Left            =   2760
      TabIndex        =   7
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "English"
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
      TabIndex        =   4
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "History"
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
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Maths"
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
      Width           =   660
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text4 = Val(Text1) + Val(Text2) + Val(Text3)
Text5 = Val(Text4) / 3
If Val(Text5) < 50 Then Text6 = "D"
If Val(Text5) >= 50 And Val(Text5) < 60 Then Text6 = "C"
If Val(Text5) >= 60 And Val(Text5) < 70 Then Text6 = "B"
If Val(Text5) >= 70 And Val(Text5) < 80 Then Text6 = "A"
End Sub
