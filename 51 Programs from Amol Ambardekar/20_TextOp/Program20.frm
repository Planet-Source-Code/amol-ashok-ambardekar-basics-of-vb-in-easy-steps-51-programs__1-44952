VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program20"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Text            =   "Program20.frx":0000
      Top             =   240
      Width           =   6255
   End
   Begin VB.Frame Frame3 
      Caption         =   "Alignment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4560
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
      Begin VB.OptionButton Option7 
         Caption         =   "Center"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   2400
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
      Begin VB.OptionButton Option4 
         Caption         =   "Black"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Effects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   1935
      Begin VB.CheckBox Check4 
         Caption         =   "Stirkethru"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Underline"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Italic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Text1.FontBold = Not Text1.FontBold
End Sub

Private Sub Check2_Click()
Text1.FontItalic = Not Text1.FontItalic
End Sub

Private Sub Check3_Click()
Text1.FontUnderline = Not Text1.FontUnderline
End Sub

Private Sub Check4_Click()
Text1.FontStrikethru = Not Text1.FontStrikethru
End Sub

Private Sub Option1_Click()
Text1.ForeColor = vbRed
End Sub

Private Sub Option2_Click()
Text1.ForeColor = vbGreen
End Sub

Private Sub Option3_Click()
Text1.ForeColor = vbBlue
End Sub

Private Sub Option4_Click()
Text1.ForeColor = vbBlack
End Sub

Private Sub Option5_Click()
Text1.Alignment = vbLeftJustify
End Sub

Private Sub Option6_Click()
Text1.Alignment = vbRightJustify
End Sub

Private Sub Option7_Click()
Text1.Alignment = vbCenter
End Sub
