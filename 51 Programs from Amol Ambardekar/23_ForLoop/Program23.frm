VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program23"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3000
      TabIndex        =   1
      Text            =   "Size"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Text            =   "Font Name"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amol"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Label1.FontName = Combo1
End Sub

Private Sub Combo2_Click()
Label1.FontSize = Combo2
End Sub

Private Sub Form_Activate()
For i = 0 To Screen.FontCount - 1
Combo1.AddItem Screen.Fonts(i)
DoEvents 'MultiTasking
Next i
Combo1.Text = "Ms Sans Serif"
For i = 8 To 72 Step 2
Combo2.AddItem i
Next i
Combo2.Text = "8"
End Sub
