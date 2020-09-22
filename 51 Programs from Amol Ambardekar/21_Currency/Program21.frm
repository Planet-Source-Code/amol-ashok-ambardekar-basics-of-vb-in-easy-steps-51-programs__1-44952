VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program21"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   2940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Amount"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check5 
      Caption         =   "10"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "20"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "50"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "100"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "500"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Text1.Visible = Not Text1.Visible
End Sub

Private Sub Check2_Click()
Text2.Visible = Not Text2.Visible
End Sub

Private Sub Check3_Click()
Text3.Visible = Not Text3.Visible
End Sub

Private Sub Check4_Click()
Text4.Visible = Not Text4.Visible
End Sub

Private Sub Check5_Click()
Text5.Visible = Not Text5.Visible
End Sub

Private Sub Command1_Click()
Text6.Visible = True
amount = 0
If Check1.Value = vbChecked Then amount = amount + 500 * Val(Text1)
If Check2.Value = vbChecked Then amount = amount + 100 * Val(Text2)
If Check3.Value = vbChecked Then amount = amount + 50 * Val(Text3)
If Check4.Value = vbChecked Then amount = amount + 20 * Val(Text4)
If Check5.Value = vbChecked Then amount = amount + 10 * Val(Text5)
Text6 = amount
End Sub
