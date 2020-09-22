VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program19"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Hide"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1680
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1680
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'To get the controls having container as frame, Draw them manally on the frame.
'To ensure that these controls have container as Frame, try to drag them outside
'the frame during design time. Failing to which shows successful drawing of controls
'on the frame so that they have container as frame
Private Sub Command1_Click()
Frame1.Visible = True
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
End Sub
