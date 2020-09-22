VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Program41"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   1680
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   1560
      MouseIcon       =   "Program41.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   445
      TabIndex        =   11
      Top             =   240
      Width           =   6735
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   2
      Left            =   240
      Max             =   10
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Open"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Save"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Color"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Filled Circle"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Circle"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Box"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Rectangle"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Line"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sketch"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   1440
      Y1              =   5760
      Y2              =   5760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim started As Boolean
Dim xst, yst, xold, yold, r1, r2, r, rold, indx As Single

Private Sub Command1_Click()
Picture1.Cls
Picture1.Picture = LoadPicture("")
End Sub

Private Sub Command10_Click()
cd.ShowOpen
Picture1.Picture = LoadPicture(cd.FileName)
End Sub

Private Sub Command2_Click()
indx = 1
End Sub

Private Sub Command3_Click()
indx = 2
End Sub

Private Sub Command4_Click()
indx = 3
Picture1.FillStyle = 1
End Sub

Private Sub Command5_Click()
indx = 3
Picture1.FillStyle = vbSolid
End Sub

Private Sub Command6_Click()
indx = 4
Picture1.FillStyle = 1
End Sub

Private Sub Command7_Click()
indx = 4
Picture1.FillStyle = vbSolid
End Sub

Private Sub Command8_Click()
cd.ShowColor
Picture1.ForeColor = cd.Color
End Sub

Private Sub Command9_Click()
cd.ShowSave
SavePicture Picture1.Image, cd.FileName
End Sub

Private Sub HScroll1_Change()
Line1.BorderWidth = HScroll1
Picture1.DrawWidth = HScroll1
End Sub

Private Sub HScroll1_Scroll()
Line1.BorderWidth = HScroll1
Picture1.DrawWidth = HScroll1
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
started = True
xst = X
yst = Y
xold = X
yold = Y
rold = 0
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If started = True Then
Select Case indx
Case 1
Picture1.Line (xold, yold)-(X, Y), cd.Color
Case 2
Picture1.DrawMode = vbXorPen
Picture1.Line (xst, yst)-(xold, yold), vbWhite
Picture1.Line (xst, yst)-(X, Y), vbWhite
Case 3
Picture1.DrawMode = vbXorPen
Picture1.Line (xst, yst)-(xold, yold), vbWhite, B
Picture1.Line (xst, yst)-(X, Y), vbWhite, B
Case 4
Picture1.DrawMode = vbXorPen
r1 = xst - X
r2 = yst - Y
r = Sqr((r1 * r1) + (r2 * r2))
Picture1.Circle (xst, yst), rold, vbWhite
Picture1.Circle (xst, yst), r, vbWhite
End Select
xold = X: yold = Y
rold = r
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
started = False
Picture1.DrawMode = vbCopyPen
If indx = 2 Then Picture1.Line (xst, yst)-(X, Y), cd.Color
If indx = 3 Then Picture1.Line (xst, yst)-(X, Y), cd.Color, B
If indx = 4 Then Picture1.Circle (xst, yst), r, cd.Color
End Sub
