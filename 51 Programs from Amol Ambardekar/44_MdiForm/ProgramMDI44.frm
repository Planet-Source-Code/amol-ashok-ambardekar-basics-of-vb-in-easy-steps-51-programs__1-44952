VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Program44"
   ClientHeight    =   3555
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6690
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnu_new 
      Caption         =   "&New"
   End
   Begin VB.Menu mnu_close 
      Caption         =   "&Close"
   End
   Begin VB.Menu mnu_window 
      Caption         =   "&Window"
      Begin VB.Menu mnu_cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnu_tile 
         Caption         =   "&Tile"
         Begin VB.Menu mnu_hori 
            Caption         =   "&Horizontal"
         End
         Begin VB.Menu mnu_vert 
            Caption         =   "&Vertical"
         End
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub mnu_cascade_Click()
MDIForm1.Arrange vbCascade
End Sub

Private Sub mnu_close_Click()
Unload MDIForm1.ActiveForm
If Forms.Count < 2 Then
mnu_close.Enabled = False
i = 0
End If
End Sub

Private Sub mnu_hori_Click()
MDIForm1.Arrange vbTileHorizontal
End Sub

Private Sub MDIForm_Activate()
Form1.Caption = "Doument 1"
i = 1
End Sub

Private Sub mnu_new_Click()
mnu_close.Enabled = True
i = i + 1
Dim x As New Form1
Load x
x.Show
x.Caption = "Document" & i
End Sub

Private Sub mnu_vert_Click()
MDIForm1.Arrange vbTileVertical
End Sub
