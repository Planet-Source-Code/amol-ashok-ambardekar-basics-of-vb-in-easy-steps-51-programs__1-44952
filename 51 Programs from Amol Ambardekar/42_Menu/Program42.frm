VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program42"
   ClientHeight    =   3300
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
      End
      Begin VB.Menu close 
         Caption         =   "&Close"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "C&ut"
      End
      Begin VB.Menu copy 
         Caption         =   "C&opy"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()
MsgBox "Type code for Close Here"
End Sub

Private Sub copy_Click()
MsgBox "Type code for Copy Here"
End Sub

Private Sub cut_Click()
MsgBox "Type code for Cut Here"
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopupMenu file
End Sub

Private Sub Form_Unload(Cancel As Integer)
rep = MsgBox("R U Sure ?", vbCritical + vbYesNo, "Quit")
If rep = vbNo Then Cancel = -1
End Sub

Private Sub new_Click()
MsgBox "Type code for New Here"
End Sub

Private Sub paste_Click()
MsgBox "Type code for Paste Here"
End Sub
