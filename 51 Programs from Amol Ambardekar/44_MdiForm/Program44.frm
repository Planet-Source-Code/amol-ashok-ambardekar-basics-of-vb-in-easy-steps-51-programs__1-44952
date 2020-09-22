VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program44"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2415
   ScaleWidth      =   3900
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
If Forms.Count <= 2 Then MDIForm1.mnu_close.Enabled = False
End Sub
