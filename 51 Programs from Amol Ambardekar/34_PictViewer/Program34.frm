VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program34"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Program34.frx":0000
      Left            =   2520
      List            =   "Program34.frx":0010
      TabIndex        =   0
      Text            =   "*.*"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
File1.Pattern = Combo1
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
If Err.Number = 68 Then
rep = MsgBox("Please Insert Floppy/CD", vbCritical + vbRetryCancel, "ERROR")
If rep = vbRetry Then
Drive1_Change
Else
Drive1.Drive = "C:\"
End If
End If
End Sub

Private Sub File1_Click()
On Error Resume Next
If Right(File1.Path, 1) = "\" Then
Text1 = File1.Path & File1.FileName
Else
Text1 = File1.Path & "\" & File1.FileName
End If
Image1.Picture = LoadPicture(Text1)
If Err.Number = 481 Then MsgBox "Inavalid Picture File", vbExclamation, "Sorry"
End Sub

Private Sub Form_Activate()
Drive1.Drive = "C:\"
Dir1.Path = "C:\"
End Sub

