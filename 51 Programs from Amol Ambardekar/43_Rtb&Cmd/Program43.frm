VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Program43"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5953
      _Version        =   393217
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Program43.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cd.Flags = cdlCFBoth
cd.ShowFont
rtb.SelFontName = cd.FontName
rtb.SelFontSize = cd.FontSize
rtb.SelBold = cd.FontBold
rtb.SelItalic = cd.FontItalic
rtb.SelUnderline = cd.FontUnderline
rtb.SelStrikeThru = cd.FontStrikethru
End Sub

Private Sub Command2_Click()
cd.Flags = 0
cd.ShowColor
rtb.SelColor = cd.Color
End Sub

Private Sub Command3_Click()
cd.Filter = "RTF Files *.rtf|*.rtf"
cd.ShowSave
rtb.SaveFile cd.FileName
End Sub

Private Sub Command4_Click()
cd.Filter = "RTF Files *.rtf|*.rtf"
cd.ShowOpen
rtb.FileName = cd.FileName
End Sub

Private Sub Command5_Click()
cd.Flags = cdlPDReturnDC + cdlPDNoPageNums
If rtb.SelLength = 0 Then
cd.Flags = cd.Flags + cdlPDAllPages
Else
cd.Flags = cd.Flags + cdlPDSelection
End If
On Error GoTo ErrorHandler
cd.ShowPrinter
rtb.SelPrint cd.hDC
ErrorHandler: Exit Sub

End Sub
