VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Program50"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Program50.frx":0000
      Height          =   1935
      Left            =   480
      OleObjectBlob   =   "Program50.frx":0014
      TabIndex        =   4
      Top             =   0
      Width           =   4935
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "EmpNo."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\50 Programs\50_DataGrid\Employee.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Main"
      Top             =   2760
      Visible         =   0   'False
      Width           =   5460
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
If Option1.Value = True Then
If IsNumeric(Text1) = False Then Text1 = ""
Data1.RecordSource = "select * from Main where (EmpNo like'" & Trim(Text1) & "*')"
Data1.Refresh
DBGrid1.Refresh
End If
If Option2.Value = True Then
Data1.RecordSource = "select * from Main where (Name like '" & Trim(Text1) & "*')"
Data1.Refresh
DBGrid1.Refresh
End If
If Option3.Value = True Then
If IsNumeric(Text1) = False Then Text1 = ""
Data1.RecordSource = "select * from Main where (Salary like '" & Trim(Text1) & "*')"
Data1.Refresh
DBGrid1.Refresh
End If
End Sub
