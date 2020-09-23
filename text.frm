VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form textviewer 
   BackColor       =   &H8000000D&
   Caption         =   "Text Viewer"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8100
   Icon            =   "text.frx":0000
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "text.frx":0442
      Top             =   0
      Width           =   4575
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   1080
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu mnusave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnusaveas 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "textviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dir As String
Private Sub Form_Load()
On Error Resume Next
Dim space As String
If VBA.Right$(Form1.Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & Form1.File1.FileName)
Else
dir = Form1.Dir1.Path & "\" & Form1.File1.FileName
End If
space = Chr$(13) + Chr$(10)
Text1.text = ""
Open dir For Input As #1
Do Until EOF(1)
Line Input #1, row
Text1.text = Text1.text + row & space
Loop
Close #1
textviewer.Caption = textviewer.Caption & " - " & Form1.File1.FileName
End Sub

Private Sub Form_Resize()
Text1.Height = textviewer.Height - 850
Text1.Width = textviewer.Width - 150
End Sub

Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnusave_Click()
dir = Form1.Dir1.Path & "\" & Form1.File1.FileName
Open dir For Output As #1
Print #1, Text1.text
Close #1
End Sub

Private Sub mnusaveas_Click()
On Error GoTo er
Dialog.Filter = "Text|*.txt"
Dialog.CancelError = True
Dialog.ShowSave
dir = Dialog.FileName
Open dir For Output As #1
Print #1, Text1.text
Close #1
Form1.File1.Refresh
er:
Exit Sub
End Sub
