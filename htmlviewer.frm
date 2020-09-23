VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form htmlviewer 
   BackColor       =   &H8000000D&
   Caption         =   "Html Viewer"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "htmlviewer.frx":0000
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   5106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "htmlviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Dim dir As String
If VBA.Right$(Form1.Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & Form1.File1.FileName)
Else
dir = (Form1.Dir1.Path & "\" & Form1.File1.FileName)
End If
WebBrowser1.Navigate (dir)
htmlviewer.Caption = htmlviewer.Caption & " - " & Form1.File1.FileName
End Sub

Private Sub Form_Resize()
On Error Resume Next
WebBrowser1.Width = htmlviewer.Width - 100
WebBrowser1.Height = htmlviewer.Height - 500
End Sub
