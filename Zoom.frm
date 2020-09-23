VERSION 5.00
Begin VB.Form Zoom 
   BackColor       =   &H8000000D&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Zoom"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2310
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Zooms: "
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         ToolTipText     =   "Enter a Number Between 1 and 100."
         Top             =   2160
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000D&
         Caption         =   "Percent"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000D&
         Caption         =   "Height And Width"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Percent"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2160
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Heigth"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
   End
End
Attribute VB_Name = "Zoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim percent As String
If Option1.Value = True Then
pictureviewer.Image1.Height = Val(Text1.text)
pictureviewer.Image1.Width = Val(Text2.text)
End If
If Option2.Value = True Then
pictureviewer.Image1.Width = pictureviewer.Image1.Width * ("." & Text3.text)
pictureviewer.Image1.Height = pictureviewer.Image1.Height * ("." & Text3.text)
End If
pictureviewer.Height = pictureviewer.Image1.Height + 150
pictureviewer.Width = pictureviewer.Image1.Width + 100
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.text = pictureviewer.Image1.Height
Text2.text = pictureviewer.Image1.Width
Text3.Enabled = False
End Sub

Private Sub Option1_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = False
End Sub

Private Sub Option2_Click()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = True
End Sub
