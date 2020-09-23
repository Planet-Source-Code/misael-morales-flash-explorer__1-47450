VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form pictureviewer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Picture Viewer"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   2745
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pictureviewer.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pictureviewer.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Zoom In"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Zoom Out"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   600
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   120
      Top             =   480
      Width           =   2535
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "Options"
      Begin VB.Menu mnusize 
         Caption         =   "Size"
      End
      Begin VB.Menu mnuzooms 
         Caption         =   "Zooms"
         Begin VB.Menu mnuundosize 
            Caption         =   "Undo"
         End
         Begin VB.Menu mnuzoomin 
            Caption         =   "Zoom In"
         End
         Begin VB.Menu mnuzoomout 
            Caption         =   "Zoom Out"
         End
         Begin VB.Menu mnuothers 
            Caption         =   "More Zooms"
         End
      End
      Begin VB.Menu mnuinfo 
         Caption         =   "Width And Height"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "pictureviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This example needs a Picture box (Picture1)
'with an picture loaded in it
Private Const IMAGE_BITMAP = 0
Private Const LR_COPYRETURNORG = &H4
Private Const CF_BITMAP = 2
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Dim picheight As String, picwidth As String
Private Sub Form_Load()
On Error Resume Next
Dim dir, der As String, search As String, dire As String
If VBA.Right$(Form1.Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & Form1.File1.FileName)
Else
dir = Form1.Dir1.Path & "\" & Form1.File1.FileName
End If
Image1.Picture = LoadPicture(dir)
If Image1.Height >= 8000 Or Image1.Width >= 8000 Then
Image1.Height = Image1.Height / 2
Image1.Width = Image1.Width / 2
picheight = Image1.Height
picwidth = Image1.Width
End If
Image1.Stretch = True
pictureviewer.Height = Image1.Height + 700
pictureviewer.Width = Image1.Width + 300
pictureviewer.Caption = pictureviewer.Caption & " - " & Form1.File1.FileName
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu mnuoptions, vbPopupMenuCenterAlign
Else
Exit Sub
End If
End Sub

Private Sub mnuabout_Click()
MsgBox "Made By Misael...!!Enjoy!!"
End Sub

Private Sub mnuagrandar_Click()
Dim percent As String
percent = InputBox("Enter a Number Between 1 and 100", "Zoom By Percent")
Image1.Width = Image1.Width * ("." & percent)
Image1.Height = Image1.Height * ("." & percent)
pictureviewer.Height = Image1.Height + 700
pictureviewer.Width = Image1.Width + 300
End Sub

Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnuinfo_Click()
MsgBox "Height: " & Format$(Image1.Height, "###,###,###") & " " & "And" & " " & "Width: " & Format$(Image1.Width, "###,###,###")
End Sub

Private Sub mnuothers_Click()
Load Zoom
Zoom.Show 1
End Sub

Private Sub mnusave_Click()
On Error Resume Next
Dim hNew As Long
Dim heigthimage As String, widthimage As String
   heigthimage = Image1.Height
   widthimage = Image1.Width
    'create an exact copy of the picture
    hNew = CopyImage(Image1.Picture, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
    'open the clipboard
    OpenClipboard Me.hwnd
    'clear the clipboard
    EmptyClipboard
    'put the picture on the clipboard
    SetClipboardData CF_BITMAP, hNew
    'close the clipboard
    CloseClipboard
    'note that we don't have to call DeleteObject(hNew)
    'from now on, the clipboard takes care of the bitmap
End Sub

Private Sub mnusize_Click()
Dim dir, der As String
If VBA.Right$(Form1.Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & Form1.File1.FileName)
Else
dir = Form1.Dir1.Path & "\" & Form1.File1.FileName
End If
MsgBox Format$(FileLen(dir), "###,###,###")
End Sub

Private Sub mnuundosize_Click()
On Error Resume Next
Image1.Height = picheight
Image1.Width = picwidth
pictureviewer.Height = Image1.Height + 700
pictureviewer.Width = Image1.Width + 300
End Sub

Private Sub mnuzoomin_Click()
Image1.Width = Image1.Width + 100
Image1.Height = Image1.Height + 100
pictureviewer.Height = Image1.Height + 700
pictureviewer.Width = Image1.Width + 300
End Sub

Private Sub mnuzoomout_Click()
Image1.Width = Image1.Width - 100
Image1.Height = Image1.Height - 100
pictureviewer.Height = Image1.Height + 700
pictureviewer.Width = Image1.Width + 300
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
Case 1
Image1.Height = Image1.Height + 250
Image1.Width = Image1.Width + 250
pictureviewer.Height = Image1.Height + 700
pictureviewer.Width = Image1.Width + 300
Case 2
Image1.Height = Image1.Height - 250
Image1.Width = Image1.Width - 250
pictureviewer.Height = Image1.Height + 700
pictureviewer.Width = Image1.Width + 300

End Select

End Sub
