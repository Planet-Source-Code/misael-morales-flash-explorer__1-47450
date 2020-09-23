VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "John Search"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   -510
   ClientWidth     =   11880
   Icon            =   "Search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Search.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Search.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Search.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Search.frx":113E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   953
      ButtonWidth     =   1217
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Run"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rename"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   240
      TabIndex        =   14
      Top             =   6960
      Width           =   11415
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H8000000D&
      Caption         =   "Date:"
      Height          =   615
      Left            =   2760
      TabIndex        =   12
      Top             =   4680
      Width           =   1215
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000D&
      Caption         =   "Time:"
      Height          =   615
      Left            =   1440
      TabIndex        =   10
      Top             =   4680
      Width           =   1215
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000D&
      Caption         =   "Filter"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      Caption         =   "Info: "
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   6255
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "File Type:"
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3480
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "File Attribute :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date: "
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Size: "
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Search "
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   11655
      Begin VB.FileListBox File1 
         Height          =   3405
         Left            =   6240
         TabIndex        =   2
         Top             =   240
         Width           =   5295
      End
      Begin VB.DirListBox Dir1 
         Height          =   3465
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4320
      Top             =   6120
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H8000000D&
      Caption         =   "Opened Files:"
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   6720
      Width           =   11655
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu files 
         Caption         =   "Files"
         Begin VB.Menu mnurun 
            Caption         =   "Run"
         End
         Begin VB.Menu mnurename 
            Caption         =   "Rename"
         End
         Begin VB.Menu mnudeletefile 
            Caption         =   "Delete"
         End
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu folders 
         Caption         =   "Folders"
         Begin VB.Menu mnunew 
            Caption         =   "New"
         End
         Begin VB.Menu mnudelete 
            Caption         =   "Delete"
         End
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub Combo1_Change()
On Error Resume Next
File1.Pattern = Combo1.text
End Sub

Private Sub Combo1_Click()
On Error Resume Next
File1.Pattern = Combo1.text
End Sub

Private Sub Combo2_Change()
On Error Resume Next
Dim dir As String
If VBA.Right$(Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & File1.FileName)
Else
dir = Dir1.Path & "\" & File1.FileName
End If
If Combo2.text = "Archive" Then
SetAttr dir, vbArchive
End If
If Combo2.text = "Read Only" Then
SetAttr dir, vbReadOnly
End If
End Sub

Private Sub Combo2_Click()
On Error Resume Next
Dim setit As String
setit = (Dir1.Path & "\" & File1.FileName)
If Combo2.text = "Archive" Then
SetAttr setit, vbArchive
End If
If Combo2.text = "Read Only" Then
SetAttr setit, vbReadOnly
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Call delete
File1.Refresh
VBA.Beep
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim foldername, dir As String
foldername = InputBox("Enter New Folder Name", "New Folder")
If Right$(Dir1.Path, 3) = "C:\" Then
dir = "C:\"
Else
dir = Dir1.Path + "\"
End If
MkDir (dir + foldername)
Dir1.Refresh
VBA.Beep
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim removefolder, dir As String
removefolder = InputBox("Enter Folder To Delete", "Delete Folder")
If Right$(Dir1.Path, 3) = "C:\" Then
dir = "C:\"
Else
dir = Dir1.Path + "\"
End If
RmDir (dir + removefolder)
Dir1.Refresh
VBA.Beep
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim dir As String
If VBA.Right$(Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & File1.FileName)
Else
dir = Dir1.Path & "\" & File1.FileName
End If
Call runit
List1.AddItem dir
End Sub

Private Sub Command5_Click()
If File1.FileName = "" Then Exit Sub
Load rename
rename.Show 1
End Sub

Private Sub Dir1_Change()
On Error Resume Next
File1.FileName = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopupMenu folders
End Sub

Private Sub File1_Click()
On Error Resume Next
Dim dir As String
Dim tam As String, vid As String
vid = VBA.Right(File1.FileName, 4)
tam = VBA.Right(File1.FileName, 3)
If VBA.Right$(Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & File1.FileName)
Else
dir = Dir1.Path & "\" & File1.FileName
End If
Label1.Caption = "Size: " & Format$(FileLen(dir), "###,###,###")
Label2.Caption = "Date : " & FileDateTime(dir)
Call filetype
If vid = "MPEG" Or vid = "Mpeg" Or vid = "mpeg" Or vid = "Jpeg" Or vid = "JPEG" Or vid = "jpeg" Then
Label6.Caption = vid
Else
Label6.Caption = tam
End If
End Sub
Private Sub File1_DblClick()
On Error Resume Next
Dim dir As String
If VBA.Right$(Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & File1.FileName)
Else
dir = Dir1.Path & "\" & File1.FileName
End If
Call runit
List1.AddItem dir
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopupMenu files
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Move 0, 0
Label5.Caption = Date
Dir1.Path = "C:\"
With Combo1
.text = "*.*"
.AddItem "*.*"
.AddItem "*.Exe"
.AddItem "*.Jpg"
.AddItem "*.Jpeg"
.AddItem "*.Bmp"
.AddItem "*.Gif"
.AddItem "*.Mpg"
.AddItem "*.Mpeg"
.AddItem "*.Avi"
.AddItem "*.Mov"
.AddItem "*.Mp3"
.AddItem "*.Txt"
.AddItem "*.Doc"
.AddItem "*.Ini"
.AddItem "*.Bat"
.AddItem "*.Com"
.AddItem "*.Zip"
End With
With Combo2
.AddItem "Read Only"
.AddItem "Archive"
End With
Call sndPlaySound("C:\WINDOWS\MEDIA\tada.wav", 1)
End Sub

Private Sub List1_Click()
Dim tam As String, vid As String
vid = VBA.Right(List1.text, 4)
tam = VBA.Right(List1.text, 3)
If vid = "MPEG" Or vid = "Mpeg" Or vid = "mpeg" Or vid = "Jpeg" Or vid = "JPEG" Or vid = "jpeg" Then
Label6.Caption = vid
Else
Label6.Caption = tam
End If
End Sub

Private Sub mnuabout_Click()
Load info
info.Show 1
End Sub

Private Sub mnuclose_Click()
On Error Resume Next
End
Set Form1 = Nothing
End Sub

Private Sub filetype()
On Error Resume Next
Dim attr, dir As String
If VBA.Right$(Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & File1.FileName)
Else
dir = Dir1.Path & "\" & File1.FileName
End If
attr = GetAttr(dir)
If attr = "1" Then Combo2.text = "Read Only"
If attr = "32" Then Combo2.text = "Archive"
End Sub

Private Sub runfile()
On Error Resume Next
If VBA.Right$(Dir1.Path, 3) = "C:\" Then
Shell ("C:\" & File1.FileName), vbNormalFocus
Else
Shell (Dir1.Path + "\" + File1.FileName), vbNormalFocus
End If
End Sub

Private Sub delete()
On Error Resume Next
Dim message, dir As String
If Right$(File1.FileName, 3) = "Com" Or Right$(File1.FileName, 3) = "COM" Or Right$(File1.FileName, 3) = "com" Or Right$(File1.FileName, 3) = "Bat" Or Right$(File1.FileName, 3) = "BAT" Or Right$(File1.FileName, 3) = "bat" Or Right$(File1.FileName, 3) = "Ini" Or Right$(File1.FileName, 3) = "INI" Or Right$(File1.FileName, 3) = "ini" Then
MsgBox "Sorry I Can't Let You Delete That File Cause You Can Crash Your Computer", vbCritical + vbOKOnly, "Aborting Delete"
File1.Refresh
Exit Sub
End If
message = MsgBox("Are you sure about permanently delete ' " & File1.FileName & " ' ?", vbCritical + vbOKCancel, "Confirm Delete")
If message = vbOK Then
If Right$(Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & File1.FileName)
Else
dir = (Dir1.Path & "\" & File1.FileName)
End If
Kill (dir)
File1.Refresh
End If
End Sub

Private Sub text()
Load textviewer
textviewer.Show
End Sub

Private Sub runit()
On Error Resume Next
Dim tam As String, vid As String
Dim dir As String
If VBA.Right$(Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & File1.FileName)
Else
dir = Dir1.Path & "\" & File1.FileName
End If
vid = VBA.Right(File1.FileName, 4)
tam = VBA.Right(File1.FileName, 3)
If tam = "exe" Or tam = "EXE" Or tam = "Exe" Then Call runfile
If tam = "mpg" Or tam = "MPG" Or tam = "Mpg" Or vid = "mpeg" Or vid = "MPEG" Or vid = "Mpeg" Or tam = "avi" Or tam = "AVI" Or tam = "Avi" Or tam = "ASF" Or tam = "Asf" Or tam = "asf" Then
Load movie
movie.Show
End If
If tam = "txt" Or tam = "TXT" Or tam = "Txt" Or tam = "INI" Or tam = "Ini" Or tam = "ini" Or tam = "DOC" Or tam = "Doc" Or tam = "doc" Then Call text
If tam = "jpg" Or tam = "Jpg" Or tam = "JPG" Or tam = "bmp" Or tam = "Bmp" Or tam = "BMP" Or VBA.Right(File1.FileName, 4) = "jpeg" Or VBA.Right(File1.FileName, 4) = "Jpeg" Or VBA.Right(File1.FileName, 4) = "JPEG" Or tam = "GIF" Or tam = "Gif" Or tam = "gif" Then
Load pictureviewer
pictureviewer.Show 1
End If
If tam = "Mp3" Or tam = "mp3" Or tam = "MP3" Or tam = "WAV" Or tam = "Wav" Or tam = "wav" Then
Load Mp3
Mp3.Show
End If
If tam = "Htm" Or tam = "htm" Or tam = "Html" Or vid = "Html" Or vid = "html" Or vid = "HTML" Then
Load htmlviewer
htmlviewer.Show
End If
If tam = "ZIP" Or tam = "Zip" Or tam = "zip" Then
Shell "start " & dir
End If
If tam = "hlp" Or tam = "HLP" Or tam = "Hlp" Or tam = "CHM" Or tam = "Chm" Or tam = "chm" Then runhelp
End Sub

Private Sub mnudelete_Click()
On Error Resume Next
Dim removefolder, dir As String
removefolder = InputBox("Enter Folder To Delete", "Delete Folder")
If Right$(Dir1.Path, 3) = "C:\" Then
dir = "C:\"
Else
dir = Dir1.Path + "\"
End If
RmDir (dir + removefolder)
Dir1.Refresh
VBA.Beep
End Sub

Private Sub mnudeletefile_Click()
On Error Resume Next
Call delete
File1.Refresh
VBA.Beep
End Sub

Private Sub mnunew_Click()
On Error Resume Next
Dim foldername, dir As String
foldername = InputBox("Enter New Folder Name", "New Folder")
If Right$(Dir1.Path, 3) = "C:\" Then
dir = "C:\"
Else
dir = Dir1.Path + "\"
End If
MkDir (dir + foldername)
Dir1.Refresh
VBA.Beep
End Sub

Private Sub mnurename_Click()
If File1.FileName = "" Then Exit Sub
Load rename
rename.Show 1
End Sub

Private Sub mnurun_Click()
On Error Resume Next
Dim dir As String
If VBA.Right$(Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & File1.FileName)
Else
dir = Dir1.Path & "\" & File1.FileName
End If
Call runit
List1.AddItem dir
End Sub

Private Sub ssss_Click()

End Sub

Private Sub Timer1_Timer()
Label4.Caption = Time
End Sub

Private Sub runhelp()
Dim dir As String
If VBA.Right$(Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & File1.FileName)
Else
dir = Dir1.Path & "\" & File1.FileName
End If
Shell "Start " & dir
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
On Error Resume Next
Dim dir As String
If VBA.Right$(Dir1.Path, 3) = "C:\" Then
dir = ("C:\" & File1.FileName)
Else
dir = Dir1.Path & "\" & File1.FileName
End If
Call runit
List1.AddItem dir
Case 2
If File1.FileName = "" Then Exit Sub
Load rename
rename.Show 1
Case 3
On Error Resume Next
Call delete
File1.Refresh
VBA.Beep
End Select
End Sub

Private Sub selectmenu()
Select Case ButtonMenus
Case 1
MsgBox "New"
Case 2
MsgBox "Delete"
End Select
End Sub




