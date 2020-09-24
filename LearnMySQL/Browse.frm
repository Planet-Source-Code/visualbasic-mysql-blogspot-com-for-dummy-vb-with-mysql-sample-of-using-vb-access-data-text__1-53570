VERSION 5.00
Begin VB.Form Browse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   3120
         Width           =   975
      End
      Begin VB.FileListBox FileList 
         Appearance      =   0  'Flat
         Height          =   2760
         Left            =   2040
         Pattern         =   "*.jpg;*.bmp;*.gif"
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.DriveListBox Drive 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.DirListBox Folder 
         Appearance      =   0  'Flat
         Height          =   2790
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
      Begin VB.Image ImageBrowse 
         Height          =   2055
         Left            =   4080
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   3960
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Browse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FilePath, FileExt  As String
Dim r$

Private Sub cmdCancel_Click()
frmPic.Text1 = ""
Unload Browse
End Sub

Private Sub cmdOK_Click()
frmPic.Text1 = FileList.Path + r$ + FileList.FileName
frmPic.Image1.Visible = True
frmPic.Image1 = LoadPicture(FilePath)
xExt = Right(FilePath, 4)
xName = FileList.FileName
Unload Browse
End Sub

Private Sub Drive_Change()
Folder.Path = Drive
End Sub

Private Sub FileList_Click()
If Right$(FileList.Path, 1) = "\" Then r$ = "" Else r$ = "\"
FilePath = FileList.Path + r$ + FileList.FileName
FileExt = FileList.FileName
Text1 = FilePath
ImageBrowse.Picture = LoadPicture(FilePath)
End Sub

Private Sub FileList_DblClick()
cmdOK_Click
End Sub

Private Sub Folder_Change()
FileList.Path = Folder.Path
End Sub

Private Sub Form_Load()
Dim lResult As Long
    lResult = SetWindowPos(Me.hwnd, HWND_TOPMOST, _
    0, 0, 0, 0, FLAGS)
End Sub

Private Sub Text1_Change()
If Text1 = "" Then
    cmdOK.Enabled = False
Else
    cmdOK.Enabled = True
End If
End Sub
