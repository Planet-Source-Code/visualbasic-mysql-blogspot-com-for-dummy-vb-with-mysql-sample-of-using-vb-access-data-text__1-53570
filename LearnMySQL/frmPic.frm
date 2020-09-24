VERSION 5.00
Begin VB.Form frmPic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmPic"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   3495
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   3255
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   1815
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
On Error Resume Next
Kill App.Path & "\Temp" & xExt
Browse.Show
End Sub

Private Sub cmdUpload_Click()
Main.ImageUpdate Main.Text1, frmPic.Text1, xName
Text1 = ""
End Sub

Private Sub Form_Load()
Dim lResult As Long
    lResult = SetWindowPos(Me.hwnd, HWND_TOPMOST, _
    0, 0, 0, 0, FLAGS)
End Sub

Private Sub Text1_Change()
If Text1 = "" Then
    cmdUpload.Enabled = False
Else
    cmdUpload.Enabled = True
End If

End Sub
