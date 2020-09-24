Attribute VB_Name = "Module1"
Option Explicit

Declare Function GetSystemMenu Lib "user32" _
        (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Declare Function GetMenuItemCount Lib "user32" _
        (ByVal hMenu As Long) As Long

Declare Function DrawMenuBar Lib "user32" _
        (ByVal hwnd As Long) As Long

Declare Function RemoveMenu Lib "user32" _
        (ByVal hMenu As Long, ByVal nPosition As Long, _
        ByVal wFlags As Long) As Long
    
    Public Const MF_BYPOSITION = &H400&
    Public Const MF_REMOVE = &H1000&
    
Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
    Public Const SWP_NOMOVE = 2
    Public Const SWP_NOSIZE = 1
    Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2
    
    Public xExt As String
    Public xName As String
    
Public Sub DisableX(frmX As Form)
        Dim hSysMenu As Long
        Dim nCnt As Long
        hSysMenu = GetSystemMenu(frmX.hwnd, False)
        If hSysMenu Then
            nCnt = GetMenuItemCount(hSysMenu)
            If nCnt Then
                RemoveMenu hSysMenu, nCnt - 1, _
                    MF_BYPOSITION Or MF_REMOVE
                RemoveMenu hSysMenu, nCnt - 2, _
                    MF_BYPOSITION Or MF_REMOVE
                DrawMenuBar frmX.hwnd
            End If
        End If
End Sub

