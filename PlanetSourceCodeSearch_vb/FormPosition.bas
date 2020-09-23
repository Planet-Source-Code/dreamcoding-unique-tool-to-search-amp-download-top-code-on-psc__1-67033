Attribute VB_Name = "modFormPosition"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
'Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Form_Activate(TheWin)
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    'Set the window position to topmost
    SetWindowPos TheWin.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
Sub Win_OnTop(TheFrm As Form)
    Dim SetOnTop
    SetWindowPos TheFrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Sub Win_ControlOnTop(TheFrm, isVisible As Boolean)
    Dim SetOnTop
    Dim hwnd
    If isVisible = True Then
        SetOnTop = SetWindowPos(TheFrm.hwnd, HWND_TOPMOST, 100, 100, 300, 100, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
    Else
        SetOnTop = SetWindowPos(TheFrm.hwnd, HWND_TOPMOST, 100, 100, 300, 100, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_HIDEWINDOW)
    End If
End Sub
Sub WIN_NotOnTop(the As Form)

Dim Flags
Dim SetWinNotOnTop As Long
 SetWinNotOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
'WinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End Sub
