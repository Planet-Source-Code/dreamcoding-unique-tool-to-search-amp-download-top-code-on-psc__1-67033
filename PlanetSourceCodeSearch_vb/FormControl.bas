Attribute VB_Name = "modFormControl"
'Please Visit - http://Rbgcode.com

'For Dragging Borderless Forms...
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'For Animating Windows
Public Const AW_HOR_POSITIVE = &H1 'Animates the window from left to right. This flag can be used with roll or slide animation.
Public Const AW_HOR_NEGATIVE = &H2 'Animates the window from right to left. This flag can be used with roll or slide animation.
Public Const AW_VER_POSITIVE = &H4 'Animates the window from top to bottom. This flag can be used with roll or slide animation.
Public Const AW_VER_NEGATIVE = &H8 'Animates the window from bottom to top. This flag can be used with roll or slide animation.
Public Const AW_CENTER = &H10 'Makes the window appear to collapse inward if AW_HIDE is used or expand outward if the AW_HIDE is not used.
Public Const AW_HIDE = &H10000 'Hides the window. By default, the window is shown.
Public Const AW_ACTIVATE = &H20000 'Activates the window.
Public Const AW_SLIDE = &H40000 'Uses slide animation. By default, roll animation is used.
Public Const AW_BLEND = &H80000 'Uses a fade effect. This flag can be used only if hwnd is a top-level window.
Public Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean

'

Public Sub DragForm(Frm As Form)

On Local Error Resume Next

'Move the borderless form...
DoEvents
Call ReleaseCapture
DoEvents
Call SendMessage(Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
DoEvents
End Sub
