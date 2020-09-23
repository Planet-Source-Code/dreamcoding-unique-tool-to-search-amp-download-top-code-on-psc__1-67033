VERSION 5.00
Begin VB.Form frmSysTray 
   Caption         =   "Sys Tray Interface"
   ClientHeight    =   1920
   ClientLeft      =   5610
   ClientTop       =   3360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   4680
   Begin VB.Menu mnuPopup 
      Caption         =   "&Popup"
      Begin VB.Menu mnuSysTray 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 03/03/2003
' * Added Unicode support
' * Added support for new tray version (ME,2000 or above required)
' * Added support for balloon tips (ME,2000 or above required)
' frmSysTray.
' Steve McMahon
' Original version based on code supplied from Ben Baird:
'Author:
'        Ben Baird <psyborg@cyberhighway.com>
'        Copyright (c) 1997, Ben Baird
'
'Purpose:
'        Demonstrates setting an icon in the taskbar's
'        system tray without the overhead of subclassing
'        to receive events.
'IMPLEMENT
'Private WithEvents m_frmSysTray As frmSysTray
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Unload m_frmSysTray
'    Set m_frmSysTray = Nothing
'End Sub
'
'Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
'   Select Case sKey
'   Case "open"
'      Me.Show
'      Me.ZOrder
'   Case "close"
'      Unload Me
'   Case Else
'      MsgBox "Clicked item with key " & sKey, vbInformation
'   End Select
'
'End Sub '
'
'Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
'    Me.Show
'    Me.ZOrder
'End Sub '
'
'Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
'    If (eButton = vbRightButton) Then
'        m_frmSysTray.ShowMenu
'    End If
'End Sub
'Private Sub SetIcon()
'    Select Case True
'    Case optIcon(0).Value
'        m_frmSysTray.IconHandle = imgIcon(0).Picture.handle
'    Case optIcon(1).Value
'        m_frmSysTray.IconHandle = imgIcon(1).Picture.handle
'    Case optIcon(2).Value
'        m_frmSysTray.IconHandle = imgIcon(2).Picture.handle
'    End Select
'End Sub
'
'Private Sub chkSysTray_Click()
'    If (chkSysTray.Value = Checked) Then
'        Set m_frmSysTray = New frmSysTray
'        With m_frmSysTray
'            .AddMenuItem "&Open SysTray Sample", "open", True
'            .AddMenuItem "-"
'            .AddMenuItem "&vbAccelerator on the Web", "vbAccelerator"
'            .AddMenuItem "&About...", "About"
'            .AddMenuItem "-"
'            .AddMenuItem "&Close", "close"
'            .ToolTip = "SysTray Sample!"
'        End With
'        SetIcon
'    Else
'        Unload m_frmSysTray
'        Set m_frmSysTray = Nothing
'    End If
'End Sub '
'
'Private Sub cmdShowBalloon_Click()
'   m_frmSysTray.ShowBalloonTip
''      "Hello from vbAccelerator.com.  This SysTray form allows Unicode text and balloon tips.",
''      "vbAccelerator SysTray Sample",
''      NIIF_INFO
'End Sub
Private Const NIF_ICON                             As Long = &H2
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const NIF_MESSAGE                          As Long = &H1
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const NIF_TIP                              As Long = &H4
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const NIF_STATE                            As Long = &H8
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const NIF_INFO                             As Long = &H10
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const NIM_ADD                              As Long = &H0
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const NIM_MODIFY                           As Long = &H1
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const NIM_DELETE                           As Long = &H2
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const NIM_SETFOCUS                         As Long = &H3
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const NIM_SETVERSION                       As Long = &H4
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const NOTIFYICON_VERSION                   As Integer = 3
'<CF> :UPDATED: UnTyped Const with numeric value changed to As Integer
Private Type NOTIFYICONDATAA
    cbSize                                             As Long    ' 4
    hWnd                                               As Long    ' 8
    uID                                                As Long    ' 12
    uFlags                                             As Long    ' 16
    uCallbackMessage                                   As Long    ' 20
    hIcon                                              As Long    ' 24
    szTip                                              As String * 128 ' 152
    dwState                                            As Long    ' 156
    dwStateMask                                        As Long    ' 160
    szInfo                                             As String * 256 ' 416
    uTimeOutOrVersion                                  As Long    ' 420
    szInfoTitle                                        As String * 64 ' 484
    dwInfoFlags                                        As Long    ' 488
    guidItem                                           As Long    ' 492
End Type
Private Type NOTIFYICONDATAW
    cbSize                                             As Long    ' 4
    hWnd                                               As Long    ' 8
    uID                                                As Long    ' 12
    uFlags                                             As Long    ' 16
    uCallbackMessage                                   As Long    ' 20
    hIcon                                              As Long    ' 24
    szTip(0 To 255)                                    As Byte    ' 280
    dwState                                            As Long    ' 284
    dwStateMask                                        As Long    ' 288
    szInfo(0 To 511)                                   As Byte    ' 800
    uTimeOutOrVersion                                  As Long    ' 804
    szInfoTitle(0 To 127)                              As Byte    ' 932
    dwInfoFlags                                        As Long    ' 936
    guidItem                                           As Long    ' 940
End Type
Private nfIconDataA                                As NOTIFYICONDATAA
Private nfIconDataW                                As NOTIFYICONDATAW
Private Const NOTIFYICONDATAA_V1_SIZE_A            As Integer = 88
'<CF> :UPDATED: UnTyped Const with numeric value changed to As Integer
Private Const NOTIFYICONDATAA_V1_SIZE_U            As Integer = 152
'<CF> :UPDATED: UnTyped Const with numeric value changed to As Integer
Private Const NOTIFYICONDATAA_V2_SIZE_A            As Integer = 488
'<CF> :UPDATED: UnTyped Const with numeric value changed to As Integer
Private Const NOTIFYICONDATAA_V2_SIZE_U            As Integer = 936
'<CF> :UPDATED: UnTyped Const with numeric value changed to As Integer
Private Const WM_MOUSEMOVE                         As Long = &H200
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const WM_LBUTTONDBLCLK                     As Long = &H203
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const WM_LBUTTONDOWN                       As Long = &H201
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const WM_LBUTTONUP                         As Long = &H202
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const WM_RBUTTONDBLCLK                     As Long = &H206
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const WM_RBUTTONDOWN                       As Long = &H204
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const WM_RBUTTONUP                         As Long = &H205
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const WM_USER                              As Long = &H400
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const NIN_SELECT                           As Long = WM_USER
'<CF> :UPDATED: UnTyped Const with API sytle Name changed set to As Long
Private Const NINF_KEY                             As Long = &H1
'<CF> :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
Private Const NIN_KEYSELECT                        As Double = (NIN_SELECT Or NINF_KEY)
'<CF> :UPDATED: UnTyped Const with value using Math or Logic operation changed to As Double
Private Const NIN_BALLOONSHOW                      As Double = (WM_USER + 2)
'<CF> :UPDATED: UnTyped Const with value using Math or Logic operation changed to As Double
Private Const NIN_BALLOONHIDE                      As Double = (WM_USER + 3)
'<CF> :UPDATED: UnTyped Const with value using Math or Logic operation changed to As Double
Private Const NIN_BALLOONTIMEOUT                   As Double = (WM_USER + 4)
'<CF> :UPDATED: UnTyped Const with value using Math or Logic operation changed to As Double
Private Const NIN_BALLOONUSERCLICK                 As Double = (WM_USER + 5)
'<CF> :UPDATED: UnTyped Const with value using Math or Logic operation changed to As Double
' Version detection:
Public Event SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseMove()
Public Event SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
Public Event MenuClick(ByVal lIndex As Long, ByVal sKey As String)
Public Event BalloonShow()
Public Event BalloonHide()
Public Event BalloonTimeOut()
Public Event BalloonClicked()
Public Enum EBalloonIconTypes
    NIIF_NONE = 0
    NIIF_INFO = 1
    NIIF_WARNING = 2
    NIIF_ERROR = 3
    NIIF_NOSOUND = &H10
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private NIIF_NONE, NIIF_INFO, NIIF_WARNING, NIIF_ERROR, NIIF_NOSOUND
#End If
'<CF> :SUGGESTION: Inserted by Code Fixer. (Must be placed after Enum Declaration for Code Fixer to recognize it properly)
Private m_bAddedMenuItem                           As Boolean
Private m_iDefaultIndex                            As Long
Private m_bUseUnicode                              As Boolean
Private m_bSupportsNewVersion                      As Boolean
Private Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, _
                                                              lpData As NOTIFYICONDATAA) As Long
Private Declare Function Shell_NotifyIconW Lib "shell32.dll" (ByVal dwMessage As Long, _
                                                              lpData As NOTIFYICONDATAW) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Public Function AddMenuItem(ByVal sCaption As String, _
                            Optional ByVal sKey As String = vbNullString, _
                            Optional ByVal bDefault As Boolean = False) As Long
Dim iIndex As Long
    On Error Resume Next
    If Not (m_bAddedMenuItem) Then
        iIndex = 0
        m_bAddedMenuItem = True
    Else
        iIndex = mnuSysTray.UBound + 1
        Load mnuSysTray(iIndex)
    End If
    With mnuSysTray(iIndex)
        .Visible = True
        .Tag = sKey
        .Caption = sCaption
    End With 'mnuSysTray(iIndex)
    If bDefault Then
        m_iDefaultIndex = iIndex
    End If
    AddMenuItem = iIndex
End Function
Public Property Get DefaultMenuIndex() As Long
    DefaultMenuIndex = m_iDefaultIndex
End Property
Public Property Let DefaultMenuIndex(ByVal lIndex As Long)
    If ValidIndex(lIndex) Then
        m_iDefaultIndex = lIndex
    Else
        m_iDefaultIndex = 0
    End If
End Property
Public Sub EnableMenuItem(ByVal lIndex As Long, _
                          ByVal bState As Boolean)
    If ValidIndex(lIndex) Then
        mnuSysTray(lIndex).Enabled = bState
    End If
End Sub
Private Sub Form_Load()
' Get version:
Dim lMajor As Long
Dim lMinor As Long
Dim bIsNt  As Long
Dim lR     As Long
    GetWindowsVersion lMajor, lMinor, , , bIsNt
    If bIsNt Then
        m_bUseUnicode = True
        If lMajor >= 5 Then
' 2000 or XP
            m_bSupportsNewVersion = True
        End If
    ElseIf (lMajor = 4) And (lMinor = 90) Then
' Windows ME
        m_bSupportsNewVersion = True
    End If
'Add the icon to the system tray...
    If m_bUseUnicode Then
        With nfIconDataW
            .hWnd = Me.hWnd
            .uID = Me.Icon
            .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon.Handle
            stringToArray App.FileDescription, .szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
            If m_bSupportsNewVersion Then
                .uTimeOutOrVersion = NOTIFYICON_VERSION
            End If
            .cbSize = nfStructureSize
        End With
        lR = Shell_NotifyIconW(NIM_ADD, nfIconDataW)
        If m_bSupportsNewVersion Then
            Shell_NotifyIconW NIM_SETVERSION, nfIconDataW
        End If
    Else
        With nfIconDataA
            .hWnd = Me.hWnd
            .uID = Me.Icon
            .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon.Handle
            .szTip = App.FileDescription & vbNullChar
            If m_bSupportsNewVersion Then
                .uTimeOutOrVersion = NOTIFYICON_VERSION
            End If
            .cbSize = nfStructureSize
        End With
        lR = Shell_NotifyIconA(NIM_ADD, nfIconDataA)
        If m_bSupportsNewVersion Then
            lR = Shell_NotifyIconA(NIM_SETVERSION, nfIconDataA)
        End If
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)
Dim lX As Long
' VB manipulates the x value according to scale mode:
' we must remove this before we can interpret the
' message windows was trying to send to us:
    lX = ScaleX(x, Me.ScaleMode, vbPixels)
    Select Case lX
    Case WM_MOUSEMOVE
        RaiseEvent SysTrayMouseMove
    Case WM_LBUTTONUP
        RaiseEvent SysTrayMouseDown(vbLeftButton)
    Case WM_LBUTTONUP
        RaiseEvent SysTrayMouseUp(vbLeftButton)
    Case WM_LBUTTONDBLCLK
        RaiseEvent SysTrayDoubleClick(vbLeftButton)
    Case WM_RBUTTONDOWN
        RaiseEvent SysTrayMouseDown(vbRightButton)
    Case WM_RBUTTONUP
        RaiseEvent SysTrayMouseUp(vbRightButton)
    Case WM_RBUTTONDBLCLK
        RaiseEvent SysTrayDoubleClick(vbRightButton)
    Case NIN_BALLOONSHOW
        RaiseEvent BalloonShow
    Case NIN_BALLOONHIDE
        RaiseEvent BalloonHide
    Case NIN_BALLOONTIMEOUT
        RaiseEvent BalloonTimeOut
    Case NIN_BALLOONUSERCLICK
        RaiseEvent BalloonClicked
    End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    If m_bUseUnicode Then
        Shell_NotifyIconW NIM_DELETE, nfIconDataW
    Else
        Shell_NotifyIconA NIM_DELETE, nfIconDataA
    End If
End Sub
Private Sub GetWindowsVersion(Optional ByRef lMajor = 0, _
                              Optional ByRef lMinor = 0, _
                              Optional ByRef lRevision = 0, _
                              Optional ByRef lBuildNumber = 0, _
                              Optional ByRef bIsNt = False)
Dim lR As Long
    lR = GetVersion()
    lBuildNumber = (lR And &H7F000000) \ &H1000000
    If (lR And &H80000000) Then
        lBuildNumber = lBuildNumber Or &H80
    End If
    lRevision = (lR And &HFF0000) \ &H10000
    lMinor = (lR And &HFF00&) \ &H100
    lMajor = (lR And &HFF)
    bIsNt = ((lR And &H80000000) = 0)
End Sub
Public Property Get IconHandle() As Long
    IconHandle = nfIconDataA.hIcon
End Property

Public Property Let IconHandle(ByVal hIcon As Long)
    If m_bUseUnicode Then
        With nfIconDataW
            If hIcon <> .hIcon Then
                .hIcon = hIcon
                .uFlags = NIF_ICON
                Shell_NotifyIconW NIM_MODIFY, nfIconDataW
            End If
        End With 'nfIconDataW
    Else
        With nfIconDataA
            If hIcon <> .hIcon Then
                .hIcon = hIcon
                .uFlags = NIF_ICON
                Shell_NotifyIconA NIM_MODIFY, nfIconDataA
            End If
        End With 'nfIconDataA
    End If
End Property

Private Sub mnuSysTray_Click(Index As Integer)
    RaiseEvent MenuClick(Index, mnuSysTray(Index).Tag)
End Sub
Private Property Get nfStructureSize() As Long
    If m_bSupportsNewVersion Then
        If m_bUseUnicode Then
            nfStructureSize = NOTIFYICONDATAA_V2_SIZE_U
        Else
            nfStructureSize = NOTIFYICONDATAA_V2_SIZE_A
        End If
    Else
        If m_bUseUnicode Then
            nfStructureSize = NOTIFYICONDATAA_V1_SIZE_U
        Else
            nfStructureSize = NOTIFYICONDATAA_V1_SIZE_A
        End If
    End If
End Property
Public Function RemoveMenuItem(ByVal iIndex As Long) As Long
Dim i As Long
    If ValidIndex(iIndex) Then
        If iIndex = 0 Then
            mnuSysTray(0).Caption = vbNullString
        Else
' remove the item:
            For i = iIndex + 1 To mnuSysTray.UBound
                mnuSysTray(iIndex - 1).Caption = mnuSysTray(iIndex).Caption
                mnuSysTray(iIndex - 1).Tag = mnuSysTray(iIndex).Tag
            Next i
            Unload mnuSysTray(mnuSysTray.UBound)
        End If
    End If
End Function
Public Sub ShowBalloonTip(ByVal sMessage As String, _
                          Optional ByVal sTitle As String, _
                          Optional ByVal eIcon As EBalloonIconTypes, _
                          Optional ByVal lTimeOutMs = 30000)
Dim lR As Long
    If m_bSupportsNewVersion Then

        If m_bUseUnicode Then
            With nfIconDataW
                stringToArray sMessage, .szInfo, 512
                stringToArray sTitle, .szInfoTitle, 128
                .uTimeOutOrVersion = lTimeOutMs
                .dwInfoFlags = eIcon
                .uFlags = NIF_INFO
            End With 'nfIconDataW
            lR = Shell_NotifyIconW(NIM_MODIFY, nfIconDataW)
        Else
            With nfIconDataA
                .szInfo = sMessage
                .szInfoTitle = sTitle
                .uTimeOutOrVersion = lTimeOutMs
                .dwInfoFlags = eIcon
                .uFlags = NIF_INFO
            End With 'nfIconDataA
            lR = Shell_NotifyIconA(NIM_MODIFY, nfIconDataA)
        End If
    Else
' can't do it, fail silently.
    End If
End Sub
Public Function ShowMenu()
    SetForegroundWindow Me.hWnd
    If m_iDefaultIndex > -1 Then
        Me.PopupMenu mnuPopup, 0, , , mnuSysTray(m_iDefaultIndex)
    Else
        Me.PopupMenu mnuPopup, 0
    End If
End Function
Private Sub stringToArray(ByVal sString As String, _
                          bArray() As Byte, _
                          ByVal lMaxSize As Long)
Dim b() As Byte
Dim i   As Long
Dim j   As Long
    If Len(sString) > 0 Then
        b = sString
        For i = LBound(b) To UBound(b)
            bArray(i) = b(i)
            If (i = (lMaxSize - 2)) Then
                Exit For
            End If
        Next i
        For j = i To lMaxSize - 1
            bArray(j) = 0
        Next j
    End If
End Sub
Public Property Get ToolTip() As String
Dim sTip As String
Dim iPos As Long
    sTip = nfIconDataA.szTip
    iPos = InStr(sTip, vbNullChar)
    If iPos <> 0 Then
        sTip = Left$(sTip, iPos - 1)
    End If
    ToolTip = sTip
End Property
Public Property Let ToolTip(ByVal sTip As String)
    If m_bUseUnicode Then
        stringToArray sTip, nfIconDataW.szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
        nfIconDataW.uFlags = NIF_TIP
        Shell_NotifyIconW NIM_MODIFY, nfIconDataW
    Else
        If (sTip & vbNullChar <> nfIconDataA.szTip) Then
            nfIconDataA.szTip = sTip & vbNullChar
            nfIconDataA.uFlags = NIF_TIP
            Shell_NotifyIconA NIM_MODIFY, nfIconDataA
        End If
    End If
End Property
Private Function unicodeSize(ByVal lSize As Long) As Long
    If m_bUseUnicode Then

        unicodeSize = lSize * 2
    Else
        unicodeSize = lSize
    End If
End Function
Private Function ValidIndex(ByVal lIndex As Long) As Boolean
    ValidIndex = (lIndex >= mnuSysTray.LBound And lIndex <= mnuSysTray.UBound)
End Function
':)Code Fixer V4.0.0 (Tuesday, 06 June 2006 03:42:17) 139 + 479 = 618 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|033322222222222222222222222222|1112222|2221222|222222222233|1111111111111|1122222222220|333333|


