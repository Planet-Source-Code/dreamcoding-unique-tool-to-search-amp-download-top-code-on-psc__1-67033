Attribute VB_Name = "modStartWithWindows"
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_CURRENT_USER = &H80000001
Private Const REG_SZ = 1
Private Const KEY_WRITE = 131078
Enum PickUser
  AllUsers = HKEY_LOCAL_MACHINE
  CurrentUser = HKEY_CURRENT_USER
End Enum

Public Sub StartWithWindows(AppTitle As String, AppPath As String, LoadOnStart As Boolean, CurrentOrAllUsers As PickUser)
Dim hKey As Long

AppPath = AppPath & "\" & App.EXEName & ".exe"

If LoadOnStart = True Then
 'Load on Start
 RegOpenKeyEx CurrentOrAllUsers, "Software\Microsoft\Windows\CurrentVersion\Run-", 0, KEY_WRITE, hKey
 RegDeleteValue hKey, AppTitle
 RegCloseKey hKey
 RegOpenKeyEx CurrentOrAllUsers, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_WRITE, hKey
 RegSetValueEx hKey, AppTitle, 0, REG_SZ, AppPath, Len(AppPath)
 RegCloseKey hKey
Else

'Do NOT Load on Start
 RegOpenKeyEx CurrentOrAllUsers, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_WRITE, hKey
 RegDeleteValue hKey, AppTitle
 RegCloseKey hKey
 RegOpenKeyEx CurrentOrAllUsers, "Software\Microsoft\Windows\CurrentVersion\Run-", 0, KEY_WRITE, hKey
 RegSetValueEx hKey, AppTitle, 0, REG_SZ, AppPath, Len(AppPath)
 RegCloseKey hKey
End If

End Sub


