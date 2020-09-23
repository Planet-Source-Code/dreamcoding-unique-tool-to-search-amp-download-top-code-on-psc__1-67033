Attribute VB_Name = "modReadWriteINI"
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias _
"GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, _
ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias _
"WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, _
ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
ByVal lpString As Any, ByVal lpFileName As String) As Long

'// INI CONTROLLING PROCEDURES
'reads an Ini string
Public Function ReadIni(Filename As String, Section As String, Key As String) As String
Dim strBuf As String * 255
Dim l As Long
l = GetPrivateProfileString(Section, Key, "", strBuf, 255, Filename)
ReadIni = Left$(strBuf, l)
End Function

'reads an Ini section
Public Function ReadIniSection(Filename As String, Section As String) As String
Dim strBuf As String * 255
l = GetPrivateProfileSection(Section, Key, 255, Filename)
ReadIniSection = Left(strBuf, l)
End Function

'writes an Ini string
Public Sub WriteIni(Filename As String, Section As String, Key As String, Value As String)
WritePrivateProfileString Section, Key, Value, Filename
End Sub

'writes an Ini section
Public Sub WriteIniSection(Filename As String, Section As String, Value As String)
WritePrivateProfileSection Section, Value, Filename
End Sub

