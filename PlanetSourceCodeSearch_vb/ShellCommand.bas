Attribute VB_Name = "modShellCommand"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function apiFindWindow Lib "user32" Alias "FindWindowA" (ByVal lpclassname As Any, ByVal lpCaption As Any) As Long

Global Const SW_SHOWNORMAL = 1

'Sub ShellExecuteExample()
'   Dim hwnd
'   Dim StartDoc
'   hwnd = apiFindWindow("OPUSAPP", "0")
'
'   StartDoc = ShellExecute(hwnd, "open", "C:\My Documents\Book1.xls", "", _
'      "C:\", SW_SHOWNORMAL)
'End Sub
