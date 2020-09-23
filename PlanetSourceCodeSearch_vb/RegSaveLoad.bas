Attribute VB_Name = "modRegSave"
Option Explicit
'Dreamstruct

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_READ = &H20019

Public Sub LoadFormPosition(NameOfForm As Form)
On Error Resume Next
Dim FormIndex      As String
Dim FormPosition() As String
    With NameOfForm
        FormIndex = .Index     'If there is no form index, resume next line of execution
        'Split the string array to get the Left, Top, Width, Height (Position) of the Form
        FormPosition = Split(GetSetting(App.EXEName, "Settings", .Name & "_" & FormIndex & "_Position", .Left & "*" & .Top & "*" & .Width & "*" & .Height), "*")
        'Move the Form to Specified Position
        .Move FormPosition(0), FormPosition(1), FormPosition(2), FormPosition(3)
    End With
End Sub

Public Function RegLoad(ControlToRemember As Control) As Variant
'You may want to add controls to this routine such as listview.
'Some controls you will want to save as different data types e.g. optionbutton is boolean

On Error Resume Next
Dim NameOfForm   As String
Dim ControlIndex As String
Dim EndResult

    With ControlToRemember
        ControlIndex = .Index  'If there is no control index, resume next line of execution
        NameOfForm = .Parent.Name
        EndResult = GetSetting(App.EXEName, "Settings", .Parent.Name & "_" & .Name & "_" & ControlIndex)
    End With
        If EndResult = vbNullString Then
            EndResult = ControlToRemember.Value
        End If
        If TypeOf ControlToRemember Is CheckBox Then
            EndResult = CInt(Val(EndResult))
        End If
        If TypeOf ControlToRemember Is ComboBox Then
            EndResult = CInt(Val(EndResult))
        End If
        If TypeOf ControlToRemember Is HScrollBar Then
            EndResult = CLng(Val(EndResult))
        End If
        If TypeOf ControlToRemember Is ListBox Then
            EndResult = CLng(Val(EndResult))
        End If
        'If TypeOf ControlToRemember Is ListView Then
        '    EndResult = CLng(Val(EndResult))
        'End If
        If TypeOf ControlToRemember Is Timer Then
            EndResult = CLng(Val(EndResult))
        End If
        If TypeOf ControlToRemember Is OptionButton Then
            EndResult = CBool(EndResult)
        End If
        
RegLoad = EndResult
End Function

'**********************************************************************************************************
'modRegSave.bas original posting by Phishbowla
'Much Thanks to Roger Gilchrist for turning good code into great code for me.
'The code is much cleaner and has been optimized. His advice and fixes has helped my learning greatly.
'Limitations to this code:
'You can only save one property for each control.
'e.g.
'If list1.text = regsave
'if list1.listindex = regsave
'Saving listindex for list1 will overwrite saving text for list1.
'**********************************************************************************************************
Public Sub RegSave(ControlToRemember As Control, ControlValue As Variant)
On Error Resume Next
Dim ControlIndex As String
    With ControlToRemember
        ControlIndex = .Index
        Call SaveSetting(App.EXEName, "Settings", .Parent.Name & "_" & .Name & "_" & ControlIndex, ControlValue)
    End With
End Sub

Public Sub RegSaveString(str1 As String, str2 As String)
    Call SaveSetting(App.EXEName, "Settings", str1, str2)
End Sub

Public Function RegLoadString(str1 As String)
    RegLoadString = GetSetting(App.EXEName, "Settings", str1)
End Function

Public Sub SaveFormPosition(NameOfForm As Form)
On Error Resume Next
Dim FormIndex As String
    With NameOfForm
        FormIndex = .Index     'If there is no form index, resume next line of execution
        Call SaveSetting(App.EXEName, "Settings", .Name & "_" & FormIndex & "_Position", .Left & "*" & .Top & "*" & .Width & "*" & .Height)
    End With
End Sub

