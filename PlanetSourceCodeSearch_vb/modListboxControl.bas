Attribute VB_Name = "modListboxControl"


Sub ListSave(Path As String, lst As ListBox)
'Ex: Call Save_ListBox("c:\windows\desktop\list.lst", list1)

    Dim Listz As Long
    On Error Resume Next
    Dim f As Long
    Dim x As Long
    
    f = FreeFile
    Open Path$ For Output As f
    For x = 0 To lst.ListCount - 1
        Print #1, lst.List(x)
    Next
    Close f
End Sub

Sub ListLoad(Path As String, lst As ListBox)
'Ex: Call Load_ListBox("c:\windows\desktop\list.lst", list1)

    
    On Error Resume Next
    Dim FileText As String
    Dim f As Long
    f = FreeFile
    Open Path$ For Input As f
    Do
        Input #1, FileText$
        DoEvents
        lst.AddItem FileText$
    Loop Until EOF(1)
    Close f
End Sub
Public Function IsDuplicate(Item As String, lst As ListBox) As Boolean
'Is string item a duplicate in list?

x = 0
If lst.ListCount > 0 Then

 Do
  If Item = lst.List(x) Then
    'This is a duplicate
  IsDuplicate = True
  Exit Function
  End If
 x = x + 1
 Loop Until x = lst.ListCount

IsDuplicate = False

Else
IsDuplicate = False

End If

End Function
