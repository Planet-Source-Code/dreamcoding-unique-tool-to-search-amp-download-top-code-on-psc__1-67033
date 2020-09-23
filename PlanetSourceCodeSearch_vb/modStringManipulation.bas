Attribute VB_Name = "modStringManipulation"
'This module is needs to be re-coded
'but for now, it gets the job done.


Public Function QuickParse(InitialPlace As Long, Page As String, Startword As String, Endword As String, ReturnPlacement As Boolean, RemoveHTMLPage As Boolean)
 
 Dim TheStart As Long
 Dim Page2 As String
 Dim FoundAt As String
 Dim TheStart1 As Long
 Dim TheStart2 As Long
 Dim EndWordCheck As Long
 Dim PageLength As Long
 PageLength = Len(Page)
 
 'Can start initial place of parsing as further down the page
 'to help narrow down possibility of picking up the wrong text parsed
 
 If InitialPlace = 0 Then
  InitialPlace = 1
 End If
 
 TheStart1 = InStr(InitialPlace, UCase(Page), UCase(Startword), vbTextCompare)
    
    If TheStart1 <> 0 Then
     Page2 = Mid(Page, TheStart1, Len(Page) - TheStart1)
     
     'Check for endword place, if not there we will just take all chr's to the right
     EndWordCheck = InStr(TheStart1, Page, Endword, vbTextCompare)
     If EndWordCheck = 0 Then
     Page2 = Right(Page, Len(Page) - TheStart1 - Len(Startword) + 1)
     GoTo Done
     End If
     
     TheStart2 = InStr(1, UCase(Page2), UCase(Endword), vbTextCompare)
      
     If TheStart2 = 0 Then TheStart2 = 1
     Page2 = Left(Page2, TheStart2 - 1)
      
      If ReturnPlacement = True Then
       FoundAt = TheStart1 & UniqueToken
       If RemoveHTMLPage = True Then
         QuickParse = FoundAt & TrimAll(RemoveHTML(Replace(Page2, Startword, "")))
       Else
         QuickParse = FoundAt & TrimAll(Replace(Page2, Startword, ""))
       End If
       
       Else
       
       If RemoveHTMLPage = True Then
       QuickParse = TrimAll(RemoveHTML(Replace(Page2, Startword, "")))
       Else
       QuickParse = Page2
       End If
      End If
     
    Else
     QuickParse = ""
    
    End If
    Exit Function
Done:
    QuickParse = Page2
End Function
Public Function RemoveHTML(ByVal str As String) As String
    ' This function finds and deletes the HTML tags
    ' To clean a tag from start to end i.e. <a html=...> ... </a> use the function CleanSHTMLTags

    Dim s, l As Long 'variables to hold the start position and the length of HTML tags found in the current string
    
    ' Not a bad idea to first normalize the code passed, just in case...
    ' Moron NormalizeTags...
    str = NormalizeTags(str)
    
    If IsNull(str) Then
        str = "no text to clean from HTML tags"
        Exit Function
    End If
    
    Do While l >= 0
        str = DeleteString(str, s, l)
        FindHTMLTag str, s, l
    Loop
    str = Replace(str, "&nbsp;", "")
    str = CleanDoubleSpaces(str)
    RemoveHTML = str
End Function

Public Function TrimAll(ToTrim As String) As String

    Dim start, Finish As Integer
    Dim ToEliminate As String

    ' Base condition test
    If Len(ToTrim) = 0 Then
        TrimAll = ""
        Exit Function
    End If

    ' Define the characters that we want to trim off
    ToEliminate = Chr(0) & Chr(8) & Chr(9) & Chr(10) & Chr(13) & Chr(32)

    ' Find the beginning of non-blank string
    start = 1
    While InStr(1, ToEliminate, Mid$(ToTrim, start, 1), vbTextCompare) <> 0 And start <= Len(ToTrim)
        start = start + 1
    Wend

    ' Find the end of non-blank string
    Finish = Len(ToTrim)
    While InStr(1, ToEliminate, Mid$(ToTrim, Finish, 1), vbTextCompare) <> 0 And Finish > 1
        Finish = Finish - 1
    Wend
    '
    ' If the string is completely blank, Start is going to be greater
    ' than Finish
    '
    If start > Finish Then
        TrimAll = ""
        Exit Function
    End If

    ' Trim out the real contents
    TrimAll = Mid$(ToTrim, start, Finish - start + 1)

End Function

Function UniqueToken()
'To avoid token character which may cause a bad parse
'Use whatever you think is best
UniqueToken = Chr$(1) & Chr$(2)
End Function

Public Function NormalizeTags(ByVal str As String) As String
    ' this function eliminates the spaces from the tag commands
    ' < script ...> will become <script ...>
    ' the main usage is Normalizing all the HTML string before looking for specific tags i.e. <title>
    
    ' It is not a bad idea to get rid of multiple space before normalization
    str = CleanDoubleSpaces(str)
    
    Do While InStr(1, str, "< ") > 0 ' continue as long as there are spaces between "<" and the tag commands
        str = Replace(str, "< ", "<")
    Loop
    
    NormalizeTags = str
End Function
Public Function CleanRepeatedAscii(ByVal str As String, ByVal ac As Long) As String
    ' This function simply gets rid of the repeated ascii charactes...
    ' It really helps after filtering an HTML code since lots of line breaks will be left
    ' Also double spaces are eliminated this way
    ' str is the HTML code, and ac is the repeated ascii code to be cleaned
    Dim s, l, ss, ls As Long
    Dim ch, ch2 As String
    
    ch = Chr(ac)
    ch2 = ch + ch
    
    If IsNull(str) Or IsEmpty(str) Then
        str = "no text to clean from specified Ascii code"
        Exit Function
    End If
    
    Do While True
        ss = InStr(str, ch2)
        If ss > 0 Then
            str = Replace(str, ch2, ch)
        Else
            Exit Do
        End If
    Loop
    CleanRepeatedAscii = str
End Function

Public Function CleanDoubleSpaces(ByVal str As String) As String
    ' Mostly after cleaning tags, double and more spaced will remain in the text
    ' this function simply gets rid of them
    
    str = CleanRepeatedAscii(str, "32")

    CleanDoubleSpaces = str
End Function
Public Function DeleteString(ByVal str As String, ByVal s As Long, ByVal l As Long) As String
    ' This function deletes the part of a string indicated by
    ' s, the starting position
    ' l, the length of the string to be deleted
    If l = 0 Then
        DeleteString = str
        Exit Function
    End If
    DeleteString = Left(str, s) + Right(str, Len(str) - s - l)
End Function
Public Sub FindHTMLTag(ByVal MainStr As String, ByRef s, ByRef l, Optional ByVal ss As Long = 0)
    ' This function finds an HTML tag in a string and puts the start position and the length of the tag in s and l respectively
    ' the HTML tags are the tags that do contain text that is displayed on your browser
    ' however, tags like <script> .... </script> containg data that should be filtered as a whole
    ' for that purpose, use FindSHTMLTag, which will get a tag from start to end
    
    ' MainStr is the string containing HTML source
    ' s is a return value that represents the position where the HTML tag starts
    ' l is the  lenght of the HTML text
    ' ss is the start string position, by default it points to the start position of the string
    ' if no tag is found l will return with a negative value...
    ' to ensure proper operation check the value of l
    
    Dim sl, ll As Long
    
    If IsNull(MainStr) Or Len(MainStr) < 3 Then 'the string passed can not be a tag at all
        s = 0
        l = -1
        Exit Sub
    End If
    
    sl = InStr(ss + 1, MainStr, "<")
    ll = InStr(sl + 1, MainStr, ">")
    
    If sl = 0 Or ll = 0 Then
        s = 0
        l = -2
        Exit Sub
    End If
    
    s = sl - 1
    l = ll - s
    
End Sub


