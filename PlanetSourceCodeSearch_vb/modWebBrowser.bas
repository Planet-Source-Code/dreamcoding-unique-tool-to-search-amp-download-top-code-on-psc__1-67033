Attribute VB_Name = "modWebBrowser"
 Private Const SW_SHOW = 5       ' Displays Window in its current size
                                      ' and position
      Private Const SW_SHOWNORMAL = 1 ' Restores Window if Minimized or
                                      ' Maximized

      Private Declare Function ShellExecute Lib "shell32.dll" Alias _
         "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
         String, ByVal lpFile As String, ByVal lpParameters As String, _
         ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

      Private Declare Function FindExecutable Lib "shell32.dll" Alias _
         "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
         String, ByVal lpResult As String) As Long
Public gCurrentBrowserURL As String
Public Function GetBrowserURL(dummyLabel As Label, strLimitURL As String)
On Error Resume Next 'do not remove

'This procedure can not be ported to .net

'If you are only searching for URLS with a particular string in them
'then specify the string using the variable strLimitURL

'We need to put at least one label on the form to get our URL so this is dummyLabel

   Dim strBrowser       As String
   Dim strBrowserURL    As String
    
    strBrowser = Replace(GetDefaultBrowser, vbNewLine, "")
    strBrowserURL = dummyLabel.Caption
    
    'Try different browsers in case user has opened non-default
    'Type in the name of your browsers exename to manually retrieve current URL
    'Program will attempt to find your default browser's current URL

   
    
    'IE
    With dummyLabel
      .Caption = ""
      .AutoSize = True
      .LinkTopic = "IExplore|WWW_GetWindowInfo"
      .LinkItem = "0xffffffff"
      .LinkMode = 2
      .LinkRequest
    End With
    If dummyLabel.Caption <> vbNullString Then
       'We only want the PSC code page for purposes of this program
       If InStr(1, dummyLabel.Caption, strLimitURL) Then
            TempURL = dummyLabel.Caption
       End If
    End If
   
    'FireFox
    With dummyLabel
      .Caption = ""
      .AutoSize = True
      .LinkTopic = "FireFox|WWW_GetWindowInfo"
      .LinkItem = "0xffffffff"
      .LinkMode = 2
      .LinkRequest
    End With
    If dummyLabel.Caption <> vbNullString Then
        If InStr(1, dummyLabel.Caption, strLimitURL) Then
            TempURL = dummyLabel.Caption
        End If
    End If
    
    'Opera
    With dummyLabel
      .Caption = ""
      .AutoSize = True
      .LinkTopic = "Opera|WWW_GetWindowInfo"
      .LinkItem = "0xffffffff"
      .LinkMode = 2
      .LinkRequest
    End With
     
     'Use Default
   With dummyLabel
      .Caption = ""
      .AutoSize = True
      .LinkTopic = strBrowser & "|WWW_GetWindowInfo"
      .LinkItem = "0xffffffff"
      .LinkMode = 2
      .LinkRequest
   End With
   If dummyLabel.Caption <> vbNullString Then
        If InStr(1, dummyLabel.Caption, strLimitURL) Then
            TempURL = dummyLabel.Caption
        End If
    End If
    If InStr(1, dummyLabel.Caption, strLimitURL) Then
        TempURL = dummyLabel.Caption
    End If
    TempURL = Split(TempURL, Chr$(34))(1)
    TempURL = Replace(TempURL, Chr$(34), vbNullString)
GetBrowserURL = TempURL
End Function
         
Public Function GetDefaultBrowser()
Dim FileName As String, Dummy As String
      Dim BrowserExec As String * 255
      Dim RetVal As Long
      Dim FileNumber As Integer
      Dim FileTitle As String
      
      ' First, create a known, temporary HTML file
      BrowserExec = Space(255)
      FileName = App.Path & "\temphtm.HTM"
      FileNumber = FreeFile                    ' Get unused file number
      Open FileName For Output As #FileNumber  ' Create temp HTML file
          Write #FileNumber, "<HTML> <\HTML>"  ' Output text
      Close #FileNumber                        ' Close file
      ' Then find the application associated with it
      RetVal = FindExecutable(FileName, Dummy, BrowserExec)
      BrowserExec = Trim(BrowserExec)
      
      ' If an application is found, launch it!
      If RetVal <= 32 Or IsEmpty(BrowserExec) Then ' Error
          MsgBox "Could not find associated Browser", vbExclamation, _
            "Browser Not Found"
      Else
          FileTitle = Replace(Trim(UCase(Right(BrowserExec, Len(BrowserExec) - InStrRev(BrowserExec, "\")))), ".EXE", vbNullString)
          If RetVal <= 32 Then        ' Error
              MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
          End If
      End If
      Kill FileName                   ' delete temp HTML file
      GetDefaultBrowser = TrimAll(FileTitle)
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
Public Function ReplaceString(sExpression As String, sFind As String, sReplace As String) As String
    ' Title: Replace
    ' Version: 1.01
    ' Author: Leigh Bowers
    ' WWW:http://www.esheep.freeserve.co.uk/
    '     compulsion
    Dim lPos As Long
    Dim iFindLength As Integer
    ' Ensure we have all required parameters
    '


    If Len(sExpression) = 0 Or Len(sFind) = 0 Then
        Exit Function
    End If
    
    ' Determine the length of the sFind vari
    '     able
    iFindLength = Len(sFind)
    
    ' Find the first instance of sFind
    
    lPos = InStr(sExpression, sFind)
    
    ' Process and find all subsequent instan
    '     ces
    


    Do Until lPos = 0
        sExpression = sExpression & Left$(sExpression, lPos - 1) + sReplace + Mid$(sExpression, lPos + iFindLength)
        lPos = InStr(lPos, sExpression, sFind)
    Loop
    
    ' Return the result
    ReplaceString = sExpression
End Function


