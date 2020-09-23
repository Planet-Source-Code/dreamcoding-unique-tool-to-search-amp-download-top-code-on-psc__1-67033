Attribute VB_Name = "modPSCQuickSearch"
Option Explicit

'Results
Public PageNum                  As Long         'Results Page Number

'PSC Login
Public Type PSC_OPTIONS
    PSCUserName                 As String   'PSC Username (Encrypted in memory)
    PSCPassword                 As String   'PSC Password (Encrypted in memory)
End Type

'Encode URLS
Private enc As New clsCoder

Public PSCOpt As PSC_OPTIONS

'Retrieve and Parse PSC
Public objXmlWeb As clsXMLWeb
Public Type varPSCResults
    CodeAuthorName          As String
    CodeSubmittedOn         As String
    CodeLevel               As String
    CodeUserRating          As String
    CodeAccessed            As String
    CodeCompatibility       As String
    CodeAuthorID            As String
    CodeAuthorPhoto()       As String
    CodeAuthorPlacement     As Long
    CodeAuthorWord          As String
    CodeTitle               As String
    CodeDescription         As String
    CodeViews               As String
    CodeUsersVoted          As String
    CodeURL                 As String
    DownloadCodeNames       As String
    ResultNum               As Long
End Type
Public PSCResults As varPSCResults


Public Function LoginURL(ReturnURL As String, strUsername As String, strPassword As String)
Dim strLoginURL         As String
Dim strTxtReturnURL     As String
Dim strLoginInfo        As String


Set enc = New clsCoder
    
    strLoginURL = "https://www.rentacoder.com/ads/authentication/LoginAction.asp?txtReturnURL="
    strTxtReturnURL = enc.sURLEncode(ReturnURL)
    strLoginInfo = "&lngWId=1&blnOutsideOfVBSubWeb=True&txtEmailAddress=" & strUsername & "&txtPassword=" & strPassword
    LoginURL = strLoginURL & strTxtReturnURL & strLoginInfo

End Function
Public Sub Pause(NbSec As Single)
 Dim Finish As Single
 Finish = Timer + NbSec
 DoEvents
 Do Until Timer >= Finish
 Loop
End Sub

Public Sub ParsePSC(OurPage As String)

Dim Page                    As String            'HTML Code of Page Returned
Dim strAuthLevel            As String            'Code Level
Dim strViewsSubmitted       As String            'Code Views
Dim arrDesParse()           As String            'Code Description
Dim strUserRating           As String            'Code User Rating
Dim strUserVoted            As String            'Code Users Voted
Dim strUserAmount           As String            'Users voted tallied
Dim strUserExcellents       As String            'Excellent votes
Dim strDesParse             As String            'Description
Dim strURL                  As String            'URL for Code's Page
Dim strPSCBalls             As String            'Vote Graphics
Dim arrFullBalls()          As String            'Full Vote Graphic
Dim arrHalfBalls()          As String            'Half Vote Graphic
Dim intFullBalls            As Integer
Dim intHalfBalls            As Integer
Dim X                       As Integer
Dim strFull                 As String
Dim column_header As ColumnHeader
Dim list_item As ListItem
'                                                'Variables below may be used in future (Keep)
'Dim AuthPhoto()             As String           'Author Photo Array
'Dim AuthorPlacement         As Long             'Author Placement
'Dim strAuthPhoto            As String           'Author Photo
'Dim arrEntries()            As String




    If PSCResults.ResultNum = 0 Then PSCResults.ResultNum = 1
    
    'Quick Parse will parse the HTML page between two given strings
    'to get the text we want, we remove the HTML
    strAuthLevel = QuickParse(PSCResults.ResultNum, OurPage, "<!--level-->", "<!--views/date submitted-->", False, True)
    strURL = QuickParse(PSCResults.ResultNum, OurPage, "<!--descrip-->", Chr$(34) & ">", False, False)
    If strURL <> vbNullString Then
        strURL = Split(strURL, "=" & Chr$(34))(1)
        strURL = "http://www.planet-source-code.com" & strURL
        PSCResults.CodeURL = strURL
    End If
    'Since Author & Code Level are one string seperated by slash
    'split this string at the slash mark to get each
    '(simliar parsing techniques used below)
    If strAuthLevel <> vbNullString Then
        PSCResults.CodeAuthorName = RemoveHTML(Split(strAuthLevel, "/")(1))
        PSCResults.CodeLevel = RemoveHTML(Split(strAuthLevel, "/")(0))
    End If
    PSCResults.CodeTitle = RemoveHTML(QuickParse(PSCResults.ResultNum, OurPage, "<!--descrip-->", "</TD>", False, True))
    strViewsSubmitted = QuickParse(PSCResults.ResultNum, OurPage, "<!--views/date submitted-->", "</TD>", False, True)
    If strViewsSubmitted <> vbNullString Then
        PSCResults.CodeViews = Split(strViewsSubmitted, " since")(0)
        PSCResults.CodeSubmittedOn = Split(strViewsSubmitted, " since")(1)
    End If
    strUserRating = QuickParse(PSCResults.ResultNum, OurPage, "<!--user rating-->", "<!description>", False, True)
    If strUserRating <> vbNullString Then
        strUserAmount = Trim(Replace(Split(strUserRating, "Users")(0), "By ", ""))
        If Not strUserRating = "Unrated" Then
            strUserExcellents = Split(strUserRating, "Users")(1)
            strUserExcellents = Replace(strUserExcellents, " Excellent Ratings", "")
        End If
            If strUserExcellents = "" Then strUserExcellents = "0"
            PSCResults.CodeUsersVoted = strUserAmount
            PSCResults.CodeUserRating = strUserExcellents
    End If
    
    PSCResults.CodeCompatibility = QuickParse(PSCResults.ResultNum, OurPage, "<!--code compat-->", "</TD>", False, True)
    
    strDesParse = QuickParse(PSCResults.ResultNum, OurPage, "<!description>", "<a href=", True, True)
    If strDesParse <> vbNullString Then
        arrDesParse = Split(strDesParse, UniqueToken)
        PSCResults.CodeDescription = arrDesParse(1)
        PSCResults.ResultNum = arrDesParse(0) + 1
    End If
    frmResults.List1.AddItem strFull & " " & PSCResults.CodeTitle & " " & PSCResults.CodeDescription
    frmResults.lstCodeURL.AddItem PSCResults.CodeURL    'Store into listbox
    
    ' frmResults.lvwResults.Icons = frmResults.imgSmall
     Set list_item = frmResults.lvwResults.ListItems.Add(, , PSCResults.CodeTitle, , 1)
    list_item.SubItems(1) = PSCResults.CodeDescription
    list_item.SubItems(2) = PSCResults.CodeAuthorName
    list_item.SubItems(3) = PSCResults.CodeUserRating
    list_item.SubItems(4) = PSCResults.CodeUsersVoted
    list_item.SubItems(5) = PSCResults.CodeLevel
    list_item.SubItems(6) = PSCResults.CodeSubmittedOn
    list_item.SubItems(7) = PSCResults.CodeViews
    
 
    'Show the graphic based on amount of votes
    
'CODE FOR ALTERNATE PSC SEARCH PAGE --- Leave this for reference, useful for future development
'
'PSCResults.CodeAuthorName = QuickParse(1, OurPage, "By:</b>", "<BR>", False, True)
'PSCResults.CodeSubmittedOn = QuickParse(1, OurPage, "<b>Submitted on:</b>", "<BR>", False, True)
'PSCResults.CodeLevel = QuickParse(1, OurPage, "Level:", "<BR>", False, True)
'PSCResults.CodeUserRating = QuickParse(1, OurPage, "User Rating:</b>", "<BR>", False, True)
'PSCResults.CodeCompatibility = QuickParse(1, OurPage, "Compatibility:</b>", "<BR>", False, True)
'PSCResults.CodeAccessed = QuickParse(1, OurPage, "Users have accessed this&nbsp;code&nbsp;", "times.<BR>", False, True)
'PSCResults.CodeAuthorID = QuickParse(1, OurPage, "lngAuthorId=", "&", False, True)
'strAuthPhoto = QuickParse(1, OurPage, "AUTHOR_PHOTO", ".", True, True)
'If strAuthPhoto <> vbNullString Then
'PSCResults.CodeAuthorPhoto = Split(strAuthPhoto, UniqueToken)
'If UBound(CodeAuthorPhoto) > 0 Then
'    PSCResults.CodeAuthorPlacement = CodeAuthorPhoto(0)
'    PSCResults.CodeAuthorWord = CodeAuthorPhoto(1)
'End If
'End If
'PSCResults.CodeDescription = QuickParse(PSCResults.CodeAuthorPlacement, OurPage, "&nbsp;", "<BR>", False, True)
'

End Sub

Public Sub UnloadAllForms()

Dim f As Form

For Each f In Forms
    Unload f
Next

End Sub
