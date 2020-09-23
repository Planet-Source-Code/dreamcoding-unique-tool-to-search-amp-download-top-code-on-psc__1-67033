VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFavorites 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Favorites"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6135
   Icon            =   "frmFavorites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdRemoveAuthor 
      Height          =   315
      Left            =   2640
      TabIndex        =   9
      Top             =   4500
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      Caption         =   "Remove Author"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmFavorites.frx":014A
      cBack           =   -2147483633
   End
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdAddAuthor 
      Height          =   315
      Left            =   1260
      TabIndex        =   8
      Top             =   4500
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "Add Author"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmFavorites.frx":02F4
      cBack           =   -2147483633
   End
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdViewPage 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   4500
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "View Page"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1170
      ScaleHeight     =   195
      ScaleWidth      =   2145
      TabIndex        =   5
      Top             =   450
      Width           =   2145
      Begin VB.Label lblStatus 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Height          =   345
         Left            =   0
         TabIndex        =   6
         Top             =   -30
         Width           =   1860
      End
   End
   Begin VB.ComboBox cmbAuthorCodeCat 
      Height          =   315
      Left            =   4305
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4485
      Width           =   1560
   End
   Begin VB.ListBox lstAuthorIDs 
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.ListBox lstAuthors 
      Height          =   3765
      Left            =   150
      TabIndex        =   1
      Top             =   645
      Width           =   5700
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4860
      Left            =   15
      TabIndex        =   0
      Top             =   90
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8573
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Authors"
            Key             =   "tabAuthor"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblBrowserURL 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   3150
      TabIndex        =   2
      Top             =   15
      Visible         =   0   'False
      Width           =   1260
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XMLGet           As clsXMLWeb        'Class using XML to Grab Web Pages
Private strAuthorName      As String           'AuthorName
Private strAuthorID        As String           'AuthorID
Private strMaxEntries      As String           'Max Entries
Private AuthorCatID        As Integer          'Author Category ID
Private iniSettingsFile    As String           'INI File Path
Private Relogin            As Boolean

Private Sub cmbAuthorCodeCat_Click()
Select Case cmbAuthorCodeCat.ListIndex
    Case 0
    '.Net
    AuthorCatID = 10
    Case 1:
    'ASP / VBScript
    AuthorCatID = 4
    Case 2:
    'C / C++
    AuthorCatID = 3
    Case 3:
    'Cold Fusion
    AuthorCatID = 9
    Case 4:
    'Delphi
    AuthorCatID = 7
    Case 5:
    'Java / Javascript
    AuthorCatID = 2
    Case 6:
    'Perl
    AuthorCatID = 6
    Case 7:
    'PHP
    AuthorCatID = 8
    Case 8:
    'SQL
    AuthorCatID = 5
    Case 9:
    'Visual Basic
    AuthorCatID = 1
    Case 10:
    'All
    AuthorCatID = 0
End Select
End Sub

Private Sub cmdAddAuthor_Click()
On Error GoTo command_error

   Dim TempURL          As String
   Dim strBrowser       As String
   Dim strBrowserURL    As String
   Dim strHTML          As String
   
    
    strBrowser = Replace(GetDefaultBrowser, vbNewLine, "")
    strBrowserURL = lblBrowserURL.Caption
    
    'Try different browsers in case user has opened non-default
    'Type in the name of your browsers exename to manually retrieve current URL
    'Program will attempt to find your default browser's current URL
GoTo skip
    'Use Default
   With lblBrowserURL
      .Caption = ""
      .AutoSize = True
      .LinkTopic = strBrowser & "|WWW_GetWindowInfo"
      .LinkItem = "0xffffffff"
      .LinkMode = 2
      .LinkRequest
   End With
   If lblBrowserURL.Caption <> vbNullString Then
        If InStr(1, lblBrowserURL.Caption, "ShowCode.asp") Then
            TempURL = lblBrowserURL.Caption
        End If
    End If
    
    'IE
    With lblBrowserURL
      .Caption = ""
      .AutoSize = True
      .LinkTopic = "IExplore|WWW_GetWindowInfo"
      .LinkItem = "0xffffffff"
      .LinkMode = 2
      .LinkRequest
    End With
    If lblBrowserURL.Caption <> vbNullString Then
       'We only want the PSC code page for purposes of this program
       If InStr(1, lblBrowserURL.Caption, "ShowCode.asp") Then
            TempURL = lblBrowserURL.Caption
       End If
    End If
   
    'FireFox
    With lblBrowserURL
      .Caption = ""
      .AutoSize = True
      .LinkTopic = "FireFox|WWW_GetWindowInfo"
      .LinkItem = "0xffffffff"
      .LinkMode = 2
      .LinkRequest
    End With
    If lblBrowserURL.Caption <> vbNullString Then
        If InStr(1, lblBrowserURL.Caption, "ShowCode.asp") Then
            TempURL = lblBrowserURL.Caption
        End If
    End If
    
    'Opera
    With lblBrowserURL
      .Caption = ""
      .AutoSize = True
      .LinkTopic = "Opera|WWW_GetWindowInfo"
      .LinkItem = "0xffffffff"
      .LinkMode = 2
      .LinkRequest
    End With
    If InStr(1, lblBrowserURL.Caption, "ShowCode.asp") Then
        TempURL = lblBrowserURL.Caption
    End If
skip:
TempURL = GetBrowserURL(Me.lblBrowserURL, "ShowCode.asp")

If InStr(1, TempURL, "ShowCode.asp") Then
   TempURL = Split(TempURL, Chr$(34))(1)
   If TempURL <> vbNullString Then
        'Retrieve a copy of source to this page
        strHTML = XMLGet.GetDataAsString(TempURL)
        'Now look for the author's code on this page
        'strAuthorName lngAuthorID
        strAuthorName = QuickParse(1, strHTML, "strAuthorName", "&", False, True)
        strAuthorName = Replace(strAuthorName, "%20", " ")  'change %20 to space
        strAuthorName = Replace(strAuthorName, "=", vbNullString) 'Trailing character removed
        strAuthorID = QuickParse(1, strHTML, "lngAuthorID", "&", False, True) 'Remove & sign
        strAuthorID = Replace(strAuthorID, "lngAuthorId=", vbNullString)    'Remove URL variable
        'strcodecatid = quickparse(1,
        'Now that we have the AuthorID & AuthorName we can utilize the
        'PSC Author web search
        '(PSC does not allow you to search for AuthorName with out the ID, would be nice though)
        If modListboxControl.IsDuplicate(strAuthorID, Me.lstAuthorIDs) = False Then
            'Add Unique AuthorName & ID to listboxes
            If strAuthorName <> vbNullString And strAuthorID <> vbNullString Then
                lstAuthors.AddItem strAuthorName
                lstAuthorIDs.AddItem strAuthorID
            End If
            'Store the AuthorName & ID as a favorite
            'For now using 2 listboxes to save
            ListSave App.Path & "\AuthorNames.lst", Me.lstAuthors
            ListSave App.Path & "\AuthorIDs.lst", Me.lstAuthorIDs
        Else
            lblStatus.Caption = "Author Already Listed!"
            Pause 1
            lblStatus.Caption = vbNullString
        End If
            lblStatus.Caption = "Author Added!"
            Pause 1
            lblStatus.Caption = vbNullString

   End If
Else
    MsgBox "You must navigate to a code download page on PSC before the author is added."
End If
   Exit Sub

command_error:
'MsgBox Err.Description
   'try the next step on error
    Resume Next

End Sub

Private Sub cmdRemoveAuthor_Click()
        If lstAuthors.ListCount > 0 And lstAuthors.ListIndex > -1 Then
            'Remove associated AuthorID
            lstAuthorIDs.RemoveItem lstAuthors.ListIndex
            'Remove AuthorName from List
            lstAuthors.RemoveItem lstAuthors.ListIndex
            'Store the AuthorName & ID as a favorite
            'For now using 2 listboxes to save
            ListSave App.Path & "\AuthorNames.lst", Me.lstAuthors
            ListSave App.Path & "\AuthorIDs.lst", Me.lstAuthorIDs
        End If
End Sub

Private Sub cmdViewPage_Click()
ViewAuthorCodeBank
End Sub

Private Sub Form_Load()
 Dim posTop     As String
 Dim posLeft    As String
 
 Set XMLGet = New clsXMLWeb              'Create Instance of Class to be used as Object
 'Load Settings Using INI
 iniSettingsFile = App.Path & "\settings.ini"
 
 'Form Position
 posTop = ReadIni(iniSettingsFile, "FormPosition", "frmFavorites_Top")
 posLeft = ReadIni(iniSettingsFile, "FormPosition", "frmFavorites_Left")

 If posTop > vbNullString And posLeft > vbNullString Then
     Me.Top = CLng(posTop)
     Me.Left = CLng(posLeft)
 End If
 'Set the form on top
 Win_ControlOnTop Me, True
 
 'Load Listboxes from Saved Files
 ListLoad App.Path & "\AuthorNames.lst", Me.lstAuthors
 ListLoad App.Path & "\AuthorIDs.lst", Me.lstAuthorIDs
 'Load Listbox with Categorys
 cmbAuthorCodeCat.AddItem ".Net"
 cmbAuthorCodeCat.AddItem "ASP / VBScript"
 cmbAuthorCodeCat.AddItem "C / C++"
 cmbAuthorCodeCat.AddItem "Cold Fusion"
 cmbAuthorCodeCat.AddItem "Delphi"
 cmbAuthorCodeCat.AddItem "Java / Javascript"
 cmbAuthorCodeCat.AddItem "Perl"
 cmbAuthorCodeCat.AddItem "PHP"
 cmbAuthorCodeCat.AddItem "SQL"
 cmbAuthorCodeCat.AddItem "Visual Basic"
 cmbAuthorCodeCat.AddItem "All"
 If cmbAuthorCodeCat.ListIndex = -1 Then cmbAuthorCodeCat.ListIndex = 10
 tmpIndex = ReadIni(iniSettingsFile, "ControlSetting", "frmFavorites_cmbAuthorCodeCat")
 If tmpIndex <> vbNullString Then
 cmbAuthorCodeCat.ListIndex = Abs(CLng(tmpIndex))
 End If
 

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Use INI File to Save Settings
WriteIni iniSettingsFile, "ControlSetting", "frmFavorites_cmbAuthorCodeCat", CStr(cmbAuthorCodeCat.ListIndex)
WriteIni iniSettingsFile, "FormPosition", "frmFavorites_Top", Me.Top
WriteIni iniSettingsFile, "FormPosition", "frmFavorites_Left", Me.Left
End Sub

Private Sub lstAuthors_DblClick()
ViewAuthorCodeBank
End Sub
Private Sub ViewAuthorCodeBank()
Dim MaxEntries          As String
Dim strAuthorCodeCat    As String
Dim strUsername         As String
Dim strPassword         As String
Dim URL                 As String
Dim strLoginNeeded      As String
Dim lngRetainLogin      As Long
Dim enc                 As clsCoder

Set enc = New clsCoder

If lstAuthors.ListCount > 0 Then
    If lstAuthors.ListIndex > -1 Then
        MaxEntries = frmMain.cmbMaxEntry.List(frmMain.cmbMaxEntry.ListIndex)
        strAuthorName = enc.sURLEncode(Me.lstAuthors.List(lstAuthors.ListIndex))
        strAuthorID = Me.lstAuthorIDs.List(lstAuthors.ListIndex)
        strAuthorCodeCat = CStr(AuthorCatID)
        
        'If ALL categorys selected then category web variable = nothing
        If AuthorCatID = 0 Then
            strAuthorCodeCat = "1"
            Relogin = True
        End If
            'http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=2210060021&strAuthorName=LaVolpe&txtMaxNumberOfEntriesPerPage=25
            'Load the Author's Page of Code
            URL = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=" & strAuthorCodeCat & "&blnAuthorSearch=TRUE&lngAuthorId=" & strAuthorID & "&strAuthorName=" & strAuthorName & "&txtMaxNumberOfEntriesPerPage=" & MaxEntries
            strURL = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=&blnAuthorSearch=TRUE&lngAuthorId=" & strAuthorID & "&strAuthorName=" & strAuthorName & "&txtMaxNumberOfEntriesPerPage=" & MaxEntries
            'Open web page through browser
            strPassword = PSCOpt.PSCPassword
            strUsername = PSCOpt.PSCUserName
            strPassword = modRC4.DecryptString(strPassword, "OurKey", True)
            strUsername = modRC4.DecryptString(strUsername, "OurKey", True)
            strUsername = strUsername
            lngRetainLogin = frmOptions.chkRetainLogin.Value
            
                    
            If lngRetainLogin = 1 Then
                URL = LoginURL(URL, strUsername, strPassword)
            End If
            ShellExecute Me.hwnd, "open", URL, vbNullString, vbNullString, SW_SHOWNORMAL
            
            If Relogin = True Then
            
                'Some reason PSC will not show all author code unless you are first logged into some sort of area
                'so we log into VB area first (lngWid=1) and now launch URL with lngWID as blank
                'strURL has blank lngwID (see above)
                
                'Now we loop until we know the first URL is loaded
                Do
                    strLoginNeeded = GetBrowserURL(Me.lblBrowserURL, "")
                    DoEvents
                Loop Until strLoginNeeded <> vbNullString And strLoginNeeded <> "about:blank"
                
                If InStr(1, strLoginNeeded, "lngWId=1") Then
                ShellExecute Me.hwnd, "open", strURL, vbNullString, vbNullString, SW_SHOWNORMAL
               
                    Do
                        strLoginNeeded = GetBrowserURL(Me.lblBrowserURL, "lngWId=&")
                        DoEvents
                    Loop Until strLoginNeeded <> vbNullString
                End If
               ' ShellExecute Me.hwnd, "open", strURL, vbNullString, vbNullString, SW_SHOWNORMAL
                Relogin = False
            End If
            
            If InStr(1, strLoginNeeded, "http://www.planet-source-code.com/vb/timeout/SessionTimeout.asp") Then
                'If the "sessiontimeout.asp" page is detected then they need to log in
                If lngRetainLogin = 0 Then
                    'No
                    MsgBox "You must log-in to PSC. This program can auto-login using the Retain Login Option in Options"
                End If
            End If
    End If
End If
End Sub
