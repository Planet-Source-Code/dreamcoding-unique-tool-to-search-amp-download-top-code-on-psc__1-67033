VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PSC Quick Search"
   ClientHeight    =   1770
   ClientLeft      =   7560
   ClientTop       =   6780
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdMiniMode 
      Height          =   315
      Left            =   315
      TabIndex        =   15
      ToolTipText     =   "Mini Mode"
      Top             =   1380
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   556
      Caption         =   "m"
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
   Begin VB.CheckBox chkDownload 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Download"
      Height          =   195
      Left            =   3255
      TabIndex        =   14
      Top             =   1020
      Width           =   1050
   End
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdIcon 
      Height          =   315
      Left            =   45
      TabIndex        =   13
      ToolTipText     =   "Minimize Systray"
      Top             =   1380
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      Caption         =   "i"
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
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdOptions 
      Height          =   330
      Left            =   630
      TabIndex        =   12
      Top             =   1365
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   582
      Caption         =   "Options"
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
      Image           =   "frmMain.frx":038A
      cBack           =   -2147483633
   End
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdFavorites 
      Height          =   330
      Left            =   3390
      TabIndex        =   9
      Top             =   1350
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      Caption         =   "&Favorites"
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
      Image           =   "frmMain.frx":0571
      cBack           =   -2147483633
   End
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdNext 
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   1365
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   582
      Caption         =   "&Next"
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
      ImgAlign        =   2
      Image           =   "frmMain.frx":073C
      cBack           =   -2147483633
   End
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdPrev 
      Height          =   330
      Left            =   1635
      TabIndex        =   2
      Top             =   1365
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   582
      Caption         =   "&Prev"
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
      Image           =   "frmMain.frx":08ED
      cBack           =   -2147483633
   End
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdSearch 
      Default         =   -1  'True
      Height          =   360
      Left            =   3495
      TabIndex        =   1
      Top             =   60
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   635
      Caption         =   "&Search"
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
      Image           =   "frmMain.frx":0AD2
      cBack           =   -2147483633
   End
   Begin VB.OptionButton optNewest 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Newest"
      Height          =   195
      Left            =   3255
      TabIndex        =   8
      Top             =   750
      Width           =   870
   End
   Begin VB.OptionButton optTopCode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Top Code"
      Height          =   195
      Left            =   3255
      TabIndex        =   6
      Top             =   480
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      Left            =   825
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   465
      Width           =   2145
   End
   Begin VB.ComboBox cmbMaxEntry 
      Height          =   315
      Left            =   1935
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   795
      Width           =   1035
   End
   Begin VB.TextBox txtSearch 
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   3255
   End
   Begin VB.Label lblBrowserURL 
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   15
      TabIndex        =   11
      Top             =   1065
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Results Per Page:"
      Height          =   165
      Left            =   45
      TabIndex        =   10
      Top             =   810
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
      Height          =   285
      Left            =   45
      TabIndex        =   7
      Top             =   495
      Width           =   780
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dreamstruct
'Use freely in your code as you please.
'Note clsCoder class is someone elses.


Public XMLGet           As clsXMLWeb        'Class using XML to Grab Web Pages


'Systray Icon
Public WithEvents objfrmSysTray    As frmSysTray
Attribute objfrmSysTray.VB_VarHelpID = -1
Private IconLoaded                  As Boolean

Const SW_SHOWNORMAL = 1


Private Sub cmdIcon_Click()
Me.Hide
If IconLoaded = False Then
    LoadIcon
End If
End Sub

Private Sub cmdMiniMode_Click()
frmMiniMode.Show
Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim File As String
Dim posTop As String
Dim posLeft As String

'Load window as Hidden?
If frmOptions.chkLoadInvisible.Value = 1 Then
    If IconLoaded = False Then
        LoadIcon
    End If
    Me.Visible = False
End If

'Log-in?
If frmOptions.chkRetainLogin.Value = 1 Then
    
End If
Set XMLGet = New clsXMLWeb              'Create Instance of XML HTTP Retrieval Class to be used as Object


'Load with Windows?
If frmOptions.chkLoadWithWindows.Value = 1 Then
    Call modStartWithWindows.StartWithWindows(App.Title, App.Path, True, AllUsers)
Else
    Call modStartWithWindows.StartWithWindows(App.Title, App.Path, False, AllUsers)
End If



'INI File Path
File = App.Path & "\settings.ini"

'Place form where it was before
posTop = ReadIni(File, "FormPosition", "frmMain_Top")
posLeft = ReadIni(File, "FormPosition", "frmMain_Left")

If posTop <> vbNullString And posLeft <> vbNullString Then
    Me.Top = CLng(posTop)
    Me.Left = CLng(posLeft)
End If

'Set the form on top
Win_ControlOnTop Me, False

'Load as Icon?
If frmOptions.chkLoadAsIcon.Value = 1 Then
    If IconLoaded = False Then
        LoadIcon
    End If
End If



'Set focus on search box
frmMain.txtSearch.SetFocus
'Load Listbox with Amount of Results on Page
cmbMaxEntry.AddItem "50"
cmbMaxEntry.AddItem "25"
cmbMaxEntry.AddItem "10"

'Load Listbox with Categorys
cmbCategory.AddItem ".Net"
cmbCategory.AddItem "ASP / VBScript"
cmbCategory.AddItem "C / C++"
cmbCategory.AddItem "Cold Fusion"
cmbCategory.AddItem "Delphi"
cmbCategory.AddItem "Java / Javascript"
cmbCategory.AddItem "Perl"
cmbCategory.AddItem "PHP"
cmbCategory.AddItem "SQL"
cmbCategory.AddItem "Visual Basic"



'Select First Item of Lists If None Selected
If cmbCategory.ListIndex = -1 Then Me.cmbCategory.ListIndex = 9
If cmbMaxEntry.ListIndex = -1 Then Me.cmbMaxEntry.ListIndex = 0

'Load Settings Using INI
cmbMaxEntry.ListIndex = Abs(CLng(ReadIni(File, "ControlSetting", "frmMain_cmbMaxEntry")))
optNewest.Value = CBool(ReadIni(File, "ControlSetting", "frmMain_optNewest"))
optTopCode.Value = CBool(ReadIni(File, "ControlSetting", "frmMain_optTopCode"))
cmbCategory.ListIndex = CLng(ReadIni(File, "ControlSetting", "frmMain_cmbCategory"))
chkDownload.Value = CLng(ReadIni(File, "ControlSetting", "frmMain_chkDownload"))

'Load Settings Using Registry
'cmbMaxEntry.ListIndex = RegLoad(cmbMaxEntry)
'If cmbMaxEntry.ListIndex = -1 Then cmbMaxEntry.ListIndex = 0
'optTopCode.Value = RegLoad(optTopCode)
'optNewest.Value = RegLoad(optNewest)
'cmbCategory.ListIndex = RegLoad(cmbCategory)
'LoadFormPosition Me



End Sub
Private Sub LoadIcon()
If IconLoaded = False Then
    'Systray Icon
    Set objfrmSysTray = New frmSysTray
    With objfrmSysTray
        .AddMenuItem "&Show", "Show", True
        .AddMenuItem "-"
        .AddMenuItem "Log Me In", "LogMeIn", False
        .AddMenuItem "&Options", "Options", False
        .AddMenuItem "&Close", "Close", False

        .ToolTip = "PSC Search"
        .IconHandle = Me.Icon.Handle
    End With
    IconLoaded = True
End If
End Sub
Private Sub UnloadIcon()
    'Systray Icon Unload
    Unload objfrmSysTray
    Set objfrmSysTray = Nothing
End Sub

Private Sub lvButtons_H1_Click()

End Sub

Private Sub objfrmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
Dim strLoginURL As String
Dim strPassword As String
Dim strUsername As String
Dim strLoginInfo As String
Dim strTxtReturnURL As String
Dim strLogin As String
Dim enc As clsCoder
Set enc = New clsCoder

    'Systray Menu Find Menu Item Clicked (sKey)
    Select Case sKey
        Case "Show"
            frmMain.Show
        Case "LogMeIn"
            'Grab their PSC URL from browser
            'Put it into return URL
            'Launch the URL
            strPassword = PSCOpt.PSCPassword
            strUsername = PSCOpt.PSCUserName
            strPassword = modRC4.DecryptString(strPassword, "OurKey", True)
            strUsername = modRC4.DecryptString(strUsername, "OurKey", True)
            If frmOptions.chkRetainLogin.Value = 1 And strUsername <> vbNullString And strPassword <> vbNullString Then
                'If it's the first time searching and they have chosen
                'to retain log-in information then launch login URL
                strLoginURL = "https://www.rentacoder.com/ads/authentication/LoginAction.asp?txtReturnURL="
                strTxtReturnURL = gCurrentBrowserURL
                strTxtReturnURL = GetBrowserURL(frmMain.lblBrowserURL, "planet")
                If strTxtReturnURL = vbNullString Then
                    strTxtReturnURL = "http://www.planet-source-code.com/"
                End If
                strLoginInfo = "&lngWId=1&blnOutsideOfVBSubWeb=True&txtEmailAddress=" & strUsername & "&txtPassword=" & strPassword ' & "&strPassKey=" & ""
                strLogin = strLoginURL & enc.sURLEncode(strTxtReturnURL) & strLoginInfo
                ShellExecute Me.hwnd, "open", strLogin, vbNullString, vbNullString, SW_SHOWNORMAL
            End If
        Case "Options"
            frmOptions.Show
        Case "Close"
            Unload objfrmSysTray
            Set objfrmSysTray = Nothing
            Unload Me
            End
    End Select
 End Sub
 Private Sub objfrmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    Me.Show
    Me.ZOrder
End Sub
Private Sub objfrmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If eButton = vbRightButton Then
        objfrmSysTray.ShowMenu
    End If
End Sub

Public Sub ExecuteSearch(PageNum As String, strSearchText As String)
Dim CatID           As Integer
Dim URL             As String
Dim rootURL         As String
Dim strPage         As String
Dim SearchItem      As String
Dim TopCode         As String
Dim Newest          As String
Dim MaxEntry        As String
Dim strCodeCatID    As String
Dim OurPage         As String
Dim strPassword     As String
Dim strUsername     As String
Dim strResultsURL   As String
Dim strLoginURL     As String
Dim strLoginInfo    As String
Dim strTxtReturnURL As String
Dim strLogin        As String
Dim enc             As clsCoder
Dim arrEntries() As String
Dim X As Integer
'New Encode URL object
Set enc = New clsCoder

'Search Text
SearchItem = enc.sURLEncode(strSearchText)

If frmMain.optTopCode.Value = True Then
  TopCode = "True"
End If
If frmMain.optNewest.Value = True Then
    Newest = "DateDescending"
    TopCode = "False"   'It's not possible to sort by date and top code simultaneously, so set to false.
End If
CatID = Me.cmbCategory.ListIndex
'<option value=10>.Net</option>
'<option value=4>ASP / VbScript</option>
'<option value=3>C / C++</option>
'<option value=9>Cold Fusion</option>
'<option value=7>Delphi</option>
'<option value=2>Java / Javascript</option>
'<option value=6>Perl</option>
'<option value=8>PHP</option>
'<option value=5>SQL</option>
'<option value=1 selected >Visual Basic</option>

'Based on listindex of lstCategory listbox select the PSC Category ID
Select Case CatID
    Case 0
    '.Net
    CatID = 10
    Case 1:
    'ASP / VBScript
    CatID = 4
    Case 2:
    'C / C++
    CatID = 3
    Case 3:
    'Cold Fusion
    CatID = 9
    Case 4:
    'Delphi
    CatID = 7
    Case 5:
    'Java / Javascript
    CatID = 2
    Case 6:
    'Perl
    CatID = 6
    Case 7:
    'PHP
    CatID = 8
    Case 8:
    'SQL
    CatID = 5
    Case 9:
    'Visual Basic
    CatID = 1
End Select
strCodeCatID = CStr(CatID)

MaxEntry = cmbMaxEntry.List(cmbMaxEntry.ListIndex)
If PageNum = vbNullString Then PageNum = "1"

'Search URL on PSC
'http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=4&grpCategories=-1&txtMaxNumberOfEntriesPerPage=50&optSort=DateDescending&chkThoroughSearch=&blnTopCode=True&blnNewestCode=False&blnAuthorSearch=False&lngAuthorId=&strAuthorName=&blnResetAllVariables=&blnEditCode=False&mblnIsSuperAdminAccessOn=False&intFirstRecordOnPage=101&intLastRecordOnPage=150&intMaxNumberOfEntriesPerPage=50&intLastRecordInRecordset=446&chkCodeTypeZip=&chkCodeDifficulty=&chkCodeTypeText=&chkCodeTypeArticle=&chkCode3rdPartyReview=&txtCriteria=email&cmdGoToPage=1&lngMaxNumberOfEntriesPerPage=50

'Search URL on PSC broken down
'http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?
'lngWId=4
'&grpCategories=-1
'&txtMaxNumberOfEntriesPerPage=50
'&optSort=DateDescending
'&chkThoroughSearch=
'&blnTopCode=True
'&blnNewestCode=False
'&blnAuthorSearch=False
'&lngAuthorId=
'&strAuthorName=
'&blnResetAllVariables=
'&blnEditCode=False
'&mblnIsSuperAdminAccessOn=False
'&intFirstRecordOnPage=101
'&intLastRecordOnPage=150
'&intMaxNumberOfEntriesPerPage=50
'&intLastRecordInRecordset=446
'&chkCodeTypeZip=
'&chkCodeDifficulty=
'&chkCodeTypeText=
'&chkCodeTypeArticle=
'&chkCode3rdPartyReview=
'&txtCriteria=email
'&cmdGoToPage=1
'&lngMaxNumberOfEntriesPerPage=50

rootURL = "http://www.planet-source-code.com/"
strPage = "vb/scripts/BrowseCategoryOrSearchResults.asp?"
URL = rootURL & strPage
URL = URL & "lngWId=" & strCodeCatID                         'Language Category
URL = URL & "&grpCategories=-1"                              '?
URL = URL & "&txtMaxNumberOfEntriesPerPage=" & MaxEntry      'Amount of Results Displayed
URL = URL & "&optSort=" & Newest                             'Ascending or Descending
'URL = URL & "&chkThoroughSearch="                           'Thorough Search throuch Code Lines
URL = URL & "&blnTopCode=" & TopCode                         'Sort by Top Code
'URL = URL & "&blnNewestCode="                               'Sort by Newst Code
'URL = URL & "&blnAuthorSearch=True"                         'Conducting Author Search?
'URL = URL & "&lngAuthorID="                                 'ID of Author
'URL = URL & "&strAuthorName=Lavolpe"                        'Author Name Search
URL = URL & "&blnResetAllVariables=False"                    'Reset these variables
'URL = URL & "&blnEditCode=False"                            '?
'URL = URL & "&mblnIsSuperAdminAccessOn="                    '?
URL = URL & "&intFirstRecordOnPage="                         'First Record
URL = URL & "&intLastRecordOnPage="                          'Last Record Displayed
URL = URL & "&intMaxNumberOfEntriesPerPage=" & MaxEntry      'Amount of Results Displayed
URL = URL & "&intLastRecordInRecordSet="                     'Last Record in Set
'URL = URL & "&chkTypeZip="                                  'Zip Files
'URL = URL & "&chkCodeDifficulty="                           'Beginner/Intermediate/expert Code?
'URL = URL & "&chkCodeTypeText="                             '
'URL = URL & "&chkCode3rdPartyReview="                       '3rd Party Review
URL = URL & "&txtCriteria=" & SearchItem                     'Search Keyword
URL = URL & "&cmdGoToPage=" & PageNum                        'Page Number
URL = URL & "&lngMaxNumberOfEntriesPerPage=" & MaxEntry      'Amount of Results Displayed

'Place URL in Public Variable
gCurrentBrowserURL = URL

'Open web page through browser
strPassword = PSCOpt.PSCPassword
strUsername = PSCOpt.PSCUserName
strPassword = modRC4.DecryptString(strPassword, "OurKey", True)
strUsername = modRC4.DecryptString(strUsername, "OurKey", True)

If frmOptions.chkRetainLogin.Value = 1 And strUsername <> vbNullString And strPassword <> vbNullString Then
    'If it's the first time searching and they have chosen
    'to retain log-in information then launch login URL
    strLogin = LoginURL(URL, strUsername, strPassword)
    ShellExecute Me.hwnd, "open", strLogin, vbNullString, vbNullString, SW_SHOWNORMAL
Else
    ShellExecute Me.hwnd, "open", URL, vbNullString, vbNullString, SW_SHOWNORMAL
End If

'If they check to download code then..
If Me.chkDownload.Value = 1 Then
    OurPage = XMLGet.GetDataAsString(URL)
    PSCResults.ResultNum = 0    'Initialize Page Parsing Position
    'This loop handles each entry on a single results page
    'Determine how many results are on the page
    arrEntries = Split(OurPage, "<!--descrip-->")   'Split page at descrip tag
    For X = 0 To UBound(arrEntries)
        Me.Caption = "PSC Quick Search - Parsing " & X & " of " & UBound(arrEntries)
        Call ParsePSC(OurPage)
        'ParsePSC stored each parsed field into a variable
        'e.g. Username, Excellents, Code Description
        'Now we take this info and add to the listview
        'and loop through all of them (on first page)
    Next
    Me.Caption = "PSC Quick Search"
    frmResults.Show
End If
End Sub


Private Sub cmdFavorites_Click()
frmFavorites.Show
End Sub

Private Sub cmdNext_Click()
PageNum = PageNum + 1
ExecuteSearch CStr(PageNum), frmMain.txtSearch.Text
End Sub

Private Sub cmdOptions_Click()
frmOptions.Show
End Sub

Private Sub cmdPrev_Click()
If PageNum >= 2 Then            'Make sure we do not go lower than first page
    PageNum = PageNum - 1
End If
ExecuteSearch CStr(PageNum), frmMain.txtSearch.Text
End Sub

Private Sub cmdSearch_Click()
PageNum = 1
ExecuteSearch CStr(PageNum), frmMain.txtSearch.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim File    As String

File = App.Path & "\settings.ini"

'Use INI File to Save Settings
WriteIni File, "ControlSetting", "frmMain_cmbMaxEntry", CStr(cmbMaxEntry.ListIndex)
WriteIni File, "ControlSetting", "frmMain_optNewest", CStr(optNewest.Value)
WriteIni File, "ControlSetting", "frmMain_optTopCode", CStr(optTopCode.Value)
WriteIni File, "ControlSetting", "frmMain_cmbCategory", CStr(cmbCategory.ListIndex)
WriteIni File, "ControlSetting", "frmMain_chkDownload", CStr(chkDownload.Value)
WriteIni File, "FormPosition", "frmMain_Top", Me.Top
WriteIni File, "FormPosition", "frmMain_Left", Me.Left

'Use Registry to Save Settings
'RegSave cmbMaxEntry, cmbMaxEntry.ListIndex
'RegSave frmMain.optTopCode, optTopCode.Value
'RegSave frmMain.optNewest, optNewest.Value
'RegSave frmMain.cmbCategory, cmbCategory.ListIndex
'SaveFormPosition Me
If frmOptions.chkMainWinRemoveIcon.Value = 1 Then
    If IconLoaded = True Then
        UnloadIcon
    End If
    End
Else
    If IconLoaded = False Then
        If frmOptions.chkMainWinRemoveIcon.Value = 0 Then
            LoadIcon
        Else
            UnloadAllForms
            End
        End If
    End If
End If
End Sub
