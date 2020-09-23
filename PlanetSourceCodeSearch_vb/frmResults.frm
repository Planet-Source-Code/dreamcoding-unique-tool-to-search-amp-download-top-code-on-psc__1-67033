VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmResults 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Results"
   ClientHeight    =   4335
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14205
   Icon            =   "frmResults.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmResults.frx":038A
   ScaleHeight     =   4335
   ScaleWidth      =   14205
   StartUpPosition =   3  'Windows Default
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdDownloadSelected 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   3900
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   556
      Caption         =   "Download Selected"
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
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdViewDownloads 
      Height          =   315
      Left            =   2115
      TabIndex        =   8
      Top             =   3900
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Caption         =   "View Download Folder"
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
   Begin VB.ListBox lstDownloadNames 
      Height          =   1035
      Left            =   2505
      TabIndex        =   6
      Top             =   4830
      Width           =   1440
   End
   Begin VB.TextBox txtStatus 
      Height          =   750
      Left            =   8160
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   4935
      Width           =   1890
   End
   Begin VB.ListBox lstDownloadURLS 
      Height          =   1035
      Left            =   285
      TabIndex        =   4
      Top             =   4620
      Width           =   2100
   End
   Begin VB.ListBox lstCodeURL 
      Height          =   840
      Left            =   6555
      TabIndex        =   3
      Top             =   4905
      Width           =   1425
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   10425
      Top             =   5205
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   34
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResults.frx":04D4
            Key             =   "CodeZip"
            Object.Tag             =   "CodeZip"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   465
      TabIndex        =   2
      Top             =   5685
      Width           =   465
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   4080
      TabIndex        =   1
      Top             =   4875
      Width           =   2250
   End
   Begin MSComctlLib.ListView lvwResults 
      Height          =   3705
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   6535
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgSmall"
      SmallIcons      =   "imgSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   7290
      TabIndex        =   7
      Top             =   4020
      Width           =   6765
   End
   Begin VB.Menu mnuRoot 
      Caption         =   "Root"
      Visible         =   0   'False
      Begin VB.Menu mnuViewCode 
         Caption         =   "View Code Page"
      End
      Begin VB.Menu mnuDownloadFolder 
         Caption         =   "View Download Folder"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "Download Selected"
      End
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public objXmlWeb As clsXMLWeb
Private Type PSCResults
    CodeAuthorName          As String
    CodeSubmittedOn         As String
    CodeLevel               As String
    CodeUserRating          As String
    CodeAccessed            As String
    CodeCompatibility       As String
    CodeAuthorID            As String
    CodeAuthorPhoto         As String
    CodeAuthorPlacement     As String
    CodeAuthorWord          As String
    CodeDescription         As String
End Type
Private objResize                 As cResize

Private Sub cmdDownloadSelected_Click()
Call DownloadSelected
End Sub
Private Sub DownloadSelected()

Dim X As Long
Dim CodeURL As String
Dim OurPage As String
PSCResults.DownloadCodeNames = vbNullString
Me.lstDownloadURLS.Clear
'Go through listview see what's selected
'grab URLs based on that
'Navigate to each page and grab the download link
'(have to do this because zip file has unique code in each one)
For X = 0 To lvwResults.ListItems.Count - 1
    If lvwResults.ListItems.Item(X + 1).Selected = True Then 'Found a selected item
        
        'Grab Name of Code
        PSCResults.DownloadCodeNames = PSCResults.DownloadCodeNames & lvwResults.ListItems.Item(X + 1).Text & UniqueToken
        frmResults.lblStatus.Caption = "Grabbing Item - " & lvwResults.ListItems.Item(X + 1).Text
        'Process the URL to find Download link
        CodeURL = Me.lstCodeURL.List(X)
        OurPage = objXmlWeb.GetDataAsString(CodeURL)
        ShowZip = QuickParse(1, OurPage, "ShowZip.asp?", Chr$(34) & ">", False, False)
        If ShowZip <> vbNullString Then
            Me.lstDownloadURLS.AddItem "http://www.planetsourcecode.com/vb/scripts/" & ShowZip  'Add it to Download list
        Else
            'Code is likely not a zip file, maybe an article or C&P post
            'Let them know
            Me.txtStatus = txtStatus & lvwResults.ListItems.Item(X).Text & " - Not a zip file and cannot be downloaded."
        End If
    End If
DoEvents
Next X
DownloadCode
End Sub
Private Sub cmdViewDownloads_Click()
OpenDownloadFolder
End Sub

Private Sub Form_Load()
Dim strFile         As String
Dim strFileName     As String
Dim posLeft         As String
Dim posTop          As String
Dim posWidth        As String
Dim posHeight       As String

Set objXmlWeb = New clsXMLWeb
Set objResize = New cResize

    strFile = App.Path & "\Settings.ini"
    objResize.InitPositions Me, False, False
    SetupListView
    posWidth = ReadIni(strFile, "FormSize", "frmResults_Width")
    posHeight = ReadIni(strFile, "FormSize", "frmResults_Height")
    posTop = ReadIni(strFile, "FormPosition", "frmResults_Top")
    posLeft = ReadIni(strFile, "FormPosition", "frmResults_Left")
    If posWidth <> vbNullString And posHeight <> vbNullString Then
        Me.Width = Val(posWidth)
        Me.Height = Val(posHeight)
    End If
    If posTop <> vbNullString And posLeft <> vbNullString Then
        Me.Left = Val(posLeft)
        Me.Top = Val(posTop)
    End If
'Set the form on top
Win_ControlOnTop Me, False

End Sub
Private Sub SetupListView()
  ' Create the column headers.
   ' Me.lvwResults.SmallIcons = Me.imgSmall
   ' With lvwResults
   ' .SmallIcons = imgSmall
    'End With
    '.View = imgSmall
    'Set .Icons = imgSmall
    'End With
   '  Me.lvwResults.SmallIcons = Me.imgSmall
    ' lvwResults.ListItems.Item = 1
         
    'list_item.SmallIcon = 1
            
  
   ' Set column_header = Me.lvwResults.ListItems.Add(, , "", , 1)
   ' Set column_header = Me.lvwResults.ColumnHeaders.Add(, , , TextWidth(" "), , 1)
    Set column_header = Me.lvwResults.ColumnHeaders.Add(, , "Title", TextWidth("                                                                                                  "))
    Set column_header = Me.lvwResults.ColumnHeaders.Add(, , "Description", TextWidth("                                       "))
    Set column_header = Me.lvwResults.ColumnHeaders.Add(, , "Author", TextWidth("                                        "))
    Set column_header = Me.lvwResults.ColumnHeaders.Add(, , "Excellents", TextWidth("                               "))
    Set column_header = Me.lvwResults.ColumnHeaders.Add(, , "Users Rating", TextWidth("                               "))
    Set column_header = Me.lvwResults.ColumnHeaders.Add(, , "Level", TextWidth("                                       "))
    Set column_header = Me.lvwResults.ColumnHeaders.Add(, , "Submitted On", TextWidth("                                "))
    Set column_header = Me.lvwResults.ColumnHeaders.Add(, , "Code Views", TextWidth("                                    "))
    Set column_header = Me.lvwResults.ColumnHeaders.Add(, , "", TextWidth(""))
    ' Initialize the ImageLists

    'CodeAuthorName          As String
    'CodeSubmittedOn         As String
    'CodeLevel               As String
    'CodeUserRating          As String
    'CodeAccessed            As String
    'CodeCompatibility       As String
    'CodeAuthorID            As String
    'CodeAuthorPhoto         As String
    'CodeAuthorPlacement     As String
    'CodeAuthorWord          As String
    'CodeDescription         As String
 End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim File As String
File = App.Path & "\Settings.ini"
WriteIni File, "FormPosition", "frmResults_Top", Me.Top
WriteIni File, "FormPosition", "frmResults_Left", Me.Left
WriteIni File, "FormSize", "frmResults_Width", Me.Width
WriteIni File, "FormSize", "frmResults_Height", Me.Height
End Sub





Private Sub lstDownloadURLS_Click()
'ShellExecute Me.hWnd, "open", strURLtoDownload, vbNullString, vbNullString, SW_SHOWNORMAL

End Sub
Private Sub DownloadCode()
Dim strURLtoDownload        As String
Dim strFileName             As String
Dim arrDownloadCodeNames()  As String
Dim lngChosenDownload       As Long
Dim strDownloadCodeName     As String

'lngChosenDownload = lstDownloadURLS.ListIndex
'strURLtoDownload = lstDownloadURLS.List(lngChosenDownload)

 For X = 1 To lstDownloadURLS.ListCount
    strURLtoDownload = lstDownloadURLS.List(X - 1)
    strdownloadcodenames = Split(PSCResults.DownloadCodeNames, UniqueToken)(X - 1)
    strshortdownloadname = Left(Replace(strdownloadcodenames, " ", "_"), 10)
    strshortdownloadname = Replace(strshortdownloadname, "/", "_")
    strCodeID = Replace(QuickParse(1, strURLtoDownload, "lngCodeId=", "&", False, False), "lngCodeId=", vbNullString)
    If DIRExists(App.Path & "\Downloads\", True) = False Then
        MkDir (App.Path & "\Downloads\")
        strFileName = App.Path & "\Downloads\" & strshortdownloadname & "_" & strCodeID & ".zip"
    Else
        strFileName = App.Path & "\Downloads\" & strshortdownloadname & "_" & strCodeID & ".zip"
    End If
    Me.lblStatus.Caption = "Downloading... - " & X & " of " & lstDownloadURLS.ListCount
    Call objXmlWeb.GetFile(strURLtoDownload, strFileName)
    Me.lblStatus.Caption = "Download #" & X & " Complete!"
Next

End Sub



Private Sub lvwResults_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
LVSortColumns Me.lvwResults, ColumnHeader
End Sub

Private Sub lvwResults_DblClick()
Dim SelItem As Long
Dim strURL As String

'Double click will launch code page
SelItem = lvwResults.SelectedItem.Index - 1
strURL = frmResults.lstCodeURL.List(SelItem)
ShellExecute Me.hwnd, "open", strURL, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub lvwResults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If the right mouse button clicked
'we will display the pop-up menu

If Button = 2 Then
    'If they selected just one code then
    'we will allow them to view that code page
    'so show this option in the pop-up menu
    If modListview.GetSelectedCount(Me.lvwResults) = 1 Then
        Me.mnuViewCode.Visible = True
    Else
        Me.mnuViewCode.Visible = False
    End If
Me.PopupMenu mnuRoot, 0
End If
End Sub

Private Sub mnuDownload_Click()
Call DownloadSelected
End Sub

Private Sub mnuDownloadFolder_Click()
OpenDownloadFolder
End Sub
Private Sub OpenDownloadFolder()
Dim X As Long
Dim OurPath As String
OurPath = App.Path & "\Downloads"
X = ShellExecute(Me.hwnd, "open", OurPath, vbNullString, vbNullString, SW_SHOWNORMAL)
If X = 2 Then
    MkDir OurPath
    X = ShellExecute(Me.hwnd, "open", OurPath, vbNullString, vbNullString, SW_SHOWNORMAL)
End If
End Sub
Private Sub mnuViewCode_Click()
Dim SelItem As Long
Dim strURL As String

'Double click will launch code page
SelItem = lvwResults.SelectedItem.Index - 1
strURL = frmResults.lstCodeURL.List(SelItem)
ShellExecute Me.hwnd, "open", strURL, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub
