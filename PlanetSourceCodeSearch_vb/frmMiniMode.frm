VERSION 5.00
Begin VB.Form frmMiniMode 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   555
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   45
      Picture         =   "frmMiniMode.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   210
      TabIndex        =   5
      Top             =   45
      Width           =   210
   End
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdMiniNext 
      Height          =   240
      Left            =   3285
      TabIndex        =   4
      ToolTipText     =   "Next"
      Top             =   135
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Caption         =   "lvButtons_H5"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   7
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
      Image           =   "frmMiniMode.frx":010B
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdMiniPrev 
      Height          =   240
      Left            =   3000
      TabIndex        =   3
      ToolTipText     =   "Previous"
      Top             =   135
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Caption         =   "lvButtons_H4"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   7
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
      Image           =   "frmMiniMode.frx":018B
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdMiniSearch 
      Default         =   -1  'True
      Height          =   360
      Left            =   2565
      TabIndex        =   2
      Top             =   90
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   635
      Caption         =   "lvButtons_H3"
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
      Image           =   "frmMiniMode.frx":020A
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtMiniSearch 
      Height          =   360
      Left            =   375
      TabIndex        =   1
      Top             =   90
      Width           =   2115
   End
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdClose 
      Height          =   195
      Left            =   3660
      TabIndex        =   0
      Top             =   30
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   344
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   7
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
      Image           =   "frmMiniMode.frx":03FA
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.Shape Shape1 
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   3885
   End
End
Attribute VB_Name = "frmMiniMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
Unload Me
frmMain.Show
End Sub

Private Sub cmdMiniNext_Click()
PageNum = PageNum + 1
frmMain.ExecuteSearch CStr(PageNum), frmMiniMode.txtMiniSearch
End Sub

Private Sub cmdMiniPrev_Click()
If PageNum >= 2 Then            'Make sure we do not go lower than first page
    PageNum = PageNum - 1
End If
frmMain.ExecuteSearch CStr(PageNum), frmMiniMode.txtMiniSearch
End Sub

Private Sub cmdMiniSearch_Click()
PageNum = 1
frmMain.ExecuteSearch CStr(PageNum), frmMiniMode.txtMiniSearch
End Sub

Private Sub Form_Load()
Dim iniSettingsFile As String

'Set the form on top
Win_ControlOnTop Me, False

'INI File Path
iniSettingsFile = App.Path & "\settings.ini"

'Place form where it was before
posTop = ReadIni(iniSettingsFile, "FormPosition", "frmMiniMode_Top")
posLeft = ReadIni(iniSettingsFile, "FormPosition", "frmMiniMode_Left")
If posTop <> vbNullString And posLeft <> vbNullString Then
    Me.Left = posLeft
    Me.Top = posTop
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  Call DragForm(Me)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim iniSettingsFile As String
iniSettingsFile = App.Path & "\settings.ini"
WriteIni iniSettingsFile, "FormPosition", "frmMiniMode_Top", Me.Top
WriteIni iniSettingsFile, "FormPosition", "frmMiniMode_Left", Me.Left
End Sub
