VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4155
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "General"
      Height          =   1650
      Left            =   45
      TabIndex        =   7
      Top             =   60
      Width           =   3870
      Begin VB.CheckBox chkLoadInvisible 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Load Main Window as Hidden"
         Height          =   210
         Left            =   165
         TabIndex        =   11
         Top             =   870
         Width           =   2475
      End
      Begin VB.CheckBox chkMainWinRemoveIcon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Closing Main Window Exits and Removes Icon"
         Height          =   270
         Left            =   165
         TabIndex        =   10
         Top             =   1110
         Width           =   3630
      End
      Begin VB.CheckBox chkLoadAsIcon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Load Icon at Start"
         Height          =   270
         Left            =   165
         TabIndex        =   9
         Top             =   600
         Width           =   1740
      End
      Begin VB.CheckBox chkLoadWithWindows 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Load with Windows"
         Height          =   285
         Left            =   165
         TabIndex        =   8
         Top             =   345
         Width           =   1740
      End
   End
   Begin VB.CheckBox chkRetainLogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Retain Log-in Info"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   3360
      Width           =   1740
   End
   Begin PlanetSourceCodeQuickSearch.lvButtons_H cmdOK 
      Default         =   -1  'True
      Height          =   360
      Left            =   3015
      TabIndex        =   5
      Top             =   3735
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   635
      Caption         =   "&OK"
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
   Begin VB.Frame fraRetainLogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Retain Log-in Info"
      Enabled         =   0   'False
      Height          =   1365
      Left            =   45
      TabIndex        =   0
      Top             =   1920
      Width           =   3870
      Begin VB.TextBox txtPassword 
         Enabled         =   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1365
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   795
         Width           =   1800
      End
      Begin VB.TextBox txtUsername 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1365
         TabIndex        =   2
         Top             =   390
         Width           =   1800
      End
      Begin VB.Label lblDispPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Enabled         =   0   'False
         Height          =   210
         Left            =   450
         TabIndex        =   3
         Top             =   870
         Width           =   870
      End
      Begin VB.Label lblDispUsername 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address:"
         Enabled         =   0   'False
         Height          =   225
         Left            =   165
         TabIndex        =   1
         Top             =   390
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strUsername         As String   'PSC Username
Private strPassword         As String   'PSC Password

Private Sub Check1_Click()

End Sub

Private Sub chkLoadAsIcon_Click()
If chkLoadAsIcon.Value = 0 Then
    Me.chkLoadInvisible.Value = 0
    Me.chkLoadInvisible.Enabled = False
Else
    Me.chkLoadInvisible.Enabled = True
End If
End Sub

Private Sub chkLoadInvisible_Click()
If chkLoadInvisible.Value = True Then
    'If they load main form invisible then we need an icon
    chkLoadAsIcon = True
End If
End Sub

Private Sub chkRetainLogin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkRetainLogin.Value = 1 Then
        'OK Fine
        txtUsername.Enabled = True
        txtPassword.Enabled = True
        lblDispPassword.Enabled = True
        lblDispUsername.Enabled = True
        fraRetainLogin.Enabled = True
Else
        txtUsername.Enabled = False
        txtPassword.Enabled = False
        lblDispPassword.Enabled = False
        lblDispUsername.Enabled = False
        fraRetainLogin.Enabled = False
        chkRetainLogin.Value = 0
End If
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim strFile As String

strFile = App.Path & "\Settings.ini"
strUsername = ReadIni(strFile, "UserSetting", "Username")
strPassword = ReadIni(strFile, "UserSetting", "Password")
If strUsername <> vbNullString And strPassword <> vbNullString Then
    'Retain them in memory of our program (see modPSCQuickSearch)
    modPSCQuickSearch.PSCOpt.PSCPassword = strPassword
    modPSCQuickSearch.PSCOpt.PSCUserName = strUsername
    
    txtUsername = modRC4.DecryptString(strUsername, "YourKey", True)
    txtPassword = modRC4.DecryptString(strPassword, "YourKey", True)
End If
chkRetainLogin.Value = CInt(Val(ReadIni(strFile, "ControlSetting", "frmOptions_chkRetainLogin")))
chkLoadAsIcon.Value = CInt(Val(ReadIni(strFile, "ControlSetting", "frmOptions_chkLoadAsIcon")))
chkLoadWithWindows.Value = CInt(Val(ReadIni(strFile, "ControlSetting", "frmOptions_chkLoadWithWindows")))
chkMainWinRemoveIcon.Value = CInt(Val(ReadIni(strFile, "ControlSetting", "frmOptions_chkMainWinRemoveIcon")))
chkLoadInvisible.Value = CInt(Val(ReadIni(strFile, "ControlSetting", "frmOptions_chkLoadInvisible")))

If chkRetainLogin.Value = 1 Then
'OK Fine
        txtUsername.Enabled = True
        txtPassword.Enabled = True
        lblDispPassword.Enabled = True
        lblDispUsername.Enabled = True
        fraRetainLogin.Enabled = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strFile As String
strFile = App.Path & "\Settings.ini"
'Store the Username & Password in the INI file
'First we will encrypt them..
If txtUsername <> vbNullString And txtPassword <> vbNullString Then
    strUsername = modRC4.EncryptString(txtUsername, "YourKey", True)
    strPassword = modRC4.EncryptString(txtPassword, "YourKey", True)
End If

'Retain them in memory of our program (see modPSCQuickSearch)
modPSCQuickSearch.PSCOpt.PSCPassword = strPassword
modPSCQuickSearch.PSCOpt.PSCUserName = strUsername

'Now store them in INI File..
WriteIni strFile, "UserSetting", "Username", strUsername
WriteIni strFile, "UserSetting", "Password", strPassword
WriteIni strFile, "ControlSetting", "frmOptions_chkRetainLogin", frmOptions.chkRetainLogin.Value
WriteIni strFile, "ControlSetting", "frmOptions_chkLoadAsIcon", frmOptions.chkLoadAsIcon.Value
WriteIni strFile, "ControlSetting", "frmOptions_chkLoadWithWindows", frmOptions.chkLoadWithWindows.Value
WriteIni strFile, "ControlSetting", "frmOptions_chkMainWinRemoveIcon", frmOptions.chkMainWinRemoveIcon.Value
WriteIni strFile, "ControlSetting", "frmOptions_chkLoadInvisible", frmOptions.chkLoadInvisible.Value

'Optionally you could store these settings in registry...

End Sub
