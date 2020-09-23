VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'   Programmer: Brent Culpepper (IDontKnow)
'   Project:    cResize.cls
'   Credits:    Class was adapted from a usercontrol written by
'               Francesco Balena, posted at http://www.vb2themax.com
'   Date:       March 17, 2004
'   Description:    A class module that traps the resizing event
'                   of a form and resizes the contained controls
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Option Explicit
Private Type tControlPosition
    Left                         As Single
    Top                          As Single
    Width                        As Single
    Height                       As Single
    FontSize                     As Single
End Type
Private WithEvents m_Form    As Form
Attribute m_Form.VB_VarHelpID = -1
Private m_Position()         As tControlPosition
Private m_Ratio              As Single
Private m_Height             As Single
Private m_Width              As Single
Private m_ResizeFonts        As Boolean
Private m_KeepRatio          As Boolean
Private m_Initialized        As Boolean
Private oResize                 As cResize
Public Sub InitPositions(OwnerForm As Form, _
                         ByVal bKeepRatio As Boolean, _
                         ByVal bResizeFonts As Boolean)
Dim lCount     As Long
Dim ctlCurrent As Control
    If Not m_Initialized Then
'<CF> :WARNING: 'Unnecessary Exit Procedure fix': 'If structure' logic reversed and extended. Old code 'Exit Sub
'<CF> :WARNING: 'Unnecessary Exit Procedure fix'  removed 'End If '<CF> Structure Expanded.
        Set m_Form = OwnerForm
        m_Height = m_Form.ScaleHeight
        m_Width = m_Form.ScaleWidth
        m_Ratio = m_Form.Height / m_Form.Width
        m_KeepRatio = bKeepRatio
        m_ResizeFonts = bResizeFonts
        ReDim m_Position(m_Form.Controls.Count - 1)
        For Each ctlCurrent In m_Form.Controls
            On Error Resume Next
            With m_Position(lCount)
                If TypeOf ctlCurrent Is Line Then
                    .Left = ctlCurrent.X1
                    .Top = ctlCurrent.Y1
                    .Width = ctlCurrent.X2 - ctlCurrent.X1
                    .Height = ctlCurrent.Y2 - ctlCurrent.Y1
                Else
                    .Left = ctlCurrent.Left
                    .Top = ctlCurrent.Top
                    .Width = ctlCurrent.Width
                    .Height = ctlCurrent.Height
                    .FontSize = ctlCurrent.Font.Size
                End If
            End With
            lCount = lCount + 1
        Next ctlCurrent
        m_Initialized = True
    End If '<CF> :WARNING: 'Unnecessary Exit Procedure fix' added this.
End Sub
Public Property Get KeepRatio() As Boolean
    KeepRatio = m_KeepRatio
End Property
Public Property Let KeepRatio(ByVal vNewValue As Boolean)
    m_KeepRatio = vNewValue
End Property
Private Sub m_Form_Resize()
    ResizeControls
End Sub
Private Sub ResizeControls()
Dim ctlCurrent As Control
Dim lCount     As Long
Dim x_Factor   As Single
Dim y_Factor   As Single
Dim MinFactor  As Single
Static Busy    As Boolean
On Error Resume Next
    If Not m_Form.WindowState = vbMinimized Then
'<CF> :WARNING: 'Unnecessary Exit Procedure fix': 'If structure' logic reversed and extended. Old code 'Exit Sub
'<CF> :WARNING: 'Unnecessary Exit Procedure fix'  removed 'End If '<CF> Structure Expanded.
        If Not m_Initialized Then
            Exit Sub
'<CF> :SUGGESTION: Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 11)
'<CF> (EXPERIMENTAL follow advice with care )
'<CF> No recommended action but consider coding around it.
        End If '<CF> Structure Expanded.
        If m_KeepRatio Then
' Prevent recursive loop
            Busy = True
            m_Form.Height = m_Ratio * m_Form.Width
            Busy = False
        End If
' Get the current scale factors:
        x_Factor = m_Form.ScaleWidth / m_Width
        y_Factor = m_Form.ScaleHeight / m_Height
' Use the lesser factor as minimum for font sizing:
        If x_Factor < y_Factor Then
            MinFactor = x_Factor
        Else
            MinFactor = y_Factor
        End If
        For Each ctlCurrent In m_Form.Controls
            On Error Resume Next
            With m_Position(lCount)
                If TypeOf ctlCurrent Is Line Then
                    ctlCurrent.X1 = x_Factor * .Left
                    ctlCurrent.Y1 = y_Factor * .Top
                    ctlCurrent.X2 = ctlCurrent.X1 + x_Factor * .Width
                    ctlCurrent.Y2 = ctlCurrent.Y1 + y_Factor * .Height
                Else
                    If m_ResizeFonts Then
                        ctlCurrent.Font.Size = MinFactor * .FontSize
                    End If
                    ctlCurrent.Left = x_Factor * .Left
                    ctlCurrent.Top = y_Factor * .Top
                    ctlCurrent.Width = x_Factor * .Width
                    If Not (TypeOf ctlCurrent Is ComboBox) Then
                        ctlCurrent.Height = y_Factor * .Height
                    End If
                End If
            End With
            lCount = lCount + 1
        Next ctlCurrent
    End If '<CF> :WARNING: 'Unnecessary Exit Procedure fix' added this.
End Sub
Public Property Get ResizeFonts() As Boolean
    ResizeFonts = m_ResizeFonts
End Property
Public Property Let ResizeFonts(ByVal vNewValue As Boolean)
    m_ResizeFonts = vNewValue
End Property
':)Code Fixer V4.0.0 (Tuesday, 06 June 2006 03:41:01) 25 + 130 = 155 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|033322222222222222222222222222|1112222|2221222|222222222233|1111111111111|1122222222220|333333|


