Attribute VB_Name = "modRC4"
Option Explicit
' Visual Basic RC4 Implementation
'
' Standard RC4 implementation with file support, hex conversion,
' speed string concatenation and overall optimisations for Visual Basic.
' RC4 is an extremely fast and very secure stream cipher from RSA Data
' Security, Inc. I recommend this for high risk low resource environments.
' It's speed is very very attractive. Patents do apply for commercial use.
'
' Information on the algorithm can be found at:
' http://www.rsasecurity.com/rsalabs/faq/3-6-3.html
''Private m_Key               As String
'<CF> :WARNING: Unused Variable 'm_Key'
'<CF> May be a prototype Variable you have not yet implimented or left over from a deleted Control.
Private m_sBox(0 To 255)    As Long
'<CF> :WARNING: Integer Private upgraded to Long.
Private byteArray()         As Byte
Private hiByte              As Long
Private hiBound             As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                  Source As Any, _
                                                                  ByVal Length As Long)
Private Sub Append(ByRef StringData As String, _
                   Optional ByVal Length As Long)
'<CF> :WARNING: 'ByVal ' inserted for Parameter 'Optional Length As Long'
Dim DataLength As Long
    If Length > 0 Then
        DataLength = Length
    Else
        DataLength = Len(StringData)
    End If '<CF> Structure Expanded.
    If DataLength + hiByte > hiBound Then
        hiBound = hiBound + 1024
        ReDim Preserve byteArray(hiBound)
    End If
    CopyMem ByVal VarPtr(byteArray(hiByte)), ByVal StringData, DataLength
    hiByte = hiByte + DataLength
End Sub
Public Sub DecryptByte(byteArray() As Byte, _
                       Optional Key As String)
    EncryptByte byteArray(), Key
End Sub
Public Function DecryptString(Text As String, _
                              Optional Key As String, _
                              Optional ByVal IsTextInHex As Boolean) As String
Dim byteArray() As Byte
If Text <> vbNullString Then
    If IsTextInHex Then
        Text = DeHex(Text)
    End If
    byteArray() = StrConv(Text, vbFromUnicode)
    DecryptByte byteArray(), Key
    DecryptString = StrConv(byteArray(), vbUnicode)
End If
End Function
Private Function DeHex(Data As String) As String
Dim iCount As Double
    Reset
    For iCount = 1 To Len(Data) Step 2
        Append Chr$(Val("&H" & Mid$(Data, iCount, 2)))
    Next iCount
    DeHex = GData
    Reset
End Function
Public Sub EncryptByte(byteArray() As Byte, _
                       Optional ByVal Key As String)
'<CF> :WARNING: 'ByVal ' inserted for Parameter 'Optional Key As String'
'<CF> :WARNING: Unused Parameter 'Key' could be removed.
Dim i              As Long
Dim j              As Long
Dim Temp           As Byte
Dim Offset         As Long
Dim OrigLen        As Long
Dim CipherLen      As Long
''Dim CurrPercent    As Long
'<CF> :WARNING: (2) One-Time Dim assignment line collapsed to in-line code.
Dim NextPercent    As Long
Dim sBox(0 To 255) As Long
'<CF> :UPDATED: Multiple Dim line separated
'<CF> :PREVIOUS CODE : Dim i As Long, j As Long, Temp As Byte, Offset As Long, OrigLen As Long, CipherLen As Long, CurrPercent As Long, NextPercent As Long, sBox(0 To 255) As Long
'<CF> :WARNING: Integer Dim upgraded to Long.
'If (Len(Key) > 0) Then Me.Key = Key
'<CF> :WARNING: Unneeded 'Call' command removed.
    CopyMem sBox(0), m_sBox(0), 512
    OrigLen = UBound(byteArray) + 1
    CipherLen = OrigLen
    For Offset = 0 To (OrigLen - 1)
        i = (i + 1) Mod 256
        j = (j + sBox(i)) Mod 256
        Temp = sBox(i)
        sBox(i) = sBox(j)
        sBox(j) = Temp
        byteArray(Offset) = byteArray(Offset) Xor (sBox((sBox(i) + sBox(j)) Mod 256))
        If Offset >= NextPercent Then
'<CF> :WARNING: Unneeded brackets removed
'<CF> :PREVIOUS CODE : If (Offset >= NextPercent) Then
''CurrPercent = Int((Offset / CipherLen) * 100)
            NextPercent = (CipherLen * (((Int((Offset / CipherLen) * 100)) + 1) / 100)) + 1
'<CF> :WARNING: Single use Variable 'CurrPercent' replaced by in-line code '(Int((Offset / CipherLen) * 100))'
        End If
    Next Offset
End Sub
Public Function EncryptString(ByVal Text As String, _
                              Optional Key As String, _
                              Optional ByVal OutputInHex As Boolean) As String
'<CF> :WARNING: 'ByVal ' inserted for Parameters 'Text As String, Optional OutputInHex As Boolean'
Dim byteArray() As Byte
    byteArray() = StrConv(Text, vbFromUnicode)
'<CF> :WARNING: Unneeded 'Call' command removed.
    EncryptByte byteArray(), Key
    EncryptString = StrConv(byteArray(), vbUnicode)
    If OutputInHex Then
'<CF> Pleonasm Removed
'<CF> :PREVIOUS CODE : If OutputInHex = True Then
        EncryptString = EnHex(EncryptString)
    End If '<CF> Structure Expanded.
End Function
Private Function EnHex(ByVal Data As String) As String
'<CF> :WARNING: 'ByVal ' inserted for Parameter 'Data As String'
Dim iCount As Double
Dim sTemp  As String
'<CF> :UPDATED: Multiple Dim line separated
'<CF> :PREVIOUS CODE : Dim iCount As Double, sTemp As String
    Reset
    For iCount = 1 To Len(Data)
        sTemp = Hex$(Asc(Mid$(Data, iCount, 1)))
        If Len(sTemp) < 2 Then
            sTemp = "0" & sTemp
        End If '<CF> Structure Expanded.
        Append sTemp
    Next iCount
    EnHex = GData
    Reset
End Function
Private Property Get GData() As String
Dim StringData As String
    StringData = Space$(hiByte)
'<CF> :PREVIOUS CODE : StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property
Private Sub Reset()
    hiByte = 0
    hiBound = 1024
    ReDim byteArray(hiBound)
End Sub
':)Code Fixer V4.0.0 (Wednesday, 10 May 2006 00:54:07) 20 + 163 = 183 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|033322222222222222222222222222|1112222|2221222|222222222233|1111111111111|1122222222220|333333|


