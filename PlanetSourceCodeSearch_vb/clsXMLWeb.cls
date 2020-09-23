VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Here's a more advance method of posting data with the XMLHTTP class
' I have sacrified code effiency for simplicity sake, so that it would be easier for
' any of you to understand.
' There are 2 methods here.....
' 1 - GetDataAsString returns the source page of the requested URL
' 2 - SubmitDataAsString posts data and returns the source page of the requested URL

Private objXMLHTTP As MSXML2.XMLHTTP
Private mURL As String

Enum enmSubmitMethod
    enmPOST = 1
    enmGET = 2
End Enum


Private Sub Class_Initialize()
    
    ' Create an object of the XMLHTTPRequest class
    Set objXMLHTTP = New MSXML2.XMLHTTP
    
End Sub

Private Sub Class_Terminate()
    
    ' Destroy the object to regain memory
    Set objXMLHTTP = Nothing
    
End Sub

Public Property Get URL() As String
    URL = mURL
End Property

'************************************************************************************
'* 2 functions below will return data as string
'* 1 - GetDataAsString requests a URL and returns the result as a string
'* 2 - SubmitDataAsString submits data to a URL and returns the result as a string
'************************************************************************************

Public Function GetDataAsString(URLString As String) As String
    
    ' Set the URL property
    mURL = Trim(URLString)
    
    ' Let it know that it is a GET operation and set the ASYNC parameter to false
    objXMLHTTP.open "GET", mURL, False
    
    ' Send the call
    objXMLHTTP.send
    
    GetDataAsString = objXMLHTTP.responseText
    
End Function

Public Function SubmitDataAsString(URLString As String, SubmitMethod As enmSubmitMethod, FormInputArray As Variant, FormValueArray As Variant) As String
    
    ' URLString is the URL that is to be requested
    ' FormInputArray are the names of the <INPUT> tags
    ' FormValueArray are the actual values of the corresponding <INPUT> tags
    
    Dim iCount As Integer
    
    ' There are 2 different types of Submit methods, POST and GET
    ' POST method has the form name/values in the message body whereas
    ' GET method shows it on the URL
    ' So we need to manipulate the parameters passed in differently based on the requested
    ' methods
    
    If SubmitMethod = enmPOST Then
    
            ' Set the URL property to the Requested URL
            mURL = URLString
            
            ' Set the method of request which is POST and the URL,and set the Async parameter to false
            objXMLHTTP.open "POST", URLString, False
            
            ' Sets the header so that the web server knows a form is going to be posted
            objXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        
            ' Construct the message body first before we send, it is a name/value pair,separated by ampersands
            ' which looks like "username=skeevs&password=aaa"
            For iCount = 0 To UBound(FormInputArray)
                strbody = strbody & FormInputArray(iCount) & "=" & FormValueArray(iCount) & "&"
            Next
            
            ' Need to remove the last ampersand which was added in the FOR Loop above, so we use the LEFT function
            objXMLHTTP.send Left(strbody, Len(strbody) - 1)
        
    ElseIf SubmitMethod = enmGET Then
        
            Dim strURLParam As String
            Dim strURL As String
            
            ' Construct the URL first , it is a name/value pair,separated by ampersands
            For iCount = 0 To UBound(FormInputArray)
                strURLParam = strURLParam & FormInputArray(iCount) & "=" & FormValueArray(iCount) & "&"
            Next
            
            ' Need to remove the last ampersand which was added in the FOR Loop above, so we use the LEFT function
            strURLParam = Left(strURLParam, Len(strURLParam) - 1)
            
            ' Construct the URL string for the GET method
            strURL = URLString & "?" & strURLParam
            
            ' Set the URL property to the Requested URL
            mURL = strURL
            
            ' Set the method of request which is POST and the URL
            objXMLHTTP.open "GET", strURL, False
            
            ' Sets the header so that the web server knows a form is going to be posted
            objXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            'objxmlhttp.
            ' Send the constructed URL
            objXMLHTTP.send strURL
        
    End If
    
    SubmitDataAsString = objXMLHTTP.responseText
    
End Function
Public Sub GetFile(strURL As String, strFile As String)

    'Dim xml As MSXML2.XMLHTTP30
    Dim X() As Byte
    Dim ff As Long
    strFile = Replace(strFile, "*", vbNullString)
    
    ff = FreeFile
     'Set xml = New MSXML2.XMLHTTP30
    objXMLHTTP.open "GET", strURL, True
    
    objXMLHTTP.send
    Do While objXMLHTTP.readyState <> 4
        DoEvents
    Loop
    X = objXMLHTTP.responseBody
    Open strFile For Binary Access Write As #ff
    Put #ff, , X()
    Close #ff
End Sub
