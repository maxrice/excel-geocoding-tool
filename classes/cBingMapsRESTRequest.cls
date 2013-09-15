VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBingMapsRESTRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'MIT License
'Copyright 2012-2013 Max Rice (max@maxrice.com)
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files
'(the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify,
'merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished
'to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
'MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
'WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'
'Enjoy!

Option Explicit

'the URL to perform the request to
Public url As String

'the response xml
Public xml As String

'performs a REST lookup to Bing location service and parses the XML
Public Function performLookup(location As String)
    Dim data(2) As String
    
    url = ("http://dev.virtualearth.net/REST/v1/Locations?query=" & URLEncode(location, True) & "&maxResults=1&key=" & Trim(CStr(Range("bingMapsKey"))) & "&o=xml")
    
    'perform an HTTP GET
    xml = HTTPGET(url)
        
    'parse the XML to get the lat/long/confidence, note for Bing, this is only "high" / "medium" / "low" vs. numerical precision indicators like Yahoo or Google
    If (InStr(xml, "<Latitude>") <> 0 And InStr(xml, "<Longitude>") <> 0 And InStr(xml, "<Confidence>") <> 0) Then
        data(0) = getElementValue("Latitude")
        data(1) = getElementValue("Longitude")
        data(2) = getElementValue("Confidence")
    Else
        data(0) = "-"
        data(1) = "-"
        data(2) = "-"
    End If
    
    performLookup = Join(data, "|")
    
End Function

'return the request URI
Public Function getRequestURI()
    getRequestURI = url
End Function

'return the response XML
Public Function getResponseXML()
    getResponseXML = xml
End Function

'wrapper for the REST request to check proxy/OS
Public Function HTTPGET(url As String) As String
    Dim useProxy As Boolean
    
    If Range("UseProxy") = "Yes" Then
        useProxy = True
    Else
        useProxy = False
    End If

    If (isWindows()) Then
        HTTPGET = windowsHTTPGET(url, useProxy)
    Else
        HTTPGET = macHTTPGET(url, useProxy)
        
    End If

End Function

'perform the REST request on windows
Private Function windowsHTTPGET(url As String, useProxy As Boolean) As String
    Dim http As Object
    
    'create http object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    'proxy HTTP - http://forums.aspfree.com/visual-basic-programming-38/proxy-auth-in-this-vb-script-20625.html
    If useProxy Then

        ' Set to use proxy - http://msdn.microsoft.com/en-us/library/aa384059%28v=VS.85%29.aspx
        Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1
        Const HTTPREQUEST_PROXYSETTING_PROXY = 2
        Const AutoLogonPolicy_Always = 0
        
        http.SetProxy HTTPREQUEST_PROXYSETTING_PROXY, [ProxyIP], "*.intra"
        http.Open "GET", url, False
        http.SetAutoLogonPolicy AutoLogonPolicy_Always
    
    Else
        http.Open "GET", url
    End If
       
    'send the request
    http.send
        
    'get response data as a string
    windowsHTTPGET = http.responseText

End Function

'perform the REST request on mac
Private Function macHTTPGET(url As String, useProxy As Boolean) As String
    Dim script As String
    
    
    'build the shell script, starts with `curl '<url>' --silent`
    If useProxy Then
        script = "do shell script " & Chr(34) & "curl '" & url & "'" & " --silent --proxy " & Range("proxyIP") & Chr(34)
    Else
        script = "do shell script " & Chr(34) & "curl '" & url & "'" & " --silent" & Chr(34)
    End If
    
    'run the shell script
    macHTTPGET = MacScript(script)
    
End Function


'URL encode a string
'From http://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba
Private Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String

On Error GoTo Catch

  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"
    
    For i = 1 To StringLen
      Char = Mid(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
Finally:
Exit Function
Catch:
URLEncode = ""
Resume Finally

End Function

'Check if Excel is running in Windows or Mac
'From http://www.rondebruin.nl/mac.htm
Private Function isWindows() As Boolean

    'Test the OperatingSystem
    If Not Application.OperatingSystem Like "*Mac*" Then
        isWindows = True
    Else
        'Mac, but test if it is Excel 2011 or higher
        If Val(Application.Version) > 14 Then
            isWindows = False
        End If
    End If

End Function

'Get the element value in an XML document
'if excel for mac had regex support, we'd use that. it does not, so use these string functions to find lat/long/precision while maintaining win/mac compatibility
Private Function getElementValue(elementName As String)
    Dim element As String
    Dim startPosition As Long
    Dim endPosition As Long
    Dim elementLength As Long

    
    'find the start position of the start tag <ElementName> and add the length of the element to the position
    element = "<" & elementName & ">"
    startPosition = InStr(xml, element) + Len(element)
    
    'find the start position of the closing tag </ElementName>
    element = "</" & elementName & ">"
    endPosition = InStr(xml, element)
    
    'calculate the length of the element value
    elementLength = endPosition - startPosition
    
    'return the sub-string
    getElementValue = Mid(xml, startPosition, elementLength)
    
End Function


