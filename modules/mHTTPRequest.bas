Attribute VB_Name = "mHTTPRequest"
'Copyright 2012-2013 Max Rice, Juice Analytics
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

'Perform a HTTP GET for the passed URL
Public Function HTTPGet(url As String, Optional UseProxy As Boolean = False) As String
    Dim http As Object
    Dim script As String
    
    If WinOrMac = "win" Then
    'Windows HTTP Request
        If UseProxy = True Then
             'Create Http object
            Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
            'proxy HTTP
            'from http://forums.aspfree.com/visual-basic-programming-38/proxy-auth-in-this-vb-script-20625.html
        
            ' Set to use proxy -- see:
            ' http://msdn.microsoft.com/en-us/library/aa384059%28v=VS.85%29.aspx
            Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1
            Const HTTPREQUEST_PROXYSETTING_PROXY = 2
            Const AutoLogonPolicy_Always = 0
            
            http.SetProxy HTTPREQUEST_PROXYSETTING_PROXY, [ProxyIP], "*.intra"
            http.Open "GET", url, False
            http.SetAutoLogonPolicy AutoLogonPolicy_Always
        Else
            'Create Http object
            Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        
            'Send request To URL
             http.Open "GET", url
        End If
       
       'TODO - error checking because of proxy
        http.send
        
        'Get response data As a string
        HTTPGet = http.responseText
        
    Else
    'Mac HTTP Request
        If UseProxy = True Then
            script = "do shell script " & Chr(34) & "curl " & url & " --proxy " & Range("ProxyIP") & Chr(34)
        Else
            script = "do shell script " & Chr(34) & "curl " & url & Chr(34)
        End If
        'Debug.Print script
        
        'TODO - error catch
        HTTPGet = MacScript(script)
        
    End If
    
End Function

Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String

'From http://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba
'with edits for error catching

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

Public Function WinOrMac() As String
'From http://www.rondebruin.nl/mac.htm
'Test the OperatingSystem
    If Not Application.OperatingSystem Like "*Mac*" Then
        WinOrMac = "win"
    Else
        'I am a Mac and will test if it is Excel 2011 or higher
        If Val(Application.Version) > 14 Then
            WinOrMac = "mac"
        End If
    End If
End Function

