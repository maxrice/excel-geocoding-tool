Attribute VB_Name = "mHTTPRequest"
Public Function HTTPGet(url As String, Optional UseProxy As Boolean = False) As String
    
    If WinOrMac = "win" Then
    'Windows HTTP Request
        If UseProxy = True Then
             'Create Http object
            If IsEmpty(Http) Then Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
            'proxy HTTP
            'from http://forums.aspfree.com/visual-basic-programming-38/proxy-auth-in-this-vb-script-20625.html
        
            ' Set to use proxy -- see:
            ' http://msdn.microsoft.com/en-us/library/aa384059%28v=VS.85%29.aspx
            Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1
            Const HTTPREQUEST_PROXYSETTING_PROXY = 2
            Const AutoLogonPolicy_Always = 0
            
            Http.SetProxy HTTPREQUEST_PROXYSETTING_PROXY, [ProxyIP], "*.intra"
            Http.Open "GET", url, False
            Http.SetAutoLogonPolicy AutoLogonPolicy_Always
        Else
            'Create Http object
            If IsEmpty(Http) Then Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")
        
            'Send request To URL
             Http.Open "GET", url
        End If
       
       'TODO - error checking because of proxy
        Http.send
        
        'Get response data As a string
        HTTPGet = Http.responseText
        
    Else
    'Mac HTTP Request
        If UseProxy = True Then
            'proxy setup for curl
        End If
        
        script = "do shell script " & Chr(34) & "curl " & url & Chr(34)
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
        If val(Application.Version) > 14 Then
            WinOrMac = "mac"
        End If
    End If
End Function

