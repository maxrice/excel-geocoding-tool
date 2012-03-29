Attribute VB_Name = "mGeoCode"
Const LATITUDECOL = 1         ' column to put longitude into
Const LONGITUDECOL = 2        ' column to put latitude into
Const PRECISIONCOL = 3        ' column to put precision (quality index) into
Const LOCATIONCOL = 4         ' column to put location info into
Const FIRSTDATAROW = 13        ' rows above this row don't contain address data
                                'TODO - edit geocodenotfound() and cleardataentryarea() to use this constant for range instead of hardcoded
Const GOOGLEMAPSLINKCOL = 7    'Stores google maps link
Dim vProxyStatus As String      'stores query when behind proxy

' holds cache of strings submitted to geocoder during this session along with results
' to ensure that duplicate strings aren't submitted
Dim geocodeResults As New Collection



'TODO - edit this to reflect changes and add to README
' GEOCODING is done using the following layers
'
'geocodeSelectedRows
'(for each row call geocodeRow)
'
'       geocodeRow(r)
'       (check that row is geocodable, pass to geocode, parse results)
'
'           geocode(street,city,state,zip)
'           (clean all variables, pass url to geocoderAddressLookup,
'            if no result then try different permuatations of address)
'
'               geocoderAddressLookup
'               (query geocoder.us, return result, marshal results)
'



' submit selected rows to the geocoder
Sub geocodeSelectedRows()
    Dim r
    Call ProxyReload
    If [GeocoderToUse] = "Yahoo" Then
        If [yahooid] <> "" Then
            For Each r In Selection.rows()
                If r.Row() >= FIRSTDATAROW Then geocodeRow (r.Row())
            Next r
            Application.StatusBar = False
        Else:
            MsgBox "Please enter a Yahoo Id for geocoding"
        End If
    End If
End Sub

Sub geocodeNotFound()
    Dim r As Integer
    Call ProxyReload
    If [GeocoderToUse] = "Yahoo" Then
        Range("A13:C65536").Select
        Selection.Replace What:="not found", Replacement:="", LookAt:=xlPart, _
                          SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                          ReplaceFormat:=False
        Cells(FIRSTDATAROW, LATITUDECOL).Select
        If [yahooid] <> "" Then
            For r = FIRSTDATAROW To LastDataRow()
                geocodeRow (r)
            Next r
            Cells(FIRSTDATAROW, LATITUDECOL).Select
            Application.StatusBar = False
        Else:
            MsgBox "Please enter a Yahoo Id for geocoding"
        End If
    End If
End Sub

Sub geocodeAllRows()
    Dim r As Integer
    Call ProxyReload
    If [GeocoderToUse] = "Yahoo" Then
        Range("A13:C65536").Select
        Selection.ClearContents
        Range("J13:j65536").Select
        Selection.ClearContents
        Cells(FIRSTDATAROW, LATITUDECOL).Select
        If [yahooid] <> "" Then
            For r = FIRSTDATAROW To LastDataRow()
                geocodeRow (r)
            Next r
            Application.StatusBar = False
        Else:
            MsgBox "Please enter a Yahoo Id for geocoding"
        End If
    End If
End Sub

' geocode a single row of data
Sub geocodeRow(r As Integer)
    Dim resultstr As String
    Dim resultarray
    
    Application.StatusBar = "Geocoding row: " & r
    
    ' can't geocode if no address data
    ' nonblank latitude means we've already geocoded this row
    If Cells(r, LOCATIONCOL) <> "" And Cells(r, LATITUDECOL) = "" Then
    
    
        ' pass the street, city, state, and zip to the function geocode
        ' geocode returns a string containing the results in comma delimited format
        ' this is crude, but works
        ' CStr casts (converts) a value to a string
        resultstr = geocode(CStr(Cells(r, LOCATIONCOL)))
        
        ' parse the results, if lat/long/precision is blank, consider it not found
        resultarray = Split(resultstr, ",")
        If resultarray(0) = "" Then resultarray(0) = "not found"
        If resultarray(1) = "" Then resultarray(1) = "not found"
        If resultarray(2) = "" And resultarray(0) = "not found" Then resultarray(2) = "not found"
        
        ' store the results
        Cells(r, LATITUDECOL) = resultarray(0)
        Cells(r, LONGITUDECOL) = resultarray(1)
        Cells(r, PRECISIONCOL) = resultarray(2)
        Cells(r, GOOGLEMAPSLINKCOL).Value = "=HYPERLINK(""http://maps.google.com/maps?f=q&hl=en&geocode=&q=" & resultarray(0) & "," & resultarray(1) & """)"
    End If
End Sub


' remove everything following the start of the string trim
Function trimstr(basestr As String, trim As String) As String
    If InStr(basestr, trim) > 0 Then
        trimstr = Left(basestr, InStr(basestr, trim) - 1)
    Else
        trimstr = basestr
    End If
End Function


' remove everything following the end of the string trim
Function trimstrafter(basestr As String, trim As String) As String
    If InStr(basestr, trim) > 0 Then
        trimstrafter = Left(basestr, InStr(basestr, trim) + Len(trim) - 1)
    Else
        trimstrafter = basestr
    End If
End Function


Function geocode(location As String) As String
    Dim result As String
    
    'Geocode at yahoo using free-form addres format (see http://developer.yahoo.com/geo/placefinder/guide/requests.html#free-form-format)
    If [GeocoderToUse] = "Yahoo" Then
        result = yahooAddressLookup(location)
    End If

    geocode = result
End Function



Function yahooAddressLookup(location As String) As String
    ' perform RESTian lookup on Yahoo
    Dim marshalledResult As String
    Dim yahoo As String
    Dim response As String
    Dim result As String
    
    ' marshal the results of this very time consuming function
    ' see if we've already looked up this address
    ' turn error handling off
    On Error Resume Next
    ' lookup the result in the collection
    ' an error will be raised if the value is not found
    marshalledResult = geocodeResults(location)
    If marshalledResult <> "" Then
        ' if a value is found then return the result
        geocodeAddressLookup = marshalledResult
        Exit Function
    End If
    ' turn error handling back on
    On Error GoTo 0
    
    Application.StatusBar = "Looking for " & location
    yahoo = trim(CStr([yahooid]))
    
    street = trim(location)
    
    'flags=C only returns basic latitude/longitude/precision, excludes address parsing and other info
    URL = "http://where.yahooapis.com/geocode?q=" & URLEncode(location, True) & "&flags=C&appid=" & yahoo
    'Debug.Print URL
    
    If vProxyStatus = "Yes" Then
         'Create Http object
        If IsEmpty(Http) Then Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")

        'proxy HTTP -- code from:
        'http://forums.aspfree.com/visual-basic-programming-38/proxy-auth-in-this-vb-script-20625.html
    
        ' Set to use proxy -- see:
        ' http://msdn.microsoft.com/en-us/library/aa384059%28v=VS.85%29.aspx
        Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1
        Const HTTPREQUEST_PROXYSETTING_PROXY = 2
        Const AutoLogonPolicy_Always = 0
        
        Http.SetProxy HTTPREQUEST_PROXYSETTING_PROXY, [ProxyIP], "*.intra"
        Http.Open "GET", URL, False
        Http.SetAutoLogonPolicy AutoLogonPolicy_Always
    Else
        'Create Http object
        If IsEmpty(Http) Then Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
        'Send request To URL
         Http.Open "GET", URL
    End If
    
    Http.send           'TODO - error checking because of proxy
    'Get response data As a string
        
    response = Http.responseText
    'Debug.Print response
    
    ' capture the latitude by regex matching the values in the tags <geo:lat> and <geo:long>
    lat = RegExMatch(response, "<latitude>([\.\-0-9]+)</latitude>")
    lng = RegExMatch(response, "<longitude>([\.\-0-9]+)</longitude>")
    precision = RegExMatch(response, "<quality>([\.\-0-9]+)</quality>")
    
    ' return a comma delimited string
    ' if values not found, this will return ","
    yahooAddressLookup = lat & "," & lng & "," & precision
    
    
    ' store the result in the cache collection
    '
    ' turn off error handling with "On Error Resume Next"
    ' an error will be raised if you try to store to an address already in the cache
    ' we can ignore this error
    On Error Resume Next
    ' store the result
    geocodeResults(location) = lat & "," & lng
End Function

' wraps string with a tag
Function tag(xmltag As String, val As String) As String
    tag = "<" & xmltag & ">" & val & "</" & xmltag & ">" & vbCrLf
End Function


' basic distance function for latitude/longitude
Public Function latLongDistance(lat1 As Double, long1 As Double, lat2 As Double, long2 As Double) As Double
    Dim x As Double
    Dim y As Double
    x = 69.1 * (lat2 - lat1)
    y = 69.1 * (long2 - long1) * Cos(lat1 / 57.3)
    
    latLongDistance = (x * x + y * y) ^ 0.5
End Function


Private Function max(a, b):
    If a > b Then
        max = a
    Else
        max = b
    End If
End Function

Sub ClearDataEntryArea()
    Range("A13:J65536").Select
    Selection.ClearContents
    Range("A13").Select
End Sub

' locate the last row containing address data
Function LastDataRow() As Integer
    Dim r As Integer
    
    activecelladdr = ActiveCell.Address

    Range("d65536").End(xlUp).Select
    r = ActiveCell.Row()
    Range("e65536").End(xlUp).Select
    r = max(r, ActiveCell.Row())
    Range("f65536").End(xlUp).Select
    r = max(r, ActiveCell.Row())
    Range("g65536").End(xlUp).Select
    r = max(r, ActiveCell.Row())
    
    Range(activecelladdr).Select
    LastDataRow = r
End Function

Sub MacrosWorking()
    MsgBox "Macros are enabled."
End Sub

'Proxy functions
Sub CheckProxy()
    vProxyStatus = ""
    [ProxyStatusStorage].Value = ""
    Select Case MsgBox("Do you use a Proxy to access the internet?" _
                       & vbCrLf & "" _
                       & vbCrLf & "Reminder: Please configure the Proxy information if you have not done so." _
                       , vbYesNo Or vbExclamation Or vbSystemModal Or vbDefaultButton1, "Proxy Check")

        Case vbYes
            vProxyStatus = "Yes"
            [ProxyStatusStorage].Value = vProxyStatus
            ' Writes the proxy usage status to a cell in the workbook for later retrievel.
        Case vbNo
            vProxyStatus = "No"
            [ProxyStatusStorage].Value = vProxyStatus
    End Select


End Sub

Public Sub ProxyReload()
    vProxyStatus = [ProxyStatusStorage]
    'This simple code is used to reload the the proxy usage.  Depending on the circumestances, the code can forget the status.
End Sub
