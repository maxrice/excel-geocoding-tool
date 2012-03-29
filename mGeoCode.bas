Attribute VB_Name = "mGeoCode"
Const LATITUDECOL = 1         ' column to put longitude into
Const LONGITUDECOL = 2        ' column to put latitude into
Const PRECISIONCOL = 3
Const LOCATIONCOL = 4
Const FIRSTDATAROW = 6        ' rows above this row don't contain address data


' holds cache of strings submitted to geocoder during this session along with results
' to ensure that duplicate strings aren't submitted
Dim geocodeResults As New Collection




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


Sub geocodeAllRows()
    Dim r As Integer
    If [GeocoderToUse] = "Yahoo" Then
        If [yahooid] <> "" Then
            For r = FIRSTDATAROW To LastDataRow()
                geocodeRow (r)
            Next r
            Application.StatusBar = False
        Else:
            MsgBox "Please enter a Yahoo Id for geocoding"
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
    marshalledResult = geocodeResults(addr)
    If marshalledResult <> "" Then
        ' if a value is found then return the result
        geocodeAddressLookup = marshalledResult
        Exit Function
    End If
    ' turn error handling back on
    On Error GoTo 0
    
    Application.StatusBar = "Looking for " & street & ", " & city & ", " & state & " " & zip
    yahoo = trim(CStr([yahooid]))
    
    street = trim(location)
    
    'flags=C only returns basic latitude/longitude/precision, excludes address parsing and other info
    URL = "http://where.yahooapis.com/geocode?q=" & URLEncode(location, True) & "&flags=C&appid=" & yahoo
    
    'Create Http object
    If IsEmpty(http) Then Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    'Send request To URL
    http.Open "GET", URL
    
    http.send
    'Get response data As a string
        
    response = http.responseText

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
    geocodeResults(addr) = lat & "," & lng
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

