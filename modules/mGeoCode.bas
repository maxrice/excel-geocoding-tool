Attribute VB_Name = "mGeoCode"
'Copyright 2012 Max Rice, Juice Analytics
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

Const LATITUDECOL = 1         ' column to put longitude into
Const LONGITUDECOL = 2        ' column to put latitude into
Const PRECISIONCOL = 3        ' column to put precision (quality index) into
Const LOCATIONCOL = 4         ' column to put location info into
Const FIRSTDATAROW = 13        ' rows above this row don't contain address data
Const GOOGLEMAPSLINKCOL = 7    'Stores google maps link

' holds cache of strings submitted to geocoder during this session along with results
' to ensure that duplicate strings aren't submitted
Dim geocodeResults As New Collection


' submit selected rows to the geocoder
Sub geocodeSelectedRows()
    Dim r
    If Range("GeocoderToUse") = "Yahoo" Then
        If Range("YahooID") <> "" Then
            For Each r In Selection.Rows()
                If r.Row() >= FIRSTDATAROW Then geocodeRow (r.Row())
            Next r
            Application.StatusBar = False
        Else:
            MsgBox "Please enter a Yahoo ID for geocoding"
        End If
    End If
End Sub

Sub geocodeNotFound()
    Dim r As Integer
    If Range("GeocoderToUse") = "Yahoo" Then

        'Loop through result range and remove "not found" cells
        'This is much easier with range.replace, but the function parameters are different between win/mac, which makes it unusable for us. The joys of cross-compatibility :)
        Dim Row As Long, Column As Long
        
        For Row = FIRSTDATAROW To 65536
            For Column = LATITUDECOL To PRECISIONCOL
                If Cells(Row, Column).Value = "not found" Then
                    Cells(Row, Column).Value = ""
                End If
            Next Column
        Next Row

        Cells(FIRSTDATAROW, LATITUDECOL).Select
        If Range("YahooID") <> "" Then
            For r = FIRSTDATAROW To LastDataRow()
                geocodeRow (r)
            Next r
            Cells(FIRSTDATAROW, LATITUDECOL).Select
            Application.StatusBar = False
        Else:
            MsgBox "Please enter a Yahoo ID for geocoding"
        End If
    End If
End Sub

Sub geocodeAllRows()
    Dim r As Integer
    If Range("GeocoderToUse") = "Yahoo" Then
        Range("A13:C65536").Select
        Selection.ClearContents
        Range("J13:j65536").Select
        Selection.ClearContents
        Cells(FIRSTDATAROW, LATITUDECOL).Select
        If Range("YahooID") <> "" Then
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
    
    
        ' pass the location to geocode
        ' geocode returns a string containing the results in comma delimited format
        resultstr = geoCode(CStr(Cells(r, LOCATIONCOL)))
        
        ' parse the results, if lat/long/precision is blank, consider it not found
        resultarray = Split(resultstr, ",")
        If resultarray(0) = "" Then resultarray(0) = "not found"
        If resultarray(1) = "" Then resultarray(1) = "not found"
        If resultarray(2) = "" And resultarray(0) = "not found" Then resultarray(2) = "not found"
        
        ' store the results
        Cells(r, LATITUDECOL) = resultarray(0)
        Cells(r, LONGITUDECOL) = resultarray(1)
        Cells(r, PRECISIONCOL) = resultarray(2)
        If Cells(r, LATITUDECOL) <> "not found" Then
            Cells(r, GOOGLEMAPSLINKCOL).Value = "=HYPERLINK(""http://maps.google.com/maps?f=q&hl=en&geocode=&q=" & resultarray(0) & "," & resultarray(1) & """)"
        End If
    End If
End Sub

Function geoCode(location As String) As String
    
    Dim result As String
    
    'Geocode at yahoo using free-form addres format (see http://developer.yahoo.com/geo/placefinder/guide/requests.html#free-form-format)
    If Range("GeocoderToUse") = "Yahoo" Then
        result = yahooAddressLookup(location)
    End If

    geoCode = result
    
End Function

Function yahooAddressLookup(location As String) As String
    ' perform RESTian lookup on Yahoo
    Dim marshalledResult As String
    Dim yahoo As String
    Dim response As String
    Dim url As String
    Dim lat As String
    Dim lng As String
    Dim precision As String
    
    ' marshal the results of this very time consuming function
    ' see if we've already looked up this address
    ' turn error handling off
    On Error Resume Next
    ' lookup the result in the collection
    ' an error will be raised if the value is not found
    marshalledResult = geocodeResults(location)
    If marshalledResult <> "" Then
        ' if a value is found then return the result
        yahooAddressLookup = marshalledResult
        Exit Function
    End If
    ' turn error handling back on
    On Error GoTo 0
    
    Application.StatusBar = "Looking for " & location
    yahoo = trim(CStr(Range("YahooID")))

    
    'flags=C only returns basic latitude/longitude/precision, excludes address parsing and other info
    url = "http://where.yahooapis.com/geocode?q=" & URLEncode(location, True) & "%26flags=C%26appid=" & yahoo
    'Debug.Print URL
   
    'Get the response via HTTP GET & use a proxy if required
    If Range("UseProxy") = "Yes" Then
        response = HTTPGet(url, True)
    Else
        response = HTTPGet(url, False)
    End If
    
    'Debug.Print response
    
    'Yahoo will return multiple results if it found more than 1 good match
    If Mid(response, (InStr(1, response, "<Found>", vbTextCompare) + 7), (InStr(1, response, "</Found>", vbTextCompare) - 7 - InStr(1, response, "<Found>", vbTextCompare))) > 0 Then
        'Found
        'if excel for mac had regex support, we'd use that. it does not, so use these silly functions to find lat/long/quality while maintaining win/mac compatibility
        lat = Mid(response, (InStr(1, response, "<latitude>", vbTextCompare) + 10), (InStr(1, response, "</latitude>", vbTextCompare) - 10 - InStr(1, response, "<latitude>", vbTextCompare)))
        lng = Mid(response, (InStr(1, response, "<longitude>", vbTextCompare) + 11), (InStr(1, response, "</longitude>", vbTextCompare) - 11 - InStr(1, response, "<longitude>", vbTextCompare)))
        precision = Mid(response, (InStr(1, response, "<quality>", vbTextCompare) + 9), (InStr(1, response, "</quality>", vbTextCompare) - 9 - InStr(1, response, "<quality>", vbTextCompare)))
        
        'return csv
        yahooAddressLookup = lat & "," & lng & "," & precision
        
    Else
        'Not found
        yahooAddressLookup = ",,"
   
    End If
    
    
    ' store the result in the cache collection
    '
    ' turn off error handling with "On Error Resume Next"
    ' an error will be raised if you try to store to an address already in the cache
    ' we can ignore this error
    On Error Resume Next
    ' store the result
    geocodeResults(location) = lat & "," & lng
End Function

Sub ClearDataEntryArea()
    Range("A13:J65536").Select
    Selection.ClearContents
    Range("A13").Select
End Sub

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
    Dim activecelladdr As String
    
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

