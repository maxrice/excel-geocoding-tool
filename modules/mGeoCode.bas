Attribute VB_Name = "mGeoCode"
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

Const LATITUDECOL = 1               'column to put longitude into
Const LONGITUDECOL = 2              'column to put latitude into
Const PRECISIONCOL = 3              'column to put precision (quality index) into
Const LOCATIONCOL = 4               'column to put location info into
Const FIRSTDATAROW = 13             'rows above this row don't contain address data
Const GOOGLEMAPSLINKCOL = 7         'column to store google maps link
Const DEBUGMODEQUERYCOL = 10        'column to store HTTP query if debug mode is on
Const DEBUGMODERESPONSECOL = 11     'column to store Response XML if debug mode is on

'Global query/response variables for debugging
Dim debugMode As Boolean
Dim debugModeQuery As String
Dim debugModeResponse As String


' geocode only selected rows
Sub geocodeSelectedRows()
    
    If checkSettings = True Then
        
        Dim r
        For Each r In Selection.rows()
            If r.Row() >= FIRSTDATAROW Then
                geocodeRow (r.Row())
            End If
        Next r
            
        Application.StatusBar = False
        
    End If

End Sub

Sub geocodeNotFound()
    
    If checkSettings = True Then
        
        
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
        
        'Now geocode
        Dim r As Integer
        For r = FIRSTDATAROW To LastDataRow()
            geocodeRow (r)
        Next r
        
        Cells(FIRSTDATAROW, LATITUDECOL).Select
        Application.StatusBar = False
        
    End If

End Sub

Sub geocodeAllRows()
    
    If checkSettings = True Then
    
        Dim r As Integer
        Range("A13:C65536").Select
        Selection.ClearContents
        Range("J13:j65536").Select
        Selection.ClearContents
        Cells(FIRSTDATAROW, LATITUDECOL).Select
        
        For r = FIRSTDATAROW To LastDataRow()
            geocodeRow (r)
        Next r
        
        Application.StatusBar = False
       
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
        resultstr = Geocode(CStr(Cells(r, LOCATIONCOL)))
        
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
        If debugMode = True Then
            Cells(r, DEBUGMODEQUERYCOL).Value = debugModeQuery
            Cells(r, DEBUGMODERESPONSECOL).Value = debugModeResponse
            Cells(r, DEBUGMODERESPONSECOL).WrapText = False
        End If
    End If
End Sub

Function Geocode(location As String) As String

    'Geocode at yahoo using free-form addres format (see http://developer.yahoo.com/geo/placefinder/guide/requests.html#free-form-format)

    Geocode = bingAddressLookup(location)

End Function


'Perform REST lookup on Bing
Function bingAddressLookup(location As String) As String
    On Error GoTo catchError:
    Dim bingMapsKey As String
    Dim response As String
    Dim geo As cGeocode
    Dim url As String
    Dim lat As String
    Dim lng As String
    
    Set geo = New cGeocode
    
    Application.StatusBar = "Looking for " & location
    bingMapsKey = Trim(CStr(Range("bingMapsKey")))

    'set the URL
    url = "http://dev.virtualearth.net/REST/v1/Locations?query=" & URLEncode(location, True) & "&maxResults=1&key=" & bingMapsKey
    
    'Log query if debug mode is on
    If debugMode = True Then debugModeQuery = url
   
    'Get the response via HTTP GET & use a proxy if required
    If Range("UseProxy") = "Yes" Then
        response = HTTPGet(url, True)
    Else
        response = HTTPGet(url, False)
    End If
    
    'Log result if debug mode is on
    If debugMode = True Then debugModeResponse = response
    
    'parse the response JSON
    geo.parseResponse (CStr(response))
    
    'return the lat/long/precision
    bingAddressLookup = geo.getLatitude() & "," & geo.getLongitude() & "," & geo.getPrecision()
    
catchError:
    bingAddressLookup = ",,"
    
End Function

Function checkSettings()
   
    'Check if Yahoo is selected as geocoder and API key is not blank
    If Range("GeocoderToUse") = "Bing" Then
        If Range("bingMapsKey") <> "" Then
            
            'Set debug mode flag if setting is enabled
            If Range("DebugMode") = "On" Then
                debugMode = True
            Else:
                debugMode = False
            End If
            
            'Ready to Geocode
            checkSettings = True
        Else:
            MsgBox "Please enter a Bing Maps Key for geocoding"
            'Not ready to geocode
            checkSettings = False
        End If
    End If

End Function

Sub ClearDataEntryArea()
    Range("A13:K65536").Select
    Selection.ClearContents
    Range("A13").Select
End Sub

Private Function max(a, B):
    If a > B Then
        max = a
    Else
        max = B
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

