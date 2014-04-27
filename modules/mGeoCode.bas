Attribute VB_Name = "mGeoCode"
'MIT License
'Copyright 2012-2013 Max Rice (max@maxrice.com), Juice Analytics
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
Const CONFIDENCECOL = 3             'column to put confidence indicator into
Const LOCATIONCOL = 4               'column to put location info into
Const FIRSTDATAROW = 13             'rows above this row don't contain address data
Const GOOGLEMAPSLINKCOL = 7         'column to store google maps link
Const DEBUGMODEREQUESTCOL = 10      'column to store request URI if debug mode is on
Const DEBUGMODERESPONSECOL = 11     'column to store response JSON  if debug mode is on

'Global request/response variables for debugging
Dim debugMode As Boolean
Dim debugModeRequest As String
Dim debugModeResponse As String


'geocode only selected rows
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

'geocode rows listed as "not found"
Sub geocodeNotFound()
    
    If checkSettings = True Then
        
        'Loop through result range and remove "not found" cells
        'This is much easier with range.replace, but the function parameters are different between win/mac, which makes it unusable for us. The joys of cross-compatibility :)
        Dim Row As Long, Column As Long
        For Row = FIRSTDATAROW To 65536
            For Column = LATITUDECOL To CONFIDENCECOL
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

'geocode ALL THE ROWS!
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

'geocode a single row of data
Sub geocodeRow(r As Integer)
    Dim rawGeocodeData As String
    Dim geocodeData
    Dim latitude As String
    Dim longitude As String
    Dim confidence As String
    
    Application.StatusBar = "Geocoding row: " & r
    
    'can't geocode if no address data
    'nonblank latitude means we've already geocoded this row
    If Cells(r, LOCATIONCOL) <> "" And Cells(r, LATITUDECOL) = "" Then
    
        ' pass the location to geocode
        ' bingAddressLookup returns an array containing the lat/long/confidence
        rawGeocodeData = bingAddressLookup(CStr(Cells(r, LOCATIONCOL)))
        
        geocodeData = Split(rawGeocodeData, "|")
        
        'set lat/long/confidence
        latitude = geocodeData(0)
        longitude = geocodeData(1)
        confidence = geocodeData(2)
        
        'if lat/long/confidence is blank, consider it not found
        If latitude = "-" Then latitude = "not found"
        If longitude = "-" Then longitude = "not found"
        If confidence = "-" Then confidence = "not found"

        ' store the results
        Cells(r, LATITUDECOL) = latitude
        Cells(r, LONGITUDECOL) = longitude
        Cells(r, CONFIDENCECOL) = confidence
        
        'add google maps link
        If Cells(r, LATITUDECOL) <> "not found" Then
            Cells(r, GOOGLEMAPSLINKCOL).Value = "=HYPERLINK(""http://maps.google.com/maps?f=q&hl=en&geocode=&q=" & latitude & "," & longitude & """)"
        End If
        
        'add logs if enabled
        If debugMode = True Then
            Cells(r, DEBUGMODEREQUESTCOL).Value = debugModeRequest
            Cells(r, DEBUGMODERESPONSECOL).Value = debugModeResponse
            Cells(r, DEBUGMODERESPONSECOL).WrapText = False
        End If
        
    End If
    
End Sub

'Perform REST lookup on Bing
Function bingAddressLookup(location As String) As String
    On Error Resume Next
    Dim bing As New cBingMapsRESTRequest
    Dim geocodeData As String

    Application.StatusBar = "Looking for " & location
    
    'perform the lookup
    geocodeData = bing.performLookup(location)
    
    'log response/request
    If (debugMode) Then
        debugModeRequest = bing.getRequestURI
        debugModeResponse = bing.getResponseXML
    End If
    
    'return the lat/long/confidence
    bingAddressLookup = geocodeData
    
End Function

'check that all settings are valid
Function checkSettings()
   
    'Check if Bing is selected as geocoder and API key is not blank
    If Range("GeocoderToUse") = "Bing" Then
        If Range("bingMapsKey") <> "" Then
            
            'Set debug mode flag if setting is enabled
            If Range("DebugMode") = "On" Then
                debugMode = True
            Else
                debugMode = False
            End If
            
            'Ready to Geocode
            checkSettings = True
        
        Else
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

'Ensure that macros are working properly
Sub MacrosWorking()
    MsgBox "Macros are enabled."
End Sub

