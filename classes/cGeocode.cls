VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGeocode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'the decoded JSON
Public jsonObject As Object

'JavaScript script engine
Public scriptEngine As Object

'Setup the script engine and define the functions upon class instantiation
Private Sub Class_Initialize()

    Set Me.scriptEngine = New ScriptControl
    Me.scriptEngine.Language = "JScript"
    Me.scriptEngine.AddCode "function getLatitude(jsonObj) { return jsonObj.resourceSets[0].resources[0].geocodePoints[0].coordinates[0]; } "
    Me.scriptEngine.AddCode "function getLongitude(jsonObj) { return jsonObj.resourceSets[0].resources[0].geocodePoints[0].coordinates[1]; } "
    Me.scriptEngine.AddCode "function getPrecision(jsonObj) { return jsonObj.resourceSets[0].resources[0].confidence; } "

End Sub

'Parse the passed JSON into a javascript object
Public Function parseResponse(ByVal jsonString As String)
    Set jsonObject = scriptEngine.Eval("(" + jsonString + ")")
End Function

'get the latitude
Public Function getLatitude() As String
    getLatitude = scriptEngine.Run("getLatitude", Me.jsonObject)
End Function

'get the longitude
Public Function getLongitude() As String
    getLongitude = scriptEngine.Run("getLongitude", Me.jsonObject)
End Function

'get the Precision - note for Bing, this is only "high" / "medium" / "low" vs. numerical precision indicators like Yahoo or Google
Public Function getPrecision() As String
    getPrecision = scriptEngine.Run("getPrecision", Me.jsonObject)
End Function
