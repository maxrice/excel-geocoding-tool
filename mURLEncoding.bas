Attribute VB_Name = "mURLEncoding"
'**************************************
' Name: URLEncode Function
' Description:Encodes a string to create legally formatted
' QueryString for URL. This function is more flexible
' than the IIS Server.Encode function because you can
' pass in the WHOLE URL and only the QueryString data
' will be converted. IIS strangely converts EVERYTHING
'(ie "http://" becomes "http%3A%2F%2F").
' By: Markus Diersbock
'
' Inputs:sRawURL - String to Encode
'
' Returns:Encoded String
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=43806&lngWId=1'for details.'**************************************

Public Function URLEncode(sRawURL As String) As String
On Error GoTo Catch
Dim iLoop As Integer
Dim sRtn As String
Dim sTmp As String
Const sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz:/.?=_-$(){}~&"
If Len(sRawURL) > 0 Then
' Loop through each char
For iLoop = 1 To Len(sRawURL)
sTmp = Mid(sRawURL, iLoop, 1)
If InStr(1, sValidChars, sTmp, vbBinaryCompare) = 0 Then
' If not ValidChar, convert to HEX and prefix with %
sTmp = Hex(Asc(sTmp))
If sTmp = "20" Then
sTmp = "+"
ElseIf Len(sTmp) = 1 Then
sTmp = "%0" & sTmp
Else
sTmp = "%" & sTmp
End If
End If
sRtn = sRtn & sTmp
Next iLoop
URLEncode = sRtn
End If
Finally:
Exit Function
Catch:
URLEncode = ""
Resume Finally
End Function
