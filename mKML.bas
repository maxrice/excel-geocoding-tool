Attribute VB_Name = "mKML"
Public folders As New Collection
Const FIRSTDATAROW = 6
Const LATITUDECOL = 1
Const LONGITUDECOL = 2
Const MARKERCOLORCOL = 10
Const MARKERSIZECOL = 10
Const MARKERIMAGECOL = 10
Const LABELCOLORCOL = 10
Const LABELSIZECOL = 10
Const NAMECOL = 3
Const DESCRIPTIONCOL = 4
Const FOLDERCOL = 10


' cheap and cheerful templating
Const PLACEMARKTEMPLATE = _
    "<Placemark>%CR%" & _
    "  <description>%description%</description>%CR%" & _
    "  <name>%name%</name>%CR%" & _
    "  <Style>%CR%" & _
    "  <IconStyle><scale>0.5</scale></IconStyle>" & _
    "%buttontemplate%" & _
    "%labeltemplate%" & _
    "  </Style>%CR%" & _
    "  <visibility>0</visibility>%CR%" & _
    "  <Point>%CR%" & _
    "    <coordinates>%longitude%,%latitude%,0</coordinates>%CR%" & _
    "  </Point>%CR%" & _
    "</Placemark>%CR%%CR%"

Const BUTTONTEMPLATE = _
    "    <IconStyle>%CR%" & _
    "%buttoncolortemplate%" & _
    "%buttonscaletemplate%" & _
    "      <Icon><href>http://www.juiceanalytics.com/images/buttons/%buttonimage%.png</href></Icon>%CR%" & _
    "    </IconStyle>%CR%"
    
    
Const BUTTONCOLORTEMPLATE = "      <color>ff%buttoncolor%</color>%CR%"
Const BUTTONSCALETEMPLATE = "      <scale>%buttonscale%</scale>%CR%"

Const LABELTEMPLATE = "    <LabelStyle>%labelcolortemplate%%labelscaletemplate%</LabelStyle>%CR%"

Const LABELCOLORTEMPLATE = "<color>ff%labelcolor%</color>"
Const LABELSCALETEMPLATE = "<scale>%labelscale%</scale>"





Sub OutputKML()
    Dim s As String
    Dim name As String
    Dim description As String
    Dim markercolor As String
    Dim markerscale As String
    Dim markerimage As String
    Dim labelcolor As String
    Dim labelscale As String
    Dim latitude As String
    Dim longitude As String
    
    Dim folder As String
    Dim prevfolder As String
    Dim lastrow As Integer
    
    
    Dim sFileName As String
    'Show the open dialog and pass the selected file name
    ' to the String variable "sFileName"
    sFileName = Application.GetSaveAsFilename("output.kml", "KML Files (*.kml),*.kml", 1, "Where do you want to save your Google Earth file?", "Save")
    
    ' if the user cancelled
    If sFileName = "False" Then Exit Sub
    
    Open CStr(sFileName) For Output As #1
    
    ' do all folders
    prevfolder = "***BLANK***"
    
    Print #1, StartKML
    
    lastrow = LastDataRow
    
    For r = FIRSTDATAROW To lastrow
        
        name = CStr(ActiveSheet.Cells(r, 8))
        description = CStr(ActiveSheet.Cells(r, 9))
        
        markercolor = ActiveSheet.Cells(r, 255)
        markerscale = "0.5"
        markerimage = ActiveSheet.Cells(r, 255)
        labelcolor = ActiveSheet.Cells(r, 255)
        labelscale = "0.5"
        folder = ActiveSheet.Cells(r, 255)
        
        latitude = ActiveSheet.Cells(r, LATITUDECOL)
        longitude = ActiveSheet.Cells(r, LONGITUDECOL)
        
        If folder <> prevfolder Then
            If prevfolder <> "***BLANK***" Then Print #1, EndFolder
            Print #1, StartFolder(folder)
        End If
        prevfolder = folder
        
        If latitude <> "" And latitude <> "not found" Then
            Print #1, KMLMakePlacemarkString(name, description, markerimage, markercolor, markerscale, labelcolor, labelscale, latitude, longitude)
        End If
    Next r
    
    Print #1, EndFolder
    Print #1, EndKML
    Close #1
    
    Shell (CStr([GoogleEarthExecutableLocation]) & " " & sFileName)
End Sub




Function StartKML() As String
    StartKML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                    "<kml xmlns=""http://earth.google.com/kml/2.0"">" & vbCrLf & "<Document>" & vbCrLf
End Function


Function EndKML() As String
    EndKML = "</Document>" & vbCrLf & "</kml>" & vbCrLf
End Function


Function StartFolder(folderName As String) As String
    StartFolder = "   <Folder>" & vbCrLf & _
                  "      <name>" & folderName & "</name>" & vbCrLf & _
                  "      <visibility>0</visibility>" & vbCrLf & _
                  "      <open>0</open>" & vbCrLf
End Function



Function EndFolder() As String
    EndFolder = "   </Folder>" & vbCrLf
End Function



Function KMLFromRow()
    r = ActiveCell.Row()
    
    Debug.Print KMLMakePlacemark(CStr(Cells(r, 10)), _
                                 CStr(Cells(r, 11)), _
                                 CStr(Cells(r, 9)), _
                                 CStr(Cells(r, 6)), _
                                 CStr(Cells(r, 7)))
                                 
End Function




Function template(templatestr As String, replacements As Collection) As String
    Dim findreplace
    Dim strFind As String
    Dim strReplace As String
    
    For Each findreplace In replacements
        strFind = findreplace(0)
        strReplace = findreplace(1)
        templatestr = Replace(templatestr, "%" & strFind & "%", strReplace)
    Next findreplace
    template = templatestr
End Function








Function KMLMakePlacemarkString(name As String, _
                                description As String, _
                                buttonimage As String, _
                                buttoncolor As String, _
                                buttonscale As String, _
                                labelcolor As String, _
                                labelscale As String, _
                                latitude As String, _
                                longitude As String) As String
    
    
    name = RegExValidate(name, "[a-zA-Z0-9 ]")
    'description = RegExValidate(description, "[a-zA-Z0-9,\(\)<>!\[\] ]")
    


    Dim repl As New Collection
        
    repl.Add Array("description", description)
    repl.Add Array("name", name)
    
    If buttonimage <> "" Then
        repl.Add Array("buttontemplate", BUTTONTEMPLATE)
        If buttoncolor <> "" Then repl.Add Array("buttoncolortemplate", BUTTONCOLORTEMPLATE)
        If buttonscale <> "" Then repl.Add Array("buttonscaletemplate", BUTTONSCALETEMPLATE)
    End If
    repl.Add Array("buttonimage", buttonimage)
    repl.Add Array("buttoncolor", buttoncolor)
    repl.Add Array("buttonscale", buttonscale)
    repl.Add Array("buttoncolortemplate", "")
    repl.Add Array("buttonscaletemplate", "")
    repl.Add Array("buttontemplate", "")
    
    If labelcolor <> "" Or labelscale <> "" Then
        repl.Add Array("labeltemplate", LABELTEMPLATE)
        If labelcolor <> "" Then repl.Add Array("labelcolortemplate", LABELCOLORTEMPLATE)
        If labelscale <> "" Then repl.Add Array("labelscaletemplate", LABELSCALETEMPLATE)
    End If
    repl.Add Array("labelcolor", labelcolor)
    repl.Add Array("labelscale", labelscale)
    repl.Add Array("labelcolortemplate", "")
    repl.Add Array("labelscaletemplate", "")
    repl.Add Array("labeltemplate", "")
    
    repl.Add Array("latitude", latitude)
    repl.Add Array("longitude", longitude)
    repl.Add Array("CR", vbCrLf)
    
    KMLMakePlacemarkString = template(PLACEMARKTEMPLATE, repl)
    
End Function

