'###################################################################################
'  Script name:		IE_Google_GetCoordinates.vbs
'  Created on:		5/17/2010
'  Author:			Dennis Hemken
'  Purpose:			Opens Microsoft Internet Explorer and navigate to Google Maps
'					with the parameter place, street and country.
'					The function returns the coordinates.
'###################################################################################

Dim strResult

' Alexanderplatz in Berlin, Germany
' strCoordinates = getCoordinatesCSV("Berlin","Alexanderplatz","Deutschland")
' Golden Gate Hotel in Las Vegas, USA
' strCoordinates = getCoordinatesCSV("Las Vegas","N Main St","USA")
' Taj Mahal in India
' strCoordinates = getCoordinatesCSV("Taj Mahal","","India")
' Royal Botanic Gardens in Australia
' strCoordinates = getCoordinatesCSV("Royal Botanic Gardens","","Australia")
' The vulcan Eyjafjallajökull in Island
 strResult = getCoordinatesCSV("Eyjafjalla Glacier","","Island")
 
wscript.echo strResult

Function getCoordinatesCSV(strPlace, strStreet, strCountry)
Dim IEApp
Dim IEDocument
Dim strArr
Dim strCoordinates
    Set IEApp = CreateObject("InternetExplorer.Application")
    With IEApp
        .Visible = False
        .Navigate "http://maps.google.com/maps/geo?q=" & strPlace & "%20" & strStreet & "%20" & strCountry & "&output=csv"
        Do: Loop Until .Busy = False
        Do: Loop Until .Busy = False
        While IEApp.Busy: Wend
        Set IEDocument = .Document
    End With

    strArr = Split(IEDocument.Body.innerText, ",")

    strCoordinates = strArr(3) & "," & strArr(2)
	getCoordinatesCSV = strCoordinates
	
    IEApp.Quit
Set IEDocument = Nothing
Set IEApp = Nothing
End Function