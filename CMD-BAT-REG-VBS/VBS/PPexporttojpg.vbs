
'#########################################################################
'  Script name:		PP_SaveAsJPG.vbs
'  Created on:		4/26/2010
'  Author:			Dennis Hemken
'  Purpose:			Opens an existing Microsoft PowerPoint presentation
'					and then saves the slides as JPG's.
'#########################################################################

Dim AppPowerPoint 
Dim OpenPresentation
Dim lngSlideCount
Const ppSaveAsJPG = 17

Set AppPowerPoint = CreateObject("PowerPoint.Application")

AppPowerPoint.Visible = True

Set OpenPresentation = AppPowerPoint.Presentations.Open("C:\Concepts\Management.ppt")
	
OpenPresentation.SaveAs "C:\Concepts\JPG\Management", ppSaveAsJPG

OpenPresentation.Close
Set OpenPresentation = Nothing

AppPowerPoint.Quit
Set AppPowerPoint = Nothing