
'#########################################################################
'  Script name:		PP_SaveAsJPG.vbs
'  Created on:		06/29/2011
'  Author:			Dennis Hemken
'  Purpose:			Opens an existing Microsoft PowerPoint presentation
'					and then saves the slides as GIF's.
'#########################################################################


Dim AppPowerPoint 
Dim OpenPresentation
Dim lngSlideCount
Const ppSaveAsGIF = 16

Set AppPowerPoint = CreateObject("PowerPoint.Application")

AppPowerPoint.Visible = True

Set OpenPresentation = AppPowerPoint.Presentations.Open("C:\Concepts\Management.ppt")
	
OpenPresentation.SaveAs "C:\Concepts\GIF\Management", ppSaveAsGIF

OpenPresentation.Close
Set OpenPresentation = Nothing

AppPowerPoint.Quit
Set AppPowerPoint = Nothing