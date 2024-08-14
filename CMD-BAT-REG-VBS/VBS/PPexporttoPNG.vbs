
'#########################################################################
'  Script name:		PPexporttoPNG.vbs
'  Created on:		07/18/2011
'  Author:			Dennis Hemken
'  Purpose:			Opens an existing Microsoft PowerPoint presentation
'					and then saves the slides as PNG's.
'#########################################################################


Dim AppPowerPoint 
Dim OpenPresentation
Const ppSaveAsPNG = 18

Set AppPowerPoint = CreateObject("PowerPoint.Application")

AppPowerPoint.Visible = True

Set OpenPresentation = AppPowerPoint.Presentations.Open("C:\Concepts\Management.ppt")
	
OpenPresentation.SaveAs "C:\Concepts\GIF\Management", ppSaveAsPNG

OpenPresentation.Close
Set OpenPresentation = Nothing

AppPowerPoint.Quit
Set AppPowerPoint = Nothing