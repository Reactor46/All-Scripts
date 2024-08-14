'#########################################################################
'  Script name:		WMI_KillProcess.vbs
'  Created on:		10/05/2010
'  Author:			Dennis Hemken
'  Purpose:			This function kills a process by name,
'                   which is running on a special pc in the network.
'#########################################################################
Dim strComputer

    strComputer = "."
    
	fct_KillProcess "acrord32", strComputer
	' or
	' strComputer = "192.168.2.13"
	' fct_KillProcess "outlook", strComputer

Public Function fct_KillProcess(strProcessName, strComputer)
  
	Dim objWMI
	Dim colServices
	Dim objService
	Dim strServicename
	Dim ret

	Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colServices = objWMI.InstancesOf("win32_process")

	For Each objService In colServices
		strServicename = LCase(Trim(CStr(objService.Name) & ""))
		If InStr(1, strServicename, LCase(strProcessName), vbTextCompare) > 0 Then
			ret = objService.Terminate
		End If
	Next
	Set colServices = Nothing
	Set objWMI = Nothing
End Function