Dim objWMIService, objComputer, colComputer, strComputer, wshArguments, ObjComp, result

set wshArguments = wscript.arguments
set objComp = getobject(wsharguments(0))
result = mid(objComp.name,4)

If result <> "" Then
	strComputer = result
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" _ 
		& strComputer & "\root\cimv2") 
	Set colComputer = objWMIService.ExecQuery _
		("Select * from Win32_ComputerSystem") 
	
	For Each objComputer in colComputer
	Wscript.Echo objComputer.UserName & " is logged on"
	Next
End If 