strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_BIOS",,48)
	For Each objItem in colItems 
		if objItem.SerialNumber = "" then
			Msgbox "Can Not Determine Serial Number"
		else
			msgbox objItem.SerialNumber
		end if
	Next