strComputer = "."
	On Error Resume Next
	Set wshNetwork = WScript.CreateObject( "WScript.Network" )
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_BIOS",,48)
		For Each objItem in colItems 
			if objItem.SerialNumber = "" then
				Serial = "Can Not Determine Serial Number"
			else
				Serial = objItem.SerialNumber
			end if
		Next
	Set objGroup = GetObject("WinNT://" & strComputer &  "/Administrators,group")
		If objGroup.PropertyCount > 0 Then
			For Each mem In objGroup.Members
				PassAge = ""
				intPasswordAge = ""
				IsDisabled = ""
				user = Right(mem.adsPath,Len(mem.adsPath) - 8)
					Set objUser = GetObject("WinNT://" & strComputer & "/" & mem.name & ", user")
						IF objUser.AccountDisabled = True then
							IsDisabled = " - Disabled"
						End If
						IF objUser.AccountDisabled = False then
							IsDisabled = " - Enabled"
						End If
				intPasswordAge = objUser.PasswordAge
				intPasswordAge = intPasswordAge * -1 
				dtmChangeDate = DateAdd("s", intPasswordAge, Now)
					PassAge = " - " & FormatNumber(Now - dtmChangeDate, 0)
				If instr(9, mem.adsPath, "DOMAIN/") = 9 then
					PassAge = ""
					IsDisabled = ""
				End if
			If results = "" then
				Results = "Serial Number: " & Serial & VbCrlf & "Computer Name: " & wshNetwork.ComputerName & VbCrlf & VbCrlf &_
				mem.name & PassAge & IsDisabled & VbCrlf
			Else
				Results = results & mem.name & PassAge & IsDisabled & VbCrlf
			End If
			Next
		End if
msgbox Results