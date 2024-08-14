Option Explicit 

Dim strComputer,strServiceName
strComputer = "." ' Local Computer
strServiceName = "wlansvc" ' 

IF isServiceRunning(strComputer,strServiceName) THEN
	Call SethostedNetwork
ELSE
	wscript.echo "The Wlan service is not running.Please turn on wireless and run this script on laptop."
END IF
	
Function SethostedNetwork
		Dim objshell,AdapterInfo,AdapterValue
		Set objshell = WScript.CreateObject("wscript.shell")
		AdapterInfo = objshell.Exec("netsh wlan show drivers").StdOut.ReadAll()
		AdapterValue = InStr(AdapterInfo,"Hosted network supported  : Yes")	
		If AdapterValue > 0 Then 
			Dim  colNamedArguments,Action
			Set colNamedArguments = WScript.Arguments.Named
			Action = colNamedArguments.Item("A")
			If colNamedArguments.Exists("A") Then 
				If InStr(UCase(Action),UCase("Stop")) > 0 Then 
					WScript.Echo objshell.Exec("Netsh wlan stop hostednetwork").StdOut.ReadAll()
				Else 	
					WScript.Echo "Invalid argument."
				End If 
			Else 
				Dim FSO,CurrentDirectory,Txtpath,txtfile
				Set FSO = CreateObject("Scripting.FileSystemObject")
				CurrentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
				Txtpath= CurrentDirectory & "\Hostednetwork.txt"
				If FSO.FileExists(Txtpath) Then 
					Set txtfile = FSO.OpenTextFile(Txtpath, 2 , true)
				Else
					Set txtfile = FSO.CreateTextFile(Txtpath)
				End If 
				Dim HostName,Password,Command
				HostName = "Mywifi"
				Password = GetRandom(8)
				Command = "netsh wlan set hostednetwork mode=allow ssid=  """ & HostName & """ key= """ & Password &""" "
				objshell.Exec(Command)
				objshell.Exec("netsh wlan start hostednetwork")
				Dim objShellEx
				Set objShellEx = objshell.Exec("netsh wlan show hostednetwork")
				Do While objShellex.Status = 0 
					WScript.Sleep 1 
				Loop 
				Dim strText,ModeValue,StatusValue
				strText = objShellEx.StdOut.ReadAll()
				ModeValue = InStr(strText,"Mode                   : Allowed")
				StatusValue = InStr(strText,"Status                 : Started")
				If Modevalue > 0 And StatusValue > 0 Then 
					WScript.Echo "Set hosted NetWork successfully. Network name is "& "'" & HostName & "'" & " and password is " & "'" & Password & "'"
					txtfile.WriteLine "NetWork Name :" & HostName
					txtfile.WriteLine "Password     :" & Password
				Else
					WScript.Echo "failed to set hosted NetWork."
				End If 
				txtfile.Close()
				Set objShellEx = Nothing
			End If  	
		Else 
			WScript.Echo "Your network adapter does not support hosted network."
		End If 
End Function 

' Function to check if a service is running on a given computer
FUNCTION isServiceRunning(strComputer,strServiceName)
	DIM objWMIService, strWMIQuery
	strWMIQuery = "Select * from Win32_Service Where Name = '" & strServiceName & "' and state='Running'"
	SET objWMIService = GETOBJECT("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	IF objWMIService.ExecQuery(strWMIQuery).Count > 0 THEN
		isServiceRunning = TRUE
	ELSE
		isServiceRunning = FALSE
	END IF
END Function

Function GetRandom(Count)
    Randomize
    Dim i 
    For i = 1 To Count
        If (Int((1 - 0 + 1) * Rnd + 0)) Then
            GetRandom = GetRandom & Chr(Int((90 - 65 + 1) * Rnd + 65))
        Else
            GetRandom = GetRandom & Chr(Int((57 - 48 + 1) * Rnd + 48))
        End If
    Next
End Function