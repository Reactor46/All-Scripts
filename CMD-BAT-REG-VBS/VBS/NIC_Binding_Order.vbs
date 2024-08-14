On Error Resume Next

const HKEY_CURRENT_USER = &H80000001
const HKEY_LOCAL_MACHINE = &H80000002

strComputer = Wscript.Arguments.Item(0)

Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter WHERE NetConnectionStatus = 2")
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

For Each objItem In colItems
	WScript.Echo ""
	WScript.Echo "HostName : " & UCase(strComputer)
	Wscript.Echo ""

	strKeyPath = "SYSTEM\Currentcontrolset\Services\TCPIP\Linkage"
	strValueName = "Bind"
	oReg.GetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,arrValues
 
	For Each strValue In arrValues
    		strNewValue = Replace(strValue,"\Device\","")   
		strNewKeyPath = "SYSTEM\Currentcontrolset\Control\Network\{4D36E972-E325-11CE-BFC1-08002be10318}\" & strNewValue & "\Connection"
		strNewValueName = "Name"
		oReg.GetStringValue HKEY_LOCAL_MACHINE,strNewKeyPath,strNewValueName,strNicName
		WScript.Echo "NIC      : " & strNicName

		strKeyPathIP = "SYSTEM\Currentcontrolset\Services\TCPIP\Parameters\Interfaces\" & strNewValue
		strValueNameIP = "IPAddress"
		oReg.GetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPathIP,strValueNameIP,arrValuesIP

		For Each IP in arrValuesIP
			If IP = "" Then
				Wscript.Echo "IP       : NOT ASSIGNED"
			Else
				Wscript.Echo "IP       : " & IP
			End If
		Next
		Wscript.Echo ""

 	Next

Next
