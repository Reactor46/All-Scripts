'Set DNS By k8l0

If WScript.Arguments.Count = 3 Then
	strInputFile = WScript.Arguments.Item(0)
	strOutputFile = WScript.Arguments.Item(1)
	strNewDNS = WScript.Arguments.Item(2)
Else
	wscript.echo "Sintaxe: cscript SetDNSv2.vbs inputfile.txt outputfile.txt 10.1.98.64,10.1.98.36,10.1.18.24"
	wscript.quit
end if	

On error resume next

Const ForReading = 1
Const ForAppending = 8
 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objTextFileIn = objFSO.OpenTextFile(strInputFile, ForReading)
Set objTextFileOut = objFSO.OpenTextFile(strOutputFile, ForAppending, True)

wscript.echo "Host		Adapter		Return Status"
wscript.echo "----		-------		-------------"
objTextFileOut.WriteLine("Inputed,Host,Adapter,Return Status")

Do Until objTextFileIn.AtEndOfStream 
    strComputer = Trim(objTextFileIn.Readline)
	
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colNicConfigs = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
	
	For Each objNicConfig In colNicConfigs
		If Not IsNull(objNicConfig.DNSServerSearchOrder) Then
			strReturn = ""
			arrNewDNSServerSearchOrder = Split(strNewDNS,",")
			intSetDNSServers = objNicConfig.SetDNSServerSearchOrder(arrNewDNSServerSearchOrder)
			If intSetDNSServers = 0 Then
				strReturn = """" & "Replaced DNS server search order list to " & strNewDNS & "." & """"
			Else
				strReturn = "Unable to replace DNS server search order list."
			End If
		Else
			strReturn = "DNS server search order is null. Nothing changed!"
		End If
		
		strDNSHostName = objNicConfig.DNSHostName
		strIndex = objNicConfig.Index
		strDescription = objNicConfig.Description
		strAdapter = "Network Adapter " & strIndex & " - " & strDescription
		wscript.echo strDNSHostName & VBTab & strAdapter & VBTab & strReturn
		objTextFileOut.WriteLine(strComputer & "," & strDNSHostName & "," & strAdapter & "," & strReturn)
		Next
Loop 

objTextFileIn.close
objTextFileOut.close

wscript.echo "Finished!!!"