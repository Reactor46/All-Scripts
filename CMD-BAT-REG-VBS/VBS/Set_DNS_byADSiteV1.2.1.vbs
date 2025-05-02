'by k8l0

strLogFile = "."

Const ForAppending = 8
Const ForReading = 1
Set objADSysInfo = CreateObject("ADSystemInfo")
Set wshNetwork = WScript.CreateObject( "WScript.Network" )
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")


strADSite = objADSysInfo.SiteName
strHostName = wshNetwork.ComputerName
strDomain = wshNetwork.UserDomain 
strDay = Day(Date())
strMonth = Month(Date())
strYear = Year(Date())
If Len(strDay) = 1 Then strDay = "0" & strDay
If Len(strMonth) = 1 Then strMonth = "0" & strMonth
strDate = strDay & "/" & strMonth & "/" & strYear

Set objLogFile = objFSO.OpenTextFile(strLogFile & "\" & strDomain & "-" & strHostName & ".txt", ForAppending, True)
Set objInpFile = objFSO.OpenTextFile(".\dnsdata.csv", ForReading)

'Before Log
objLogFile.Write(strDate & "," & Time() & ",Before Change")
objLogFile.Write("," & strDomain)
objLogFile.Write("," & strHostName)
objLogFile.Write("," & strADSite)

Set colNicConfigsBefore = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
For Each objNicConfig In colNicConfigsBefore
	strDNSSuffixSO = ""
	strDNSServerSO = ""
	strDNSSuffixSO = ""
	If Not IsNull(objNicConfig.DNSDomainSuffixSearchOrder) Then
		For Each strDNSSuffix In objNicConfig.DNSDomainSuffixSearchOrder
			If strDNSSuffixSO = "" Then
				strDNSSuffixSO = strDNSSuffix
			Else	
				strDNSSuffixSO = strDNSSuffixSO & " / " & strDNSSuffix
			End If
		Next
	End If
	strDNSServerSO = ""
	If Not IsNull(objNicConfig.DNSServerSearchOrder) Then
		For Each strDNSServer In objNicConfig.DNSServerSearchOrder
			If strDNSServerSO = "" Then
				strDNSServerSO = strDNSServer
			Else
				strDNSServerSO = strDNSServerSO & "," & strDNSServer
			End If
		Next
	End If
	strDomainDNSRegistrationEnabled = objNicConfig.DomainDNSRegistrationEnabled
	strFullDNSRegistrationEnabled = objNicConfig.FullDNSRegistrationEnabled
	strDNSServerSearchOrder = strDNSServerSO
	objLogFile.Write("," & strDNSServerSearchOrder)
Next
objLogFile.WriteLine("")

Set colNicConfigs = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
Do Until objInpFile.AtEndOfStream
	arrRecord = split(objInpFile.Readline, ",")
	strInpADSite = arrRecord(0)
	strInpDomain = arrRecord(1)
    strInpDNS1 = arrRecord(2)
	strInpDNS2 = arrRecord(3)
    If strADSite = strInpADSite and strDomain = strInpDomain Then
		arrNewDNSServerSearchOrder = Split(strInpDNS1 & "," & strInpDNS2,",")
		For Each objNicConfig In colNicConfigs
			If Not IsNull(objNicConfig.DNSServerSearchOrder) Then
				intSetDNSServers = objNicConfig.SetDNSServerSearchOrder(arrNewDNSServerSearchOrder)
			End If
		Next
	End If
Loop

'After Log

objLogFile.Write(strDate & "," & Time() & ",After Change")
objLogFile.Write("," & strDomain)
objLogFile.Write("," & strHostName)
objLogFile.Write("," & strADSite)

Set colNicConfigsAfter = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
For Each objNicConfig In colNicConfigsAfter
	strDNSSuffixSO = ""
	strDNSServerSO = ""
	strDNSSuffixSO = ""
	If Not IsNull(objNicConfig.DNSDomainSuffixSearchOrder) Then
		For Each strDNSSuffix In objNicConfig.DNSDomainSuffixSearchOrder
			If strDNSSuffixSO = "" Then
				strDNSSuffixSO = strDNSSuffix
			Else	
				strDNSSuffixSO = strDNSSuffixSO & " / " & strDNSSuffix
			End If
		Next
	End If
	strDNSServerSO = ""
	If Not IsNull(objNicConfig.DNSServerSearchOrder) Then
		For Each strDNSServer In objNicConfig.DNSServerSearchOrder
			If strDNSServerSO = "" Then
				strDNSServerSO = strDNSServer
			Else
				strDNSServerSO = strDNSServerSO & "," & strDNSServer
			End If
		Next
	End If
	strDomainDNSRegistrationEnabled = objNicConfig.DomainDNSRegistrationEnabled
	strFullDNSRegistrationEnabled = objNicConfig.FullDNSRegistrationEnabled
	strDNSServerSearchOrder = strDNSServerSO
	objLogFile.Write("," & strDNSServerSearchOrder)
Next
objLogFile.WriteLine("")

objInpFile.close
objLogFile.close