Option Explicit
Dim sStartAddress, iStartAddress, sEndAddress, iEndAddress, sCurrentAddress, iCurrentAddress
Dim oLocalWMIService, iIPWarningThreshold, iResponse, bTrace, bExportToCSV, sCSVPath, oFs, oOutput, IE, x, ICMPResponseTimeout

sStartAddress = "192.168.0.1"
sEndAddress = "192.168.0.254"
bTrace = True
bExportToCSV = False
iIPWarningThreshold = 255
ICMPResponseTimeout = 10

If bTrace = False And bExportToCSV = False Then
	Wscript.echo "IE and CSV output are turned off so the output doesn't have anywhere to go!" & vbcrlf & vbcrlf &_
								"Take a look at changing either the bTrace or bExportToCSV variables at the top of the script to resolve this..." & vbcrlf & vbcrlf &_
								"Goodbye."
	Wscript.quit
End If

sCSVPath = Left(Wscript.ScriptName, Len(Wscript.ScriptName)-4) & " on " & Replace(Date,"/",".") & " - " & sStartAddress & " - " & sEndAddress & ".csv"

Trace "<b>Validating the IP Address passed to the script</b>"
Trace "--"
ValidateIP(sStartAddress)
Trace "--"
ValidateIP(sEndAddress)
Trace "--"

iStartAddress = ConvertIPtoInteger(sStartAddress)
iEndAddress = ConvertIPtoInteger(sEndAddress)

If iStartAddress > iEndAddress Then
		Trace "The start address cannot be larger than the end address. Please correct this and try again."
		Wscript.quit
End If

Trace iEndAddress - iStartAddress & " IP addresses to go though."
If (iEndAddress - iStartAddress) > iIPWarningThreshold Then
	Trace "Prompt to continue sent as " & iEndAddress - iStartAddress & " is higher then the warning threshold of " & iIPWarningThreshold
	iResponse = MsgBox("There are about " & iEndAddress - iStartAddress & " addresses in the IP range you provided." & vbcrlf & vbcrlf &_
	"Start Address: " & sStartAddress & vbcrlf &_
	"End Address: " & sEndAddress & vbcrlf & vbcrlf &_
	"This may take a some time. Are you sure about this?!?", vbYesNo, "Ping Sweeper")
	If iResponse = vbNo Then Wscript.quit
End If


If bExportToCSV Then
	Set oFs = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	Set oOutput = oFs.CreateTextFile(sCSVPath)
	If Err Then
		If bTrace Then 
			bExportToCSV = False
			Trace "--"
			Trace "Unable to create the output file: " & sCSVPath
			Trace "Check your permissions to the current directory and that the file is not locked."
			Trace "Trace is enabled so the script will continue but results will not be pushed into a CSV file."
			Trace "--"
		Else
			Wscript.echo "Unable to create the output file: " & sCSVPath & "." & vbcrlf &_
			"Check your permissions to the current directory and that the file is not locked."
			Wscript.quit
		End If
	Else
		oOutput.WriteLine Chr(34) & "IP Address" & Chr(34) & "," &_
			Chr(34) & "Ping Response" & Chr(34) & "," &_
			Chr(34) & "Host Name" & Chr(34)
	End If
	On Error GoTo 0
End If

Trace "Now going through the IP addresses<br>--<br>"

Set oLocalWMIService = GetObject("winmgmts:\\.\root\cimv2")
x = iStartAddress
While x <= iEndAddress
	iCurrentAddress = x
	sCurrentAddress = ConvertIntegertoIP(iCurrentAddress)
	Ping(sCurrentAddress)
	x = x+1
Wend

Trace "--"

'--

Sub Ping(sIP)
	Dim cPingStatus, oPing, sHostName
	Set cPingStatus = oLocalWMIService.ExecQuery ("Select * from Win32_PingStatus Where Address = '" & sIP &_
	"' AND resolveAddressNames = true AND  statusCode = 0 AND timeout = " & ICMPResponseTimeout & "")
	For Each oPing in cPingStatus
		sHostName = oPing.ProtocolAddressResolved
		If sHostName = sIP Then
			Trace sIP & " (Unresolvable host) is responding"
			If bExportToCSV Then oOutput.WriteLine Chr(34) & sIP & Chr(34) & "," & Chr(34) & "True" & Chr(34) & "," & Chr(34) & "Unresolvable host" & Chr(34)
		Else
			Trace sIP & " (" & sHostName & ") is responding"
			If bExportToCSV Then oOutput.WriteLine Chr(34) & sIP & Chr(34) & "," & Chr(34) & "True" & Chr(34) & "," & Chr(34) & sHostName & Chr(34)
		End If
	Next
End Sub

Function ConvertIPtoInteger(sIP)
	Dim aIpOctets
	aIpOctets = Split(sIP, ".")
	ConvertIPtoInteger = (aIpOctets(0) * 16777216) + (aIpOctets(1) * 65536) + (aIpOctets(2) * 256) + (aIpOctets(3))
End Function

Function ConvertIntegertoIP(iNumber)
	Dim i, iTemp
	For i = 1 To 4
		iTemp = Int(iNumber/256^(4 -i))
		iNumber = iNumber-(iTemp*256^(4-i))
			If i = 1 Then
			ConvertIntegertoIP = iTemp
		Else
			ConvertIntegertoIP = ConvertIntegertoIP & "." & iTemp
		End If
	Next
End Function

Sub ValidateIP(sIP)
	Dim i, aIpOctets
	aIpOctets = Split(sIP, ".")
	Trace "Checking the IP address " & sIP & " is valid."
	If (UBound(aIpOctets) <> 3) Then
		Trace sIP & " doesn't appear to have the correct amount of octets."
		Wscript.quit
	Else
		Trace sIP & " appears to have the correct amount of octets."
	End If
	Trace "Checking each octet in each IP is numeric and below 256"
	For i = 0 To 3
		If (IsNumeric(aIpOctets(i))=False) Or (aIpOctets(i) > 256) Then
			Wscript.echo "Octect " & i=1 & " doesn't appear to be valid in the address: " & sIP
			Wscript.quit
		End If
	Next
	Trace sIP & " appears to have a valid number in each octet."
End Sub

Sub Trace(sMsg)
	If bTrace = True Then
	If Not IsObject(IE) Then
		Set IE = CreateObject("InternetExplorer.Application")
		IE.navigate "about:blank"
		IE.ToolBar = False
		IE.AddressBar = False
		IE.Width = 850
		IE.Height = 600
		IE.Visible = True
		IE.menubar = False
		IE.StatusBar = False
		IE.Document.writeln "<Html><Head>" & vbcrlf &_
			"<Title>:: VB Script Status :: </Title><Style>" & vbcrlf &_
			"Body {Background-Color:silver;font-size:10pt;font-family: verdana, sans-serif;color:navy;font-weight:normal;" &_
			"scrollbar-face-color:#C0C0C0;scrollbar-highlight-color:#FFFFFF;scrollbar-shadow-color:#C0C0C0;" &_
			"scrollbar-3dlight-color:#808080;scrollbar-arrow-color:#003752;" &_
			"scrollbar-track-color:#C9C7C7;scrollbar-darkshadow-color:#98AAB1;}" & vbcrlf &_
			"#table1 {font-size:10pt;}" & vbcrlf &_
			"</Style></Head><BODY>"
	End If
	IE.Document.writeln sMsg & "<br>"
	End If
End Sub