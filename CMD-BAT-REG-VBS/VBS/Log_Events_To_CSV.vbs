'''''' Logon Track event to CSV
' Scott Morris - 2012-04-04
' logon.vbs - a logon script that logs who logs into what machine
' Info logged: the date, the time in 24-hour format, the person's username,
'              the person's full name, the computer name that they logged into,
'              the IP address of that computer, and the MAC address of that computer

Option Explicit
On Error Resume Next ' Blanket error-handling statement

' Set up the target file
const strFile="\\\\turkish\\shared\\sessions\\computerusers.csv"

Dim WshNetwork, outputStr, strIP, strMAC, strQuery, objWMIService, colItems, _
objItem, objFSO, objTextFile, strFullName

'Create the network object
Set WshNetwork = WScript.CreateObject("WScript.Network")

'Connect to the WMI service for the machine being logged into
Set objWMIService = GetObject( "winmgmts://" & WshNetwork.ComputerName & "/root/CIMV2" )

'Set up the query to pull the WMI info
strQuery = "SELECT IPAddress, MACAddress FROM Win32_NetworkAdapterConfiguration WHERE MACAddress is not null"

'Run the query to pull the info
Set colItems = objWMIService.ExecQuery( strQuery, "WQL", 48 )

' Grab the IP and MAC addresses
For Each objItem In colItems
    If IsArray( objItem.IPAddress ) Then
    	strIP = objItem.IPAddress(0)
    	strMAC = objItem.MACAddress(0)
    End If
Next

' Grab the user's account information
Set colItems = objWMIService.ExecQuery("Select FullName from Win32_UserAccount where Name='" & _
WshNetwork.Username & "'",,48)

' Take note of the user's full name
For Each objItem in colItems
	strFullName = objItem.FullName
Next

' Format the output to be CSV-friendly
outputStr = Date() & "," & FormatDateTime(Time, 4) & "," & WshNetwork.UserName & "," & _
strFullName & "," & WshNetwork.ComputerName & "," & strIP & "," & strMAC

' Write out the information to the target file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Const ForAppending = 8
Set objTextFile = objFSO.OpenTextFile (strFile, ForAppending, True)
objTextFile.WriteLine(outputStr)
objTextFile.Close

''''' End Script