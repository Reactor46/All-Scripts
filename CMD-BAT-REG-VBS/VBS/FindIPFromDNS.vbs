'=*=*=*=*=*=*=*=*=*=*=*=*=
' Author : Assaf Miron 
' Http://assaf.miron.googlepages.com
' Date : 24/11/2009
' FindIPFromDNS.vbs
' Description : Finds The IP Address of a List of Servers from the DNS A Records
' The Script Creates a Hash Table (Dictionary) of all A Records from the Container 'Mydomain.com'
' The Script Writes to a Log file all the IP Addresses of the Servers found from the DNS
'=*=*=*=*=*=*=*=*=*=*=*=*=
Option Explicit
On Error Resume Next

Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2

Dim colItems
Dim objWMIService, objDictionary, objItem
Dim objFSO, objSRVNames, objLogFile
Dim strDNSServer, strOwner, strServerName, strServersFile
Dim strLogFile : strLogFile = "C:\ServersIP.txt"
Dim tStartTime, tEndTime, diff

' Record the Start Time
tStartTime = Now

' Set the DNS Server Name
strDNSServer = "DNSSrv"
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strDNSServer & _
        "\root\MicrosoftDNS")

' Set a Hash Table        
Set objDictionary = CreateObject("Scripting.Dictionary")

' Set a File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' List Host Address DNS Records
Set colItems = objWMIService.ExecQuery("Select * from MicrosoftDNS_AType WHERE ContainerName = 'Mydomain.com'")

' Create a Hash Table of DNS Records (A Records)
For Each objItem in colItems
	strOwner = Mid(objItem.OwnerName, 1, InStr(objItem.OwnerName,".Mydomain.com"))
	If strOwner <> "" Then ' Check that the Host Name is not Empty
		strOwner = Left(strOwner,len(strOwner)-1) ' Trim the End '.'
		If Not objDictionary.Exists(strOwner) Then ' Check that Hash table doesnt have the Host Name
			objDictionary.Add strOwner , objItem.IPAddress ' Add the Host name and IP
		End If
	End if
Next

' Get the Servers List File Name from Arguments
strServersFile = WScript.Arguments(0)

' Open the Servers File
Set objSRVNames = objFSO.OpenTextFile(strServersFile, ForReading)

' Check if the log File Exists
If objFSO.FileExists(strLogFile) Then
	' Overwrite any Existing Values - Keep the Log Updated
	Set objLogFile = objFSO.OpenTextFile(strLogFile,ForWriting)
Else
	' Create a New Log File
	Set objLogFile = objFSO.CreateTextFile(strLogFile)
End If

Do Until objSRVNames.AtEndOfStream
	strServerName = objSRVNames.ReadLine ' Get the Server Name
	' Check if Server Exists in the Hash Table
	If objDictionary.Exists(strServerName) Then
		' Write the Server's IP to the Log file
		 objLogFile.WriteLine objDictionary.Item(strServerName)
	End if
Loop

' Close Files
objSRVNames.Close
objLogFile.Close

' Clean Up
Set objLogFile = Nothing
Set objSRVNames = Nothing
Set objFSO = Nothing
' Record the End Time
tEndTime = Now

' Calculate the Time Differenecs
diff = DateDiff("s",tStartTime, tEndTime)
WScript.Echo "Script Done in " & diff & " Secodns!"