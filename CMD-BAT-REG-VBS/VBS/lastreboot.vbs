'********************************************************************
'* File: LastReboot.vbs
'* http://gallery.technet.microsoft.com/ScriptCenter/82588289-4e07-455e-8322-c635cc719f00/
'*
'* Author: Manoj Nair | Created on 09/10/2009
'* Version 1.0
'*
'* Main Function: Displays the last reboot time of a computer
'*
'* Modified to incude test for Online Status
'* Modified so output is formatted for .csv file
'* Modified to include IP Address in output
'* Modified by Michael Martin | 6/6/2014
'*
'********************************************************************

On Error Resume Next
Const ForReading = 1
Const HKEY_LOCAL_MACHINE = &H80000002


Set objFSO = CreateObject("Scripting.FileSystemObject")


    ' =====================================================================
     'Gets the script to run against each of the computers listed 
     'in the text file path for which should be specified in the syntax below
    ' =====================================================================
Set objTextFile = objFSO.OpenTextFile("uptime_sources.txt", ForReading)
Set outfile = objFSO.CreateTextFile("Report.csv")
Outfile.Writeline "Machine Name, IP Address, Online Status, User Name, Last Startup Time, Uptime"
Do Until objTextFile.AtEndOfStream 
    strComputer = objTextFile.Readline
	OnlineStatus =  Reachable(strComputer)
	strAddress = IPV4(strComputer)
	If OnlineStatus = True then
		' ===============================================================================
		' Code to get the Last Boot Time using LastBootupTime from Win32_Operating System
		' ===============================================================================
		' Code to get the Last Logged On User from StdRegProv
		' ===============================================================================
		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
		Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
		Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
		strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"
		strValueName = "LastLoggedOnUser"
		objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strUsername
		myArray = split(strUsername, "\")
		strUsername = myArray(1)
		For Each objOS in colOperatingSystems
			dtmBootup = objOS.LastBootUpTime
			dtmLastBootupTime = WMIDateStringToDate(dtmBootup)
			'OutFile.WriteLine "=========================================="
			dtmSystemUptime = DateDiff("h", dtmLastBootUpTime, Now)
			OutFile.WriteLine strComputer & "," & strAddress & ", " & "Online" & "," & strUsername & ", " & dtmLastBootupTime & ","  & dtmSystemUptime
		Next
	Else 
		OutFile.WriteLine strComputer & ", " & "," & "Offline"
	End If

objTextFile.Close
' =====================================================================
' End of Main Body
' =====================================================================
Loop
 ' ===============================================================================
 ' Displaying to the user that the script execution is completed
 ' ===============================================================================
MsgBox "Script Execution Completed. The Report is saved as Report.csv in the current directory"


 ' ===============================================================================
 ' Functions
 ' ===============================================================================

 
 ' ===============================================================================
 ' Function to check for online status
 ' ===============================================================================
 Function Reachable(strComputer) 
   On Error Resume Next 
   Dim wmiQuery, objWMIService, objPing, objStatus 
   wmiQuery = "Select * From Win32_PingStatus Where Address = '" & strComputer & "'" 
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2") 
   Set objPing = objWMIService.ExecQuery(wmiQuery) 
   For Each objStatus In objPing
     If IsNull(objStatus.StatusCode) Or objStatus.Statuscode<>0 Then 
        Reachable = False 'if computer is unreachable, return false 
      Else 
         Reachable = True 'if computer is reachable, return true 
      End If 
   Next  
End Function

 ' ===============================================================================
 ' Function to check for IP Address
 ' ===============================================================================
Function IPV4(strComputer)
	Set objWMIService = GetObject( _ 
		"winmgmts:\\" & strComputer & "\root\cimv2")
	Set IPConfigSet = objWMIService.ExecQuery _
		("Select IPAddress from Win32_NetworkAdapterConfiguration ")
	For Each IPConfig in IPConfigSet
			If Not IsNull(IPConfig.IPAddress) Then 	
                IPV4 =  IPConfig.IPAddress(i)
			End If	
	Next
End Function
 
 '================================================================================
 ' Function to convert UNC time to readable format
 ' ===============================================================================

 Function WMIDateStringToDate(dtmBootup)
    WMIDateStringToDate = CDate(Mid(dtmBootup, 5, 2) & "/" & _
         Mid(dtmBootup, 7, 2) & "/" & Left(dtmBootup, 4) _
         & " " & Mid (dtmBootup, 9, 2) & ":" & _
         Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup, _
         13, 2))
End Function
