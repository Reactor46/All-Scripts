'=*=*=*=*=*=*=*=*=*=*=*=
' Coded By Assaf Miron 
' Date : 01/12/08
'=*=*=*=*=*=*=*=*=*=*=*=

On Error Resume Next
'======================================================================================================================
' Consts
'======================================================================================================================
' File Consts
Const ForReading = 1
Const LogFile = "C:\ShutdownComputers-Log.txt"
'Sched Consts
Const Sunday = 64
Const  Monday = 1
Const  Tuesday = 2
Const  Wednesday = 4
Const  Thursday = 8
Const  Friday = 16
Const  Saturday = 32
Const SchedTime = 184000 				' 19:30:00
Const TimeZone = "-480" 					' Represents your time zone offset from Greenwich Mean Time in Minutes (-480 is USA)
Const RUN_REPEATEDLY = True 	' Indicates whether or not the scheduled job should run repeatedly
'======================================================================================================================
' Set Objects
'======================================================================================================================
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject") 

Sub SchedJob(strComputer)
'======================================================================================================================
' This sub creates a schedule job on a computer that it recives.
' It uses the SchedTime const to define the time of running the comand, the TimeZone const to define the time zone the script is running in, whether to run the job repeatedly or not (from the consts also).
' It then creates the new job with all the data and run it under system privleges and the remote computer name - no user interfirence.
'======================================================================================================================
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objNewJob = objWMIService.Get("Win32_ScheduledJob")

Days = Sunday	' MultiDays are enterd with a OR 
If InStr(Days,"OR") = true Then
	RUN_REPEATEDLY = True
End If

' Create a task to Shutdown the Computer with a 30 Seconds timeout
errJobCreated = objNewJob.Create ("cmd /c shutdown -s -t 30", "********" & SchedTime & ".000000" & TimeZone, RUN_REPEATEDLY , Days, , , JobID)

If Err.Number = 0 Then
 wscript.echo "New Job ID: " & JobID
 wscript.echo "The scheduled task was successfully created."
Else
	wscript.echo "An error occurred: " & errJobCreated ' & "("
	Select Case errJobCreated
		Case 1 : wscript.echo "The request is not supported."
		Case 2 : wscript.echo "The user did not have the necessary access."
		Case 8 : wscript.echo "Interactive Process."
		Case 9 : wscript.echo "The directory path to the service executable file was not found."
		Case 21 : wscript.echo "Invalid parameters have been passed to the service."
		Case 22: wscript.echo "The account under which this service is to run either is invalid or lacks the permissions to run the service."
	End Select
	' wscript.echo ")"
End If

End Sub

' Opening File using a Common Dialog
Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "Text Files|*.txt|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
Else
    FileLoc = objDialog.FileName
End If

' Check if the Log File Exists
If objFSO.FileExists (LogFile) THEN
	set objFile = objFSO.GetFile (LogFile)
	' Delete the log file
	objFile.Delete
end If 
' Create the log file
Set objFile = objFSO.CreateTextFile (LogFile, True)

objFile.WriteLine "Log Started : " & Now
objFile.WriteLine

'Get List Of Computers
set objReadFile = objFSO.OpenTextFile(FileLoc, ForReading)
Do Until objReadFile.AtEndOfStream
    strComputer = objReadFile.Readline
    SchedJob strComputer
    Next
Loop

' Cleanup - Close Files
objReadFile.Close
objFile.WriteLine
objFile.WriteLine "Log Ended : " & Now
objFile.Close
wscript.echo "Done!"