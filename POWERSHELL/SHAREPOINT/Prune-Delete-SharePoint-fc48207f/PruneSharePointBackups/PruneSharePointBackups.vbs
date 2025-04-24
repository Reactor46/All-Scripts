'==================================================================================================================================================
' Name:				PruneSharePoint.vbs
' Description:		Prunes SharePoint Backups older than the specified number of days (30 days by default) 
'
' Author:			Matthew Yarlett
' Creation date:	12/04/10
'
' Notes:
' ------
' 
' Modifications:
' --------------
' Date:     Author: Version:  Description:
' --------  ------- --------  -------------------------------------------------------------------
' 12/04/10  MY        	1.0   Created
'==================================================================================================================================================

'==================================================================================================================================================
'***********' Section: Script Parameters / Arguments '***********'
'================================================================================================================================================== 
'Comment out to allow undefined variables
Option Explicit
'Comment out to turn error handling off
'On Error Resume Next
'Set the timeout for script (in seconds)
'WScript.timeout = 5000
'Check environment variables
Dim arg
Dim enableLog
enableLog = False
Dim sTOCfile
sTOCfile = Empty
Dim sOutputPath
Dim iDeleteBackupsOlderThanXDays
iDeleteBackupsOlderThanXDays = 30

Dim oArguments
Set oArguments = wscript.arguments
Dim a
Dim abort
If oArguments.count > 2 Then
    For a = 0 To (oArguments.count -1)
        arg = Wscript.arguments(a)
        If lcase(arg) = "-p" Then
       		'Directory to search
       		sTOCfile = WScript.Arguments(a + 1) 
  		ElseIf lcase(arg) = "-a" Then
	  		'Delete files that are older than the number of days (postive value)
	  		iDeleteBackupsOlderThanXDays = WScript.Arguments(a + 1)      	       
       	ElseIf LCase(arg) = "-l" Then
       		'Enable logging
       		enableLog = True
        End If        
    Next
Else	
	TerminateScript("2")
End If

'==================================================================================================================================================
'***********' Section: Set Constants '***********' 
'==================================================================================================================================================
Const ForAppending = 8
Const ForWriting = 2
Const ForReading = 1
'This should be a negative value. I.e. to delete files older than 30 days, put a value of -30
Const DeleteLogsOlderThan = 5

'==================================================================================================================================================
'***********' Section: Delcare Global Variables '***********' 
'==================================================================================================================================================

Dim oShell
Dim oFileSystem
Dim File
Dim Folder
Dim oLog
Dim sLogPath
Dim oLogPath
Dim sLogFile
Dim oDirectoryToCheck
Dim oLogFolder
Dim dCurrentDate
Dim sErrText
Dim Errors
Dim sLogName
Dim Message
Dim DaysOld


'==================================================================================================================================================
'***********' Section: Setup and define Global Objects '***********' 
'==================================================================================================================================================

'Setup shell and file system access
Set oShell = CreateObject("Wscript.Shell")
Set oFileSystem = CreateObject("Scripting.FileSystemObject")

dCurrentDate = Date
sLogPath = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
sLogFile = "PruneSharePointBackups_" & Replace(Replace(Date, "\", "-"), "/", "-") & ".log"


'Setup logging if the logging argument was Set
If enableLog = True Then	
    Set oLog = oFileSystem.OpenTextFile(sLogPath & sLogFile, ForAppending, True)
    sLog("Script processing starting.")
    sLog("Arguments are: ")
    sLog("-p ('Path to the SharePoint backup xml file'): " & sTOCfile )    
    sLog("-a ('Delete files that are older than the number of days (postive value)'): " & iDeleteBackupsOlderThanXDays)         
    sLog("-l ('Enable Logging'): " & enableLog) 
End If

If IsEmpty(sTOCfile) = True Then	
	'No Directory specified - abort script
	TerminateScript("2")
End If

'Access files and folders
Set oFileSystem = CreateObject("Scripting.FileSystemObject")
'Set oDirectoryToCheck = oFileSystem.GetFolder(sTOCfile)
Set oLogFolder = oFileSystem.GetFolder(sLogPath)



'==================================================================================================================================================
'***********' Section: Run Sub Routines '***********' 
'==================================================================================================================================================


'Prune SharePoint Backups
Call PruneSharePointBackups(sTOCfile, iDeleteBackupsOlderThanXDays)

'Manage the scripts own log files
Call ManageOwnLogs(cint(DeleteLogsOlderThan))

'==================================================================================================================================================
'***********' Section: Cleanup '***********'
'==================================================================================================================================================
Call TerminateScript("0")

'==================================================================================================================================================
'***********' Sub Routines: '***********' 
'==================================================================================================================================================

Sub ManageOwnLogs(DeleteLogsOlderThan)
    On Error Resume Next    
	sLog("Cleaning up own log files.")
    For Each File In oLogFolder.files
        If strcomp(lcase(right(file.name, 4)), "mlog") = 0 Then
            Err.Clear
            DaysOld = DateDiff("d", File.DateLastModified, Now)                       
            If err.number <> 0 Then                
                sLog("Error: " & err.Description & " " & err.number)                
            Else
                If DaysOld > DeleteLogsOlderThan Then
                    'Delete the file
                    sLog(File.Name & " --- " & File.DateLastModified & " --- " & "This file is " & Replace(DaysOld, "-", "") & " days Old: " & " File will be deleted.")
                    file.Delete 
                    If err.number <> 0 Then                        
                        sLog("Error: " & err.Description & " " & err.number)                        
                    End IF          
                Else 
                    sLog(File.Name & " --- "  & File.DateLastModified & " --- " & "This file is " & Replace(DaysOld, "-", "") & " days Old: " & " File will be kept.")
                End If
            End If
        Else
            'sLog(File.Name & " --- " & File.DateLastModified & " --- " & "This file is not a .log backup file and will not be deleted.")
        End If    
    Next
End Sub

Sub TerminateScript(code)
	'Set 0 to indicate success, 1 for an unexpected error, 2 when arguments are incorrect.      
    If StrComp(code, "1") = 0 Then
    	If enableLog = True Then
			sLog("The script is terminating unexpectedly. This indicates that it encountered an unrecoverable error.")
		End If        
    ElseIf  StrComp(code, "2") = 0 Then 
    	WScript.Echo ("Aborting. Required arguments not supplied.") & vbcrlf & _
    	"Arguments are: " & vbcrlf & _
    	"-p 'Path to the SharePoint backup xml file': " & sTOCfile & vbcrlf & _    
    	"-a 'Delete backups that are older than the specified number of days (postive value)': " &  vbcrlf & _        
    	"-l 'Enable Logging'): " & enableLog		       	 
    ElseIf  StrComp(code, "0") = 0 Then
    	If enableLog = True Then
			sLog("Script finished processing successfully.")       		    
		End If             
    End If
    
	If enableLog = True Then
		sLog("Script Exiting" & vbcrlf)
		oLog.Close
	End If
	
    Set oArguments = Nothing
    Set oLog = Nothing    
    Set oShell = Nothing
    Set oFileSystem = Nothing
    Set oArguments = Nothing
    WScript.Quit(code)    
End Sub

Sub sLog(Logstring)
	'Write a line of text to the log file
    If enableLog = True Then
        oLog.writeline Date & " " & Time & " " & LogString        
    End If
End Sub

Sub PruneSharePointBackups (sTOCfile, iDeleteBackupsOlderThanXDays)
	On Error Resume Next
	Dim oXML
	Dim oXMLNode 
	Dim sDirToDelete
	Dim dMaxAge
	Dim dUKFormatedDate 
	
	Set oXML = CreateObject("Microsoft.XMLDOM")		
	dMaxAge = DateAdd("d",-(iDeleteBackupsOlderThanXDays), Now)
	
	sLog("Delete backups prior to: " & dMaxAge) 
	
	'Load XML File
	oXML.Async = false
	oXML.Load(sTOCfile)
	
	If oXML.ParseError.ErrorCode <> 0 Then
	    sLog("Error: Could not load the SharePoint Backup / Restore History. Reason: " & oXML.ParseError.Reason& ".") 
	    TerminateScript(1)
	End If	
	
	' Delete backup nodes that are older than the deletion date.
	For Each oXMLNode in oXML.DocumentElement.ChildNodes
		'SharePoint Backups are recorded using American date format. Reformat.
		Err.Clear
		Dim sDateParts
		Dim FinishTime,StartTime
		FinishTime = oXMLNode.SelectSingleNode("SPFinishTime").Text
		StartTime = oXMLNode.SelectSingleNode("SPStartTime").Text		
		if Err.Number = 0 Then
			sDateParts = Split(FinishTime, "/")
		else
			sDateParts = Split(StartTime, "/")
		End if
		dUKFormatedDate = sDateParts(1) & "/" & sDateParts(0) & "/" & sDateParts(2)		
	    If CDate(dUKFormatedDate) < Cdate(dMaxAge) Then
	        If oXMLNode.SelectSingleNode("SPIsBackup").Text = "True" Then
	        	sDirToDelete = mid(oXMLNode.SelectSingleNode("SPBackupDirectory").Text,1,len(oXMLNode.SelectSingleNode("SPBackupDirectory").Text)-1)
				sLog("Deleting backup with SPID: " & oXMLNode.SelectSingleNode("SPId").Text & " (Backup Finish Time: " & dUKFormatedDate & ")")				
				oFileSystem.DeleteFolder sDirToDelete	
	            sLog("Deleted: " & sDirToDelete)
	            oXML.DocumentElement.RemoveChild(oXMLNode)
	        End If 
	    Else
	    	  sLog("Keeping backup with SPID: " & oXMLNode.SelectSingleNode("SPId").Text & " (Backup Finish Time: " & dUKFormatedDate & ")")  
	    End If
	Next	
	' Save the XML file with the old nodes removed.
	oXML.Save(sTOCfile)	
End Sub
