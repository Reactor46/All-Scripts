dim wshShell, filesys, newfolder
set wshShell = WScript.CreateObject( "WScript.Shell" )
set filesys = CreateObject("Scripting.FileSystemObject")

dim Logs
dim Servers
dim strLogDir
dim boolClear
dim intDays

dim strDate, thisDay, thisMonth, thisYear
thisDay = Day(date)
thisMonth = Month(date)
thisYear = Year(date)
strDate = thisMonth & "-" & thisDay & "-" & thisYear & "-" & Timer

'List of Event Logs to Backup
Logs=Array("Application", "Security", "System")

'List of Servers to retrieve Logs from remotely
Servers=Array(".")
'Servers=Array("Svr1", "Svr2")

'Use following for running on individual PC
Servers=Array(wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" ))

'Path for Saving Logs
strLogDir = "\\server\logdir"

'Clear Event logs after backup TRUE/FALSE
boolClear = True

'Number of days to save files
intDays = 365

For Each svr in Servers
	For Each evtLog in Logs		
		Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate,(Backup)}!\\" & svr & "\root\cimv2")
		Set colLogFiles = objWMIService.ExecQuery("Select * from Win32_NTEventLogFile where LogFileName='" & evtLog & "'")
		
		'Save Log file as "ServerName_Log_Date
		For Each objLogfile in colLogFiles		
			'Check to see if Folder for Server exists, if Not, create one			
			If Not filesys.FolderExists(strLogDir & svr) Then
				newfolder = filesys.CreateFolder(strLogDir & svr)
			End If
			
			'Name and backup Event Log
			errBackupLog = objLogFile.BackupEventLog(strLogDir & svr & "\" & svr & "_" & evtLog & "_" & strDate & ".evt")
			If errBackupLog <> 0 Then        
				Wscript.Echo "The " & evtLog & " Event log could not be backed up."
			Else
				'Clear Event Log if boolClear = True
				If boolClear then
					objLogFile.ClearEventLog()
				End If
			End If			
		Next		
	Next
	
	'Keep for x days then remove from individual folders
	dim files, folder
	set folder = filesys.GetFolder(strLogDir & svr)
	for each files in folder.Files
		If files.DateLastModified < (Now - intDays) then
			files.Delete
		End If
	Next

Next	

WScript.Quit

