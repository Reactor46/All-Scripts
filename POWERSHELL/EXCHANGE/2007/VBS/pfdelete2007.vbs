days = wscript.arguments(1)
servername = wscript.arguments(0)
SB = 0
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\temp\PfDeletes.csv",2,true)
wfile.writeline("DateDeleted,FolderName,FolderID,DeletedBy-MailboxName,DeletedBy-UserName")
dtmStartDate = CDate(Date) - days
dtmStartDate = Year(dtmStartDate) & Right( "00" & Month(dtmStartDate), 2) & Right( "00" & Day(dtmStartDate), 2)
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & servername & "\root\cimv2")
Set colLoggedEvents = objWMIService.ExecQuery("Select Timewritten,Message from Win32_NTLogEvent Where Logfile='Application' and Eventcode = '9682' and TimeWritten >= '" & dtmStartDate & "' ",,48)
wscript.echo "Date Deleted	FolderName	FolderID	Deleter-Mbox	Deleter-UserName"
For Each objEvent in colLoggedEvents
    SB = 1
    Time_Written = cdate(DateSerial(Left(objEvent.TimeWritten, 4), Mid(objEvent.TimeWritten, 5, 2), Mid(objEvent.TimeWritten, 7, 2)) & " " & timeserial(Mid(objEvent.TimeWritten, 9, 2),Mid(objEvent.TimeWritten, 11, 2),Mid(objEvent.TimeWritten,13, 2)))
    FolderName = Mid(objEvent.Message,8,instr(objEvent.Message,"with folder")-9)
    FolderID = Mid(objEvent.Message,instr(objEvent.Message,"with folder ID")+15,(instr(objEvent.Message,"was deleted by")-(instr(objEvent.Message,"with folder ID")+16)))
    MailboxName = Mid(objEvent.Message,instr(objEvent.Message,"was deleted by")+15,instr(objEvent.Message,", user account")-(instr(objEvent.Message,"was deleted by")+15))
    UserName = Mid(objEvent.Message,instr(objEvent.Message,", user account")+15,(instr(objEvent.Message,".")-(instr(objEvent.Message,", user account")+15)))  
    wscript.echo Time_Written & "	" & FolderName & "	" & FolderID & "	" & MailboxName & "	" & UserName
    wfile.writeline(Time_Written & "," & FolderName & "," & FolderID & "," & MailboxName & "," & UserName)
next
wfile.close
set wfile = nothing
if SB = 0 then wscript.echo "No Public Folder Deletes recorded in the last " & days & " Days"