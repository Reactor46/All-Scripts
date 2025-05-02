On Error Resume Next
Set FSO = CreateObject("Scripting.FileSystemObject")													' class used for getting file info
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set logFile = FSO.OpenTextFile("\\NETWORK\PATH\TO\LOGFILE\FileCheck.txt", ForAppending, True) 			' create log file					
Set WshNetwork = WScript.CreateObject("WScript.Network")												' class used for getting computer name

filespec = "C:\FILEPATH\FILENAME"																		' path to file
today = date																							' create variable out of todays date
compName = WshNetwork.ComputerName																		' get the computer name of checked computer
retention = 10																							' set retention period 
if FSO.FileExists(filespec) then 																		' check if file exists
		 set file =	FSO.GetFile(filespec)																' create object out of file
			FileModified = FormatDateTime(file.DateLastModified, "2")                         			' get the Last Date Modified from specified file
		If DateDiff("d", FileModified, today) > retention then											' file has been modified out of retention period, do this
					logFile.Write compName & " " & "is NOT updated!" & VbCrLf  
				else																					' if file has been modified within the retention period, do this
					logFile.Write compName & " " & "is updated!" & VbCrLf  
			End if
			
else
logFile.Write compName & " " & "- This file does not exist!" & VbCrLf									' if file does not exist, write that in log	
End if
logFile.Close																							' close the log file