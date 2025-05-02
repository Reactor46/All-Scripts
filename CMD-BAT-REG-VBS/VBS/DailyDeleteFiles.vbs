'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' Http://assaf.miron.googlepages.com
' Date : 25/06/2007
' DailyDeleteFiles.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
On Error Resume Next

' Defining & initializing Variables

Const DeleteReadOnly = TRUE
Const ForWriting = 2
Const strServerPath = "\\Server\MainFolder"
Const LogName = "DBLogs"

Dim objFSO

' Setting the Log File object
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("DailyDeleteLog-" & LogName & ".txt", ForWriting, True)


' Procedures and Functions

Public Sub ShowSubFolders(strFolder,intFolder)
' This Function Searches A Folder and Its Sub Folders for 
	Dim SubFolder
	Dim objFolder
	
	Set objFolder = objFSO.GetFolder(strFolder)

    For Each Subfolder in objFolder.SubFolders
    	' Delete All Files in This Sub Folder that are Older than 1 Day
		DeleteFilesInFoder SubFolder.Path,"d",1 
		intFolder = intFolder + 1
	    ShowSubFolders Subfolder,intFolder ' Call The Function Recursevlly For each Sub Folder
		intFolder = intFolder - 1
	Next

' If The Folder Count is Greater than 0 - We are in a Sub Folder
If intFolder > 0 Then
    Set objFolder = objFSO.GetFolder(strFolder)
    Set colFiles = objFolder.Files
    ' If No Files in The Folder - Delete the Folder
	If colFiles.Count = 0 Then
		DeleteFolders(objFolder)    
	End If
End If

End Sub

Sub DeleteFilesInFoder (strFolder,DateInterval,Diff)
' This Sub Deletes All Files in  A folder that Exceed the Diff Date Interval
	Dim objFolder
	Dim colFiles
	Dim objFile
	Dim strDate

	Set objFolder = objFSO.GetFolder(strFolder)
	Set colFiles = objFolder.Files

	strDate = NOW	
	For Each objFile in colFiles
		objTextFile.Write objFile.Path
		If IsDateDiff(DateInterval,objFile.DateCreated,strDate,Diff) Then
			Err = objFSO.DeleteFile(objFile.Path, DeleteReadOnly)
			If Err.Number = 0 Then
				objTextFile.WriteLine " - Deleted"
			Else
				objTextFile.WriteLine " - Error ! - " & Err.Description
			End If
		End If
		
		objTextFile.WriteLine
	Next
End Sub

Public Sub DeleteFolders(Folder)
' This Sub Delets A Folder
    Set objFolder = objFSO.GetFolder(Folder.Path)
	objTextFile.Write objFolder.Path & " - Deleted"
	objTextFile.writeLine
	objFSO.DeleteFolder(objFolder.Path)

End Sub

Public Function IsDateDiff(interval,strDate1,strDate2,diff)
	' Function Recieves Two Dates, an Interval of Time (d - Days, w - Weeks) and The Differrent Time Between Them
	' Function Returns True or False Wheter the Time has Exceeded or not
	Dim objDate1
	Dim objDate2
	Dim intDiff
	Dim returnedValue
	' Convert String Date To Date Time
	objDate1 = CDate(strDate1)
	objDate2 = CDate(strDate2)
	' Check Date Time Differences
	intDiff = DateDiff(interval,objDate1,objDate2)
	
	If Abs(intDiff) >= Abs(diff) Then
		returnedValue = True
	Else
		returnedValue = False
	End If
	' Return Result
	IsDateDiff = returnedValue 
	
End Function


'Main 
objTextFile.WriteLine "Log Started " & Now

' Delete All Files in This Sub Folder that are Older than 1 Day
DeleteFilesInFoder strServerPath,"d",1

intFolder = 0
' Run Recursevlly on the Sub Folders
ShowSubfolders strServerPath,intFolder

objTextFile.WriteLine
objTextFile.WriteLine "Log Ended " & Now
objTextFile.Close
