on error resume next

' ########################################
'set up variables
' ########################################




lDate = "30/1/2000"
uDate = "10/11/2010"

lDate = CDate(lDate)
uDate = CDate(uDate)
' ########################################
' set file name for output text file
' ########################################
wscript.echo "Delete Files Between" & ldate & " - " & uDate


dtmThisDay = Day(Date)
dtmThisMonth = Month(Date)
dtmThisYear = Year(Date)
strBackupName =  "Deleted files " & dtmThisDay & "_" & dtmThisMonth & "_" & dtmThisYear & ".txt"


Set fso = CreateObject("Scripting.FileSystemObject") 

' ########################################
'setup output file
' ########################################


Set objInputFile = fso.OpenTextFile("FolderList.txt", 1, True) 
Set objOutputFile = fso.OpenTextFile(strBackupName, 2, True) 
 
' ########################################
'Main Program Loop
' ########################################


Do While objInputFile.AtEndOfLine <> True 
    strFoldername = objInputFile.ReadLine 
    Set objDIR = FSO.GetFolder(strFoldername)
sDIR = STRFOlderName

wscript.echo "Deleting files From " & objDIR

GoSubFolders objDIR


 
Loop

wscript.echo "please keep a copy of the file " & strbackupname & " located in this script folder"
' ########################################
' Endo of main program Loop
' ########################################




' ########################################
' Sub Routines called by main program loop
' ########################################

' ########################################
'Main SubRoutine
' ########################################

Sub MainSub (objDIR)

For Each efile in objDIR.files
'objoutputfile.writeline (efile) & DateLastModified

if ldate = Null and udate = Null Then
	If efile.DateLastAccessed = Date Then
		objOutputfile.writeline(efile)
		FSO.DeleteFile efile, True
wscript.echo "1"
	End If

ElseIf ldate <> Null and uDate = Null Then
	If efile.DateLastAccessed < lDate then
		objOutputfile.writeline(efile)
		FSO.DeleteFile efile, True
wscript.echo "2"
	End IF

ElseIf lDate = Null and uDate <> Null Then
	If efile.DateLastAccessed > uDate Then
		objOutputfile.writeline(efile)
		FSO.DeleteFile efile, True
wscript.echo "3"
	End IF

ElseIf lDate = uDate Then
	if efile.DateLastAccessed = lDate Then
		objOutputfile.writeline(efile)
		FSO.DeleteFile efile, True
wscript.echo "4"
	End IF

Else
	If efile.DateLastAccessed > lDate and efile.DateLastAccessed <uDate Then
	objOutputfile.writeline(efile)
	FSO.DeleteFile efile, True
wscript.echo "5"
	End IF
End IF
Next
End Sub




' ########################################
' Sub for recursive folder search
' ########################################

Sub GoSubFolders (objDIR)
on Error Resume Next
	if objDIR <> "\System Volume Information" Then 
		MainSub objDIR
		For each eFolder in objDIR.SubFolders 
			'objOutputfile.writeline(efolder)
'wscript.echo efolder
			GoSubFolders eFolder
		Next
	End IF
End Sub

Sub DelFile(sfile)
FSO.DeleteFile sfile, True
end sub
