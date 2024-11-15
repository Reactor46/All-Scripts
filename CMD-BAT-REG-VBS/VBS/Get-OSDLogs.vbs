'*******************************************************************************
'* File:	Get-OSDLogs.vbs
'* Author:	Jonathon Mitchell	
'* Purpose:	Gets all "LOG" files and creates a copy keeping file path
'* Usage:       cscript.exe Get-OSDLogs.vbs
'* Version:     1.0
'* History:		
'*		Version		Author		Date			What
'*		1.0		JM219		01/10/2013	Initial Version
'*******************************************************************************

On Error Resume Next

Dim oDrive, sComputerName, sLogFolder, sLogStore 

sLogFolder = "<Enter Your Path Here>"

Set oFSO = CreateObject("Scripting.FileSystemObject") 

sComputerName = InputBox("Please enter the name of the computer","LogNabber")
sLogStore = sLogFolder & "\" & sComputerName
If Not oFSO.FolderExists(sLogStore) Then
	oFSO.CreateFolder(sLogStore)
End If

For Each oDrive In oFSO.Drives
	WScript.Echo "Checking for log files on " & oDrive.Path
	GetLogs oDrive.Path
Next

Sub GetLogs(ByVal rootName)

	On Error Resume Next

	Dim oFile, oFolder, cPathFolder, sPath
	
	Set cDrive = oFSO.GetFolder(rootName)

	If cDrive.Files.Count >= 1 Then
		For Each oFile In cDrive.Files
			If UCase(Right(oFile.Name,3)) = "LOG" Then
				WScript.Echo oFile.Name & " to be copied..."
				cPathFolder = Replace(Replace(oFile.Path,oFile.Name,""),":","")
				CreateFolders(sLogStore & "\" & cPathFolder)
				If Not oFSO.FolderExists(cPathFolder) Then
					oFSO.CreateFolder(cPathFolder)
				End If
				oFSO.CopyFile oFile.Path, sLogStore & "\" & cPathFolder & oFile.Name
			End If
		Next
	End If
	
	For Each oFolder In cDrive.SubFolders
		GetLogs oFolder.Path
	Next

End Sub

Sub CreateFolders(byval sPath)
    
    Dim nLevel, iLevel, j 
	
	nLevel = 0
	
	strParentPath = sPath
	Do Until strParentPath = ""
	    strParentPath = oFSO.GetParentFolderName(strParentPath)
	    nLevel = nLevel + 1
	Loop
	
	For iLevel = 1 To nLevel
	    strParentPath = sPath
	    For j = 1 To nLevel - iLevel
	        strParentPath = oFSO.GetParentFolderName(strParentPath)
	    Next
	
	    If oFSO.FolderExists(strParentPath) = False Then
	        Set newFolder = oFSO.CreateFolder(strParentPath)
	    End If
	Next
   	
End Sub
