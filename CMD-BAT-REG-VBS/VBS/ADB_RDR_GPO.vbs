' This script will take any Adobe Reader EXE file and create an Administrative Install Point, MSP patches applied.
' 
' The basic steps are as follows:
' 1) Extract the EXE to a temporary location (in the base directory of the script)
' 2) Extract the MSI from the temporary extract location to an AIP directory (base directory of script), and into a versioned sub-directory
' 3) Apply any included MSP patches from the temporary directory to the MSI in the AIP directory (in order!)
' 4) Copy the setup.ini file from the temporary directory to the AIP directory
' 
' Tested on Windows 7 Pro, 64-bit with 11.0.00, 11.0.02, 11.0.07
' 
' Feel free to make any modifications or changes as needed.
' This script is "AS-IS" with no guarantees.  Use at your own discretion.
' 
' http://www.adobe.com/devnet-docs/acrobatetk/tools/AdminGuide/aip.html
' http://www.adobe.com/devnet-docs/acrobatetk/tools/AdminGuide/basics.html
' 
' @author Jon Dehen

Option Explicit
Dim WshShell, objFSO
Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Make sure the script is launched as elevated
If WScript.Arguments.Named.Exists("elevated") = False Then
	' Use the full UNC path to relaunch the program since we have no way to share drive maps between elevated and non-elevated
	' FYI a network drive is indeed mapped once the scripted begins execution
	CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & forceUNCRef(WScript.ScriptFullName) & """ /elevated", "", "runas", 1
	WScript.Quit
End If

' ================================
' Initialization the script and base directories
' ================================
Dim strBaseDir, strNewEXE

' Get the base directory for the script
strBaseDir = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\"

' Get the path to our new Adobe EXE installer
MsgBox "Please select the Adobe Reader EXE file.", 64, "Adobe Reader Updater"
strNewEXE = ChooseFile()

' Obtain the version number, formatted as XX.X.XX
Dim strFileVersion
strFileVersion = getFileVersion(strNewEXE)

' Create the base AIP directory if not existing
Dim strAIPDir
If Not objFSO.FolderExists(strBaseDir & "AIP\") Then
	objFSO.CreateFolder (strBaseDir & "AIP")
End If

' Get our AIP versioned directory
strAIPDir = strBaseDir & "AIP\" & strFileVersion

' Clear any existing AIP versioned directory, just to be safe
If objFSO.FolderExists(strAIPDir) Then
	objFSO.DeleteFolder(strAIPDir)
End If

' Create the new AIP versioned directory
objFSO.CreateFolder(strAIPDir)

' ================================
' Start the extractions
' ================================
Dim strExtractDir
strExtractDir = strBaseDir & "Extract"

' Extract the EXE contents to a temp directory
WshShell.Run "cmd.exe /c """"" & strNewExe & """ -sfx_o""" & strExtractDir & """ -sfx_ne -sfx_nd""", 0, True

' Create the AIP extraction from the extracted MSI file in the temp directory
WshShell.Run "cmd.exe /c msiexec /a """ & strExtractDir & "\AcroRead.msi"" /qr TARGETDIR=""" & strAIPDir & """", 0, True

' Copy setup.ini into AIP directory (needed for transforms)
objFSO.CopyFile strExtractDir & "\setup.ini", strAIPDir & "\"

' ================================
' Apply our MSP patches to our AIP
' ================================
Dim strMSPExt, objExtractDir, objFile, strMSPFile

' We need to search file the "msp" file extension
Set objExtractDir = objFSO.GetFolder(strExtractDir)
Set strMSPExt = CreateObject("Scripting.Dictionary")
strMSPExt.CompareMode = vbTextCompare  'case-insensitive
strMSPExt.Add "msp", True

' Rename all MSP files so they're sorted in the correct order (alphabetical) for applying
For Each objFile In objExtractDir.Files
	If strMSPExt.Exists(objFSO.GetExtensionName(objFile.Name))  Then
		strMSPFile = strExtractDir & "\" & objFile.Name
		objFSO.MoveFile strMSPFile, strExtractDir & "\" & getFileVersion(objFile.Name) & ".msp"		
	End If	
Next

' Now apply any MSP files to the AIP, as found alphabetically
For Each objFile In objExtractDir.Files
	If strMSPExt.Exists(objFSO.GetExtensionName(objFile.Name))  Then
		strMSPFile = strExtractDir & "\" & objFile.Name
		WshShell.Run "cmd.exe /c msiexec /a """ & strAIPDir & "\AcroRead.msi"" /qr /p """ & strMSPFile & """ TARGETDIR=""" & strAIPDir & """", 0, True
	End If	
Next

' Delete the temporary folder
objFSO.DeleteFolder(strExtractDir)

' We're all done, prompt with verification
MsgBox "Adobe Reader " & strFileVersion & " AIP created and patched successfully!", 64, "Adobe Reader Updater"


' ================================
' Functions
' ================================

' Function does not alter paths containing non-network drives or paths that are already UNCs - safe to use on all paths
Function forceUNCRef(byref path)
	forceUNCRef = vbEmpty
	
	If objFSO.GetDrive(objFSO.GetDriveName(path)).drivetype <> 3 OR left(objFSO.GetDriveName(path),2) = "\\" Then
		forceUNCRef = path
		Exit Function
	Else
		Dim WshNetwork, objDrives, UNC
		Set WshNetwork = CreateObject("WScript.Network") 
		Set objDrives = WshNetwork.EnumNetworkDrives 
		
		Dim i
		For i = 0 to (objDrives.Count - 1) Step 2
			If lcase(objDrives.Item(i)) = lcase(objFSO.GetDriveName(path)) Then 
				UNC = LCase(objDrives.Item(i+1))
				forceUNCRef = replace(path,objFSO.GetDriveName(path),UNC)
				Exit Function
			End If    
		Next   
		If IsEmpty(forceUNCRef) Then
			forceUNCRef = path 
		End If
	End If
	
End Function

' Prompts user with a select file dialogue (thanks HTML!)
Function ChooseFile()
	Dim WshShell, WshExec
	Set WshShell = CreateObject( "WScript.Shell" )
	Set WshExec = WshShell.Exec("mshta.exe ""about: <input type=file id=X><script>X.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(X.value);close();resizeTo(0,0);</script>""" )
	
	ChooseFile = Replace(WshExec.StdOut.ReadAll, vbCRLF, "" )	
End Function

' Returns the version of a given EXE file
Function getFileVersion(strFileName)	
	' Acrobat and Reader version numbers contain 3 integers and two dots:
	' The first integer identifies the major release; for example, 9.0.0, 10.0.0, or 11.0.0 (sometimes XX.0 for short).
	' The second integer is incremented when a quarterly update is delivered as a full installer; for example, 9.1.0, 9.2.0, and 10.1.0. These always include previously released out of cycle patches.
	' The third integer is incremented for all other updates and out of cycle patches; for example, 9.4.3, 9.4.4, 10.0.1, 10.0.2, etc. Note that 11.0 products use the format: 11.x.xx (with two integers identifying minor updates starting with 01).
	Dim strRegExp, strFileVersion, arrMatches
	Set strRegExp = New RegExp
	strRegExp.Pattern = "\d{5}" ' grab the 5 digit version number

	' Strip out path if a full path is provided (in case the path has 5 digits in it)
	If Not InStrRev(strFileName, "\") = 0 Then
		Set arrMatches = strRegExp.Execute(Mid(strFileName, InStrRev(strFileName, "\")))
	Else
		Set arrMatches = strRegExp.Execute(strFileName)
	End If
	
	' We should have only one match, the version in format XXXXX
	strFileVersion = arrMatches(0)
	
	' Format the version number, based on the above information, to get XX.X.XX	
	strFileVersion = Mid(strFileVersion, 1, 2) & "." & Mid(strFileVersion, 3, 1) & "." & Mid(strFileVersion, 4, 2)
	
	getFileVersion = strFileVersion
End Function