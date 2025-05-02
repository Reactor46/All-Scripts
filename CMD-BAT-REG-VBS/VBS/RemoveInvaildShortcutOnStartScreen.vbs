'#--------------------------------------------------------------------------------- 
'#The sample scripts are not supported under any Microsoft standard support 
'#program or service. The sample scripts are provided AS IS without warranty  
'#of any kind. Microsoft further disclaims all implied warranties including,  
'#without limitation, any implied warranties of merchantability or of fitness for 
'#a particular purpose. The entire risk arising out of the use or performance of  
'#the sample scripts and documentation remains with you. In no event shall 
'#Microsoft, its authors, or anyone else involved in the creation, production, or 
'#delivery of the scripts be liable for any damages whatsoever (including, 
'#without limitation, damages for loss of business profits, business interruption, 
'#loss of business information, or other pecuniary loss) arising out of the use 
'#of or inability to use the sample scripts or documentation, even if Microsoft 
'#has been advised of the possibility of such damages 
'#--------------------------------------------------------------------------------- 

'This VBScript is used to delete invalid shortcuts on Start Menu and All Apps in Windows 8
Option Explicit 
On Error Resume Next 
Dim objFSO,objshell,envUSER
Dim objUserFolder,objSysFolder,UserFolder,SysFolder 	
'Create "Scripting.FileSystemObject" object 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
'Create "Wscript.shell" object 
Set objshell = CreateObject("wscript.shell")
'Get the current username
envUSER = objshell.expandEnvironmentStrings("%username%")
'The path of personal start menu 
objUserFolder = "C:\Users\" & envUSER & "\AppData\Roaming\Microsoft\Windows\Start Menu"
'The path of the system start menu
objSysFolder = "C:\ProgramData\Microsoft\Windows\Start Menu"
Set UserFolder = objFSO.GetFolder(objUserFolder)
DeleteShortcut UserFolder
Set SysFolder = objFSO.GetFolder(objSysFolder)
DeleteShortcut SysFolder

'This sub is used to find the invalid shortcut and delete it in specified folder
Sub DeleteShortcut(Folder)
	Dim Subfolder,objFolder,colFiles
	Dim objshortcut,filepath,objFile
	'This for-each loop is used to get all files in the specified folder and its subfolders
    For Each Subfolder in Folder.SubFolders
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files 
        For Each objFile in colFiles
        	If  objFile.Name <> "Desktop.lnk"  And  UCase(objFile.Type) = "SHORTCUT" Then 
        		Set objshortcut = objshell.CreateShortcut(objFile.Path)
        		filepath =  objshortcut.TargetPath 
        		If objFSO.FileExists(filepath) = False Then 
        		    WScript.Echo  "Removing invalid ShortCut :" & objFile.Path
        		 	objFSO.DeleteFile(objFile.Path) 
        		End If 
        	End If    
        Next
        DeleteShortcut Subfolder
    Next
End Sub