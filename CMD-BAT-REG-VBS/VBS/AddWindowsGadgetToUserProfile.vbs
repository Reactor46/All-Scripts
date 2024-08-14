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
'The script shows how to add windows gadgets to user profile
Option Explicit 
Sub Main()
	Dim wsArgNum
	Dim objshell,CompName
	'Get the count of arguments
	wsArgNum = WScript.Arguments.Count
	'Create "wscript.shell" object
	Set objshell = CreateObject("Wscript.shell")
	Select Case wsArgNum
		Case 0 
			CompName = objshell.expandEnvironmentStrings("%Computername%")
			CopySidebarProfile(CompName)	 
		Case 1
			Dim InputText 
			InputText = WScript.Arguments(0)
			Dim objFSO,File
			'Create "Scripting.FileSystemObject" object
			Set objFSO = CreateObject("Scripting.FileSystemObject") 
			Set File = objFSO.OpenTextFile(InputText,1)
			'Get the computername in the text file
			Do Until File.AtEndOfStream
			CompName = File.ReadLine
			CopySidebarProfile(CompName)
			Loop
			File.Close
		Case Else 
			WScript.Echo "Please just drag one txt file to the script"
	End Select 
End Sub 

'This function is add the current user gadgets profile to all users
Function CopySidebarProfile(ComputerName)
	Dim objshell,envUSER,Source,objFSO 
	Dim UsersOfComp,UserFolders,Folder,SidebarPath,UserFolder
	Set objshell = CreateObject("Wscript.shell")
	'Get the current username 
	envUSER = objshell.expandEnvironmentStrings("%username%")
	'Get the user sidebar profile
	Source = "C:\Users\" & envUSER & "\AppData\Local\Microsoft\Windows Sidebar"
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
	UsersOfComp = "\\" & ComputerName & "\c$\Users\"
	If objFSO.FolderExists(UsersOfComp) = True Then 
		Set UserFolders = objFSO.GetFolder(UsersOfComp)
		For Each UserFolder In UserFolders.SubFolders
			Set Folder = objFSO.GetFolder(UserFolder.Path)
			SidebarPath = UsersOfComp & Folder.name & "\AppData\Local\Microsoft\Windows Sidebar"
			If objFSO.FolderExists(SidebarPath) = True And Folder.name <> envUSER  Then
				'Copy the sidebar profile
				objFSO.CopyFolder Source,SidebarPath
			End If 
		Next 
	Else 
		WScript.Echo "It can not connet to computer : " & ComputerName
	End If
End Function 

Call Main 