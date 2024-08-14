
'---------------------------------------------------------------------------------
' The sample scripts are not supported under any Microsoft standard support
' program or service. The sample scripts are provided AS IS without warranty
' of any kind. Microsoft further disclaims all implied warranties including,
' without limitation, any implied warranties of merchantability or of fitness for
' a particular purpose. The entire risk arising out of the use or performance of
' the sample scripts and documentation remains with you. In no event shall
' Microsoft, its authors, or anyone else involved in the creation, production, or
' delivery of the scripts be liable for any damages whatsoever (including,
' without limitation, damages for loss of business profits, business interruption,
' loss of business information, or other pecuniary loss) arising out of the use
' of or inability to use the sample scripts or documentation, even if Microsoft
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------
Option Explicit

Sub main()
	Dim objshell
	'Create wscript.shell
	Set objShell = CreateObject("Wscript.shell")
	Dim CurrentTime,command1,SysDrive,FolderPath,UserProfile,destination
	'Get the current date 
	CurrentTime = Date 
	'Start perfmon.exe by using powershell
	Command1 = "powershell.exe -nologo -command Start-Process -FilePath ""C:\windows\System32\perfmon"" -ArgumentList ""/report"" -WindowStyle Hidden"
	objShell.Run Command1, 0 
	WScript.Echo "It will take some time£¬Please be patient."
	'Script sleep 10 seconds
	WScript.Sleep 10000
	'Get the path of system drive
	SysDrive = objshell.ExpandEnvironmentStrings("%SystemDrive%")
	'Get the path of the report
	FolderPath = SysDrive & "\PerfLogs\System\Diagnostics"
	UserProfile = Objshell.ExpandEnvironmentStrings("%USERPROFILE%")
	Destination = UserProfile & "\DeskTop\" 
	Dim FSO,folders,folder,path 
	'Create scripting.filesystemobject 
	Set FSO = CreateObject("Scripting.filesystemobject")
	Set folders = FSO.GetFolder(FolderPath)
	For Each folder In folders.SubFolders 	
	'Get the newest folder 
	If DateDiff("s",folder.DateCreated,CurrentTime) < 0  Then 
	 	CurrentTime  = folder.DateCreated
		Path = Folder.Path
	End If 
	Next 
	WScript.Sleep 110000
	Dim FilePath,oShell,Test
	Filepath = path & "\report.html"
	'Copy the report to desktop by using powershell
 	Set oShell = CreateObject("Shell.Application")  
 	oShell.ShellExecute "powershell", "-command Copy-Item -Path " & Filepath & " -Destination " & Destination & " -Force", "", "runas", 0
 	If Err.Number = 0 Then 
		WScript.Echo "Generate system health report successfully."
	Else 
		WScript.Echo "Fail to generate system health report."
	End If 
End Sub 

Call main 

