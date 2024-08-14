'--------------------------------------------------------------------------------- 
'The sample scripts are not supported under any Microsoft standard support 
'program or service. The sample scripts are provided AS IS without warranty  
'of any kind. Microsoft further disclaims all implied warranties including,  
'without limitation, any implied warranties of merchantability or of fitness for 
'a particular purpose. The entire risk arising out of the use or performance of  
'the sample scripts and documentation remains with you. In no event shall 
'Microsoft, its authors, or anyone else involved in the creation, production, or 
'delivery of the scripts be liable for any damages whatsoever (including, 
'without limitation, damages for loss of business profits, business interruption, 
'loss of business information, or other pecuniary loss) arising out of the use 
'of or inability to use the sample scripts or documentation, even if Microsoft 
'has been advised of the possibility of such damages 
'--------------------------------------------------------------------------------- 
Option Explicit

Dim ProgramPath

ProgramPath = InputBox("Please enter the full path of program that you want to run as administrator with UAC promts:")

If IsEmpty(ProgramPath) Then
	WScript.Quit
Else
	CreateShortcut ProgramPath
End If

Sub CreateShortcut(ProgramPath)
	Dim objShell
	Dim objWSHShell
	Dim objFSO
	
	'Create an instance of the Shell.Application
	Set objShell = CreateObject("Shell.Application")
	
	Set objWSHShell = CreateObject("Wscript.Shell")
	
	'Create an instance of the Scripting.FileSystemObject
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Dim WshNetwork
	Dim currentDirectory
	Dim ComputerName
	
	Set WshNetwork = CreateObject("WScript.Network")
	ComputerName = WshNetwork.ComputerName
	
	'Get current folder location
	currentDirectory = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
	
	' ####################################
	' Create Windows shortcut of Program #
	' ####################################
	Dim ProgramShortcutPath
	Dim objProgramShortcut
	Dim objFile
	Dim FileName
	Dim objBaseName
	Dim SystemRoot
	
	objFile = objFSO.GetFile(ProgramPath)
	objBaseName = objFSO.GetBaseName(objFile)
	
	
	SystemRoot = objWSHShell.expandEnvironmentStrings("%SystemRoot%")
	
	'create a new shortcut of program
	ProgramShortcutPath = currentDirectory & objBaseName & ".lnk"
	Set objProgramShortcut = objWSHShell.CreateShortcut(ProgramShortcutPath)
	objProgramShortcut.TargetPath = SystemRoot & "\system32\runas.exe"
	objProgramShortcut.Arguments = "/user:" & ComputerName &"\Administrator /savecred " & ProgramPath
	
	'change the default icon of shortcut
	objProgramShortcut.IconLocation = SystemRoot & "\System32\SHELL32.dll,111"
	objProgramShortcut.Save
	WScript.Echo "Successfully creatred a shortcu of program on " & currentDirectory
End Sub