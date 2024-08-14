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
'################################################
'This script is to add "Scan with Windows Defender" to context menu
'################################################

If WScript.Arguments.Count = 0 Then
	Dim objShell
	Set objShell = CreateObject("Shell.Application")
	objShell.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "runas", 1
Else

'################################################
' This script is to verify if the registry key exists
'################################################


Function KeyExists(Path)
	On Error Resume Next 
	Dim objshell,Flag,value
	Set objShell = CreateObject("WScript.Shell")
	value = objShell.RegRead(Path) 
	Flag = False 
	If Err.Number = 0 Then 	
	 	Flag = True 
	End If
	Keyexists = Flag
End Function 

'################################################
'This function is to delete registry key.
'################################################
Sub DeleteSubkeys(HKEY_CLASSES_ROOT, strKeyPath) 
	Dim strSubkey,arrSubkeys,strComputer,objRegistry
	strComputer = "."
	Set objRegistry = GetObject("winmgmts:\\" & _
    strComputer & "\root\default:StdRegProv")
    objRegistry.EnumKey HKEY_CLASSES_ROOT, strKeyPath, arrSubkeys 
    If IsArray(arrSubkeys) Then 
        For Each strSubkey In arrSubkeys 
            DeleteSubkeys HKEY_CLASSES_ROOT, strKeyPath & "\" & strSubkey 
        Next 
    End If 
    objRegistry.DeleteKey HKEY_CLASSES_ROOT, strKeyPath 
End Sub

Dim thekey, Keyvalue, CurrentDirectory, Windir
Dim FSObject, strKeyPath, PSPath, strComputer,objRegistry 
Const HKEY_CLASSES_ROOT = &H80000000
Set objshell = CreateObject("Wscript.shell") 'Create "wscript.shell" object
Windir = objshell.expandEnvironmentStrings("%Windir%") & "\System32\" 
strComputer = "."
CurrentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName))) 'Get the script current directory
Dim source 
Source = CurrentDirectory & "\scan.ps1"
Set FSObject = CreateObject("Scripting.FileSystemObject") 'Create "scripting.FileSystemObject" object 
FSObject.CopyFile Source, Windir
PSPath = Windir & "scan.ps1"
Dim Path1, path2
Path1 = "HKEY_CLASSES_ROOT\Folder\shell\WindowsDefender\"
Path2 = "HKEY_CLASSES_ROOT\*\shell\WindowsDefender\"
'Add Windows Defender 
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	If KeyExists(Path1) then 	'Verify if the key has existed	
		Dim Choice 
		Choice = MsgBox("""Scan with Windows defender"" exists, do you want to remove it?",4, "system message")
		If Choice = vbYes Then 
			'Delete Windows Defender
			If FSObject.FileExists(PSPath) Then 
				FSObject.DeleteFile(PSPath)
			End If 	
			DeleteSubkeys HKEY_CLASSES_ROOT, "Folder\shell\WindowsDefender"
			DeleteSubkeys HKEY_CLASSES_ROOT, "*\shell\WindowsDefender"
			If KeyExists(Path1) And KeyExists(Path2) Then
				MsgBox "Fail to delete 'Scan with Windows Defender' from context menu."
			Else 
				MsgBox "Delete 'Scan with Windows Defender' from context menu successfully."
			End If 
		End If 
	Else 
		Dim strKeyPath1, strKeyPath2, CommandValue
		strKeyPath1 = "Folder\shell\WindowsDefender"
		strKeyPath2 = "*\shell\WindowsDefender"
		If FSObject.FileExists(PSPath) Then 
			FSObject.DeleteFile(PSPath)
		End If 
		If FSObject.FileExists(Source) Then
			FSObject.CopyFile Source, Windir  'Copy the vbscript to "C:\windwos\system32"
			CommandValue = "powershell.exe """ & PSPath & " -file %1"""
			objRegistry.CreateKey HKEY_CLASSES_ROOT, strKeyPath1   'Create a HiddenFiles registry key
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, strKeyPath1, "Icon", "%ProgramFiles%\\Windows Defender\\EppManifest.dll"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, strKeyPath1, "MUIVerb", "Scan with Windows Defender"
			objRegistry.CreateKey HKEY_CLASSES_ROOT, strKeyPath1 & "\Command"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, strKeyPath1 & "\Command", "", CommandValue
			
			
			objRegistry.CreateKey HKEY_CLASSES_ROOT, strKeyPath2   'Create a HiddenFiles registry key
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, strKeyPath2, "Icon", "%ProgramFiles%\\Windows Defender\\EppManifest.dll"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, strKeyPath2, "MUIVerb", "Scan with Windows Defender"
			objRegistry.CreateKey HKEY_CLASSES_ROOT, strKeyPath2 & "\Command"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, strKeyPath2 & "\Command", "", CommandValue
			
			If KeyExists(Path1) And KeyExists(Path2)  Then  'Verify if the key is created
				MsgBox "Add 'Scan with Windows Defender' to context menu successfully."
			Else 
				MsgBox "Fail to Add 'Scan with Windows Defender' to context menu."
			End If 
		Else 
			MsgBox "Not find 'scan.ps1' file, failed to execute script."
		End If 
	End If 



		

End If 