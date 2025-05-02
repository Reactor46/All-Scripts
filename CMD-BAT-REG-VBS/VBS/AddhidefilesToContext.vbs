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
'This script is to add "Show/HideFiles" to context menu
'################################################
Sub main()

Dim UserName
	Dim objshell, thekey, setHidden, CurrentDirectory
	Dim FSObject, strKeyPath, VBPath, Path, input
	Const HKEY_CLASSES_ROOT = &H80000000
	Set objshell = CreateObject("Wscript.shell") 'Create "wscript.shell" object
	UserName = objshell.expandEnvironmentStrings("%UserName%")
	theKey = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden"
	setHidden = objshell.RegRead(theKey) 'Read the hidden value 
	CurrentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName))) 'Get the script current directory
	Set FSObject = CreateObject("Scripting.FileSystemObject") 'Create "scripting.FileSystemObject" object 
	strKeyPath = "Directory\Background\shell\HidenFiles" 
	VBPath = "C:\Users\" & UserName & "\HideFiles.vbs"
	Path = "HKEY_CLASSES_ROOT\Directory\Background\shell\HidenFiles\"
	Input = InputBox("Enter 'A'('D') to add(delete) 'Show(Hide)HidenFiles'","Enter your choice")
	If UCase(input) = UCase("a") Then  'Add "Show/HideFiles" to context menu
		Dim strComputer, strValueName, objRegistry, CommandValue
		strComputer = "."
		strValueName = ""
		Set objRegistry = GetObject("winmgmts:\\" & _
	    strComputer & "\root\default:StdRegProv")
		If KeyExists(Path) Then 	'Verify if the key has existed
			MsgBox "'Show/HideFiles' has existed."
			WScript.Quit
		Else 
			If FSObject.FileExists(VBPath) Then 
				FSObject.DeleteFile(VBPath)
			End If 
			FSObject.CopyFile CurrentDirectory & "\HideFiles.vbs","C:\Users\" & UserName & "\"  'Copy the vbscript to "C:\windwos\system32"
			If FSObject.FileExists(VBPath) Then 
				CommandValue = "wscript.exe " & VBPath
				objRegistry.CreateKey HKEY_CLASSES_ROOT, strKeyPath   'Create a HiddenFiles registry key
				objRegistry.SetStringValue HKEY_CLASSES_ROOT, strKeyPath, strValueName, "Show/HideFiles"
				objRegistry.CreateKey HKEY_CLASSES_ROOT, strKeyPath & "\Command"
				objRegistry.SetStringValue HKEY_CLASSES_ROOT, strKeyPath & "\Command", strValueName, CommandValue
				If KeyExists(Path) Then  'Verify if the key is created
					MsgBox "Add 'Show/HideFiles' to Context successfully."
				Else 
					MsgBox "Fail to Add 'Show/HideFiles' to Context."
				End If 
			Else 
				MsgBox "Can not copy the VBScript to C:\windows\system32"
			End If 
		End If 
	ElseIf UCase(input) = UCase("d") Then   'Delete "Show/HideFiles" from context menu
			If KeyExists(Path) Then 
				If FSObject.FileExists(VBPath) Then 
					FSObject.DeleteFile(VBPath)
				End If 	
				DeleteSubkeys HKEY_CLASSES_ROOT, strKeyPath
				If KeyExists(Path) Then
					MsgBox "Fail to delete 'Show/HideFiles' from Context."
				Else 
					MsgBox "Delete 'Show/HideFiles' from Context successfully."
				End If 
			Else 
				MsgBox "Can not find 'Show(Hide)HiddenFils' from Context."
			End If 
	ElseIf IsEmpty(input) Then 
		WScript.Quit
	Else  
		MsgBox "Invalid input,please try again."
	End If 
		
End Sub 

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

Call main 