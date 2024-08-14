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
	Dim Count,Flag 
	Dim objshell, thekey, setHidden, CurrentDirectory,strComputer,objRegistry
	Count = WScript.Arguments.Count
	Select Case Count 
	Case 1 
		Flag = WScript.Arguments(0)
		If InStr(UCase(Flag),"ADD") > 0 Then 
			Const HKEY_CLASSES_ROOT = &H80000000
			strComputer = "."
			Set objRegistry = GetObject("winmgmts:\\" & _
			    strComputer & "\root\default:StdRegProv")
			Set objshell = CreateObject("Wscript.shell") 'Create "wscript.shell" object
			
			If KeyExists("HKEY_CLASSES_ROOT\*\shell\runas\") = False Then 
			objRegistry.CreateKey HKEY_CLASSES_ROOT, "*\shell\runas"   
			End If 
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "*\shell\runas", "", "take ownership"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "*\shell\runas", "Icon", "C:\Windows\System32\imageres.dll,-78"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "*\shell\runas", "NoWorkingDirectory", ""
			objRegistry.CreateKey HKEY_CLASSES_ROOT, "*\shell\runas\Command"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "*\shell\runas\Command", "", "cmd.exe /c takeown /f ""%1"" && icacls ""%1"" /grant administrators:F"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "*\shell\runas\Command", "IsolatedCommand", "cmd.exe /c takeown /f ""%1"" && icacls ""%1"" /grant administrators:F"
			
			If KeyExists("HKEY_CLASSES_ROOT\Directory\shell\runas\") = False Then 
			objRegistry.CreateKey HKEY_CLASSES_ROOT, "Directory\shell\runas"
			End If 
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "Directory\shell\runas", "", "Take Ownership"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "Directory\shell\runas", "Icon", "C:\Windows\System32\imageres.dll,-78"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "Directory\shell\runas", "NoWorkingDirectory", ""
			objRegistry.CreateKey HKEY_CLASSES_ROOT, "Directory\shell\runas\Command"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "Directory\shell\runas\Command", "", "cmd.exe /c takeown /f ""%1"" /r /d y && icacls ""%1"" /grant administrators:F /t"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "Directory\shell\runas\Command", "IsolatedCommand", "cmd.exe /c takeown /f ""%1"" /r /d y && icacls ""%1"" /grant administrators:F /t"	
			
			If KeyExists("HKEY_CLASSES_ROOT\dllfile\shell\") = False Then 
			objRegistry.CreateKey HKEY_CLASSES_ROOT, "dllfile\shell"
			End If 
			objRegistry.CreateKey HKEY_CLASSES_ROOT, "dllfile\shell\runas"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "dllfile\shell\runas", "", "Take Ownership"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "dllfile\shell\runas", "HasLUAShield", ""
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "dllfile\shell\runas", "NoWorkingDirectory", ""
			objRegistry.CreateKey HKEY_CLASSES_ROOT, "dllfile\shell\runas\Command"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "dllfile\shell\runas\Command", "", "cmd.exe /c takeown /f ""%1"" && icacls ""%1"" /grant administrators:F"
			objRegistry.SetStringValue HKEY_CLASSES_ROOT, "dllfile\shell\runas\Command", "IsolatedCommand", "cmd.exe /c takeown /f ""%1"" && icacls ""%1"" /grant administrators:F"
			
			WScript.Echo "Add 'Take Ownership' to context menu successfully."
		ElseIf InStr(UCase(Flag),"REMOVE")  > 0 Then 
			If KeyExists("HKEY_CLASSES_ROOT\*\shell\runas\") = True Then 
				DeleteSubkeys  HKEY_CLASSES_ROOT, "*\shell\runas"  
			End If 
			If KeyExists("HKEY_CLASSES_ROOT\Directory\shell\runas\") = True Then 
				DeleteSubkeys  HKEY_CLASSES_ROOT, "Directory\shell\runas"
			End If 	
			If KeyExists("HKEY_CLASSES_ROOT\dllfile\shell\") = True Then 
				DeleteSubkeys  HKEY_CLASSES_ROOT, "dllfile\shell"
			End If 
			WScript.Echo "Remove 'Take Ownership' from context menu successfully."
		Else 
			WScript.Echo "Invalid input value, please try again"
		End If 
	 Case Else 
	 	WScript.Echo "Invalid input value, please try again"
	 End Select 	
				
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