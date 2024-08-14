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
On Error Resume Next 
Err.Clear
Dim wshshell, path 
Set wshshell= CreateObject("wscript.shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoControlPanel"
If KeyExists(Path) Then  
	'enable control panel
	wshshell.RegDelete(path)
	If Err.Number = 0 Then 
		WScript.Echo "Control panel has been enabled."
		Call RefreshExplorer
	Else 
		WScript.Echo Err.Description
	End If 
	Err.Clear
Else
	'disable control panel
 	wshshell.RegWrite Path,1,"REG_DWORD"
 	If Err.Number = 0 Then 
	 	WScript.Echo "Control panel has been disabled."
	 	Call RefreshExplorer
 	Else 
		WScript.Echo Err.Description
	End If 
	Err.Clear
End If 

End Sub 

'################################################
'Check registry exists
'################################################
Function KeyExists(Path)
	On Error Resume Next 
	Dim objshell,Flag,value
	Set objShell = CreateObject("WScript.Shell")
	'read the registry key value
	value = objShell.RegRead(Path) 
	Flag = False 
	If Err.Number = 0 Then 	
	 	Flag = True 
	End If
	Keyexists = Flag
	Err.Clear
End Function 

Function RefreshExplorer()
'################################################
'Refresh explorer
'################################################
	dim strComputer, objWMIService, colProcess, objProcess 
	strComputer = "."
	'Get WMI object 
	Set objWMIService = GetObject("winmgmts:" _
	  & "{impersonationLevel=impersonate}!\\" _ 
	  & strComputer & "\root\cimv2") 
	Set colProcess = objWMIService.ExecQuery _
	  ("Select * from Win32_Process Where Name = 'explorer.exe'")
	For Each objProcess in colProcess
	   objProcess.Terminate()
	Next 

End Function 

Call main 