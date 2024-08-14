Sub main()
Set sh = CreateObject("WScript.Shell")
theKey = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden"
setHidden = sh.RegRead(theKey)
If setHidden = 1 Then
	setHidden = 0
Else
	setHidden = 1
End If
sh.RegWrite theKey,setHidden,"REG_DWORD"
Call RefreshExplorer
If setHidden = 0 Then
	msgbox "Hide hidden files successfully "
Else
	msgbox "Show hidden files successfully"
End If
Set sh = Nothing
End Sub 
		
Function RefreshExplorer()
	dim strComputer, objWMIService, colProcess, objProcess 
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
	  & "{impersonationLevel=impersonate}!\\" _ 
	  & strComputer & "\root\cimv2") 
	Set colProcess = objWMIService.ExecQuery _
	  ("Select * from Win32_Process Where Name = 'explorer.exe'")
	For Each objProcess in colProcess
	   objProcess.Terminate()
	Next 
End Function 



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

Call main

