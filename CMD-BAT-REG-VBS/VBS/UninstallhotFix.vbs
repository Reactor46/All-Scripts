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

Sub main()
    Dim objshell,objExecObject,count,HotFixID
	If Wscript.Arguments.Count = 0 Then
	   	WScript.Echo "Invalid Arguments "
	Else
	    Dim arrHotFix()
	    For i = 0 to Wscript.Arguments.Count - 1
	        Redim Preserve arrHotFix(i)
	        arrHotFix(i) = Wscript.Arguments(i)
	    Next
	    
	    For Each HotFixId In arrHotFix
		 	UninstallHotFix(HotFixId)
		Next 
	End If
End Sub 

Function UninstallHotFix(HotFixID)
	ID = Right(HotfixID,Len(hotfixid)-2)
	If  GetHotFix(HotFixID) = True Then 
		Set objshell = CreateObject("wscript.shell")
		Set objExecObject = objshell.Exec("Cmd /c wusa.exe /uninstall /KB:" & ID & " /quiet /norestart")
		Do
		 	WScript.Sleep 3000
		Loop Until FindWusa = False 
		If GetHotFix(HotFixID) Then 
			WScript.Echo "Failed to uninstall " & HotFixID
		Else 
			WScript.Echo "Uninstall " & HotFixID & " Successfully."
		End If 
	Else 
		WScript.Echo "Not find hotfix " & HotFixID
	End If 
End Function 

Function FindWusa 'This function is to verify "wusa.exe" is running
	set service = GetObject ("winmgmts:")
	for each Process in Service.InstancesOf ("Win32_Process")
		If Process.Name = "wusa.exe" then
			Flag = True 
			Exit For 
		End If
	Next
	If Flag  Then 
		FindWusa = True 
	Else 
		FindWusa = False 
	End If 
End Function 

Function GetHotFix(HotFixID)
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
	Set colItems = objWMIService.ExecQuery("Select * from Win32_QuickFixEngineering where HotFixID = '" & HotFixID & "'")
	For Each colitem In colitems 
		If InStr(UCase(colitem.HotFixId),UCase(HotFixID)) > 0 Then 
			Flag = True 
		End If 
	Next 
	If 	Flag = True  Then 
		GetHotFix = True 
	Else
		GetHotFix = False 
	End If 
End Function 

Call main 
