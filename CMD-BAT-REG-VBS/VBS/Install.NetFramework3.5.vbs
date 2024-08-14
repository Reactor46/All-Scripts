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

Sub main() 
	Dim sExecutable,SystemSet,objShell,objExecObject,System,caption
	sExecutable = LCase(Mid(Wscript.FullName, InstrRev(Wscript.FullName,"\")+1))
	If sExecutable <> "cscript.exe" Then 
	  WScript.Echo "Please run this script with cscript.exe"
	  Wscript.Quit
	End If
	Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
	Set objShell = WScript.CreateObject("WScript.Shell")
	for each System in SystemSet 
		caption = System.Caption 
	next 
	If InStr(caption,"Microsoft Windows 8") Then 
		If GetError = True Then 
			WScript.StdOut.WriteLine "Please run this script as Administrator"
		Else 
			If GetNetFramewordstatus = True  Then
				wscript.stdout.WriteLine ".Net Framework 3.5 has been installed and enabled."
			Else	
				Set objExecObject = objShell.Exec("Dism /online /Enable-feature /featurename:NetFx3 /All")
				WScript.StdOut.WriteLine "Installing .Net Framework 3.5 online,please wait.... "
				While objExecObject.Status = 0
					WScript.Sleep 1
				Wend
				If GetNetFramewordstatus = True Then 
					WScript.StdOut.WriteLine "Install .Net Framework 3.5 successfully."
				Else 
					WScript.StdOut.WriteLine "Failed to install .Net Framework 3.5 online. You can use local source to install it."
					WScript.StdOut.Write "Local source :"
					Source  = WScript.StdIn.ReadLine
					WScript.StdOut.WriteLine "Installing .Net Framework 3.5 in local,please wait.... "
					Set objExecObject = objShell.Exec("DISM /Online /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:" & Source )	
					While objExecObject.Status = 0
						WScript.Sleep 1
					Wend	
					If GetNetFramewordstatus = True Then 
						WScript.StdOut.WriteLine "Install .Net Framework 3.5 successfully."
					Else 
						WScript.StdOut.WriteLine "Failed to install .Net Framework 3.5,please make sure the local source is correct."
					End If 
				End If 
			End If 
		End If 
	Else 
		WScript.StdOut.WriteLine "Please run this script in Windows 8"
	End If 

End Sub 

Function GetNetFramewordstatus
	Dim objShell,objExecObject,Flag,strText,returnValue,Result
	Set objShell = WScript.CreateObject("WScript.Shell")
	Set objExecObject = objShell.Exec("Dism /online /Get-FeatureInfo /FeatureName:NetFx3")
	Flag = 0
	Do While Not objExecObject.StdOut.AtEndOfStream
	    strText = objExecObject.StdOut.ReadLine()
	    returnValue = InStr(strText,"Enabled")
		Flag = Flag + returnValue
	Loop
	If Flag > 0 Then 
		Result = True 
	Else 
		Result = False 
	End If 
	GetNetFramewordstatus = Result 
End Function 

Function GetError 
	Dim objShell,objExecObject,Flag,strText,returnValue,Result
	Set objShell = WScript.CreateObject("WScript.Shell")
	Set objExecObject = objShell.Exec("Dism /online /Get-FeatureInfo /FeatureName:NetFx3")
	Flag = 0
	Do While Not objExecObject.StdOut.AtEndOfStream
	    strText = objExecObject.StdOut.ReadLine()
	    returnValue = InStr(strText,"Error")
		Flag = Flag + returnValue
	Loop
	If Flag > 0 Then 
		Result = True 
	Else 
		Result = False 
	End If 
	GetError = Result 
End Function 

Call main 