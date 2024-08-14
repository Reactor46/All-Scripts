
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

On Error Resume Next 
'Check if the script run with administrator privilege
If WScript.Arguments.Count = 0 Then
	Set objshell = CreateObject("Shell.Application")
	objshell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "Runas", 1
Else
	Const HKEY_LOCAL_MACHINE = &H80000002
	Public count 
	Public oReg
	count = 0
	strComputer = "."
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Wow6432Node\Policies\Microsoft\Internet Explorer\BrowserEmulation\PolicyList"
	oReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
	'Check if the compatibility view list exists 
	If IsArray(arrSubKeys) Then 
		GetChoice 
	Else 
		AddUrlToList 
	End If 
	'This function is to confirm user selection 
	Function GetChoice 
		WScript.StdOut.Write "There are some website in compability view list, do you want to add more or remove?(A:Add, R:Remove, Q:Quit):"
		Choice = WScript.StdIn.ReadLine
		If UCase(choice) = "A" Then 
			'Add Uri to view list 
			AddUrlToList  
		ElseIf UCase(Choice) = "R" Then 
			'Delete the view list 
			RemoveList
		ElseIf UCase(choice) = "Q" Then 
			'quit the script
			WScript.Quit
		Else 
			WScript.StdOut.WriteLine "Invalid input please try again."
			GetChoice 
		End If  
	End Function 
	'This script is to add Uri to compatibility view list 
	Function AddUrlToList 
		strComputer = "."
		Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
		strKeyPath = "SOFTWARE\Wow6432Node\Policies\Microsoft\Internet Explorer\BrowserEmulation\PolicyList\"
		oReg.CreateKey HKEY_LOCAL_MACHINE,"SOFTWARE\Wow6432Node\Policies\Microsoft\Internet Explorer"
		oReg.CreateKey HKEY_LOCAL_MACHINE,"SOFTWARE\Wow6432Node\Policies\Microsoft\Internet Explorer\BrowserEmulation"
		oReg.CreateKey HKEY_LOCAL_MACHINE,strKeyPath
		Dim FSObject, wshell 
		Set FSObject = CreateObject("Scripting.FileSystemObject")
		Set wshell = CreateObject("wscript.shell")
		userprofile = wshell.ExpandEnvironmentStrings("%userprofile%")
		Set Folders =FSObject.GetFolder(userprofile & "\Favorites") 
		GetFile Folders 
		WScript.StdOut.WriteLine "Add " & count & " to list succesffully."
		WScript.StdOut.WriteLine  "Script will exist in five seconds."
		WScript.Sleep  5000
		
		Set oReg = Nothing 
		Set FSObject = Nothing 
		Set wshell = Nothing 
	End Function 
	
		Function GetFile(Folders)
		For Each subFolder In Folders.SubFolders
			GetFile(subfolder)
		Next 
		For Each File  In Folders.Files
			GetURL File
		Next  
	End Function 
	
	Function GetURL(File)
		Set FileObj = File.OpenAsTextStream
		Do While Fileobj.AtEndOfStream <> True 
			Line = fileobj.ReadLine
			If InStr(UCase(Line),UCase("Url=")) =1 Then 	
			    Arr = Split(line,"/") 
			    strValueName = GetTLD(Arr(2))
			    If ValueExists(strValueName) = False  Then 	
			    	oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValueName
			    	Count = count +1
			    End If 
			   	
			End If 
		Loop 
	    Fileobj.Close 
	End Function 
	
	Function GetTLD(Str)
		If InStr( UCase(Str), ucase("WWW")) = 1 Then
	     	Str =  Right(Str,Len(Str)-InStr(Str,".")) 	
	   	End If 
	    If  InStr(str,".") <> InStrRev(str,".") Then 
	    	If (Len(str)-InStrRev(str, ".")) < 3 Then 
	    		Do Until  InStr(str,".") = InStrRev(str,".")
	    			Temp = str 
	    			Str =  Right(Str,Len(Str)-InStr(Str,"."))
					
	    		Loop 
	    		Str  =  Temp 
	    	Else 
	    		Do Until  InStr(str,".") = InStrRev(str,".")
	    			Str =  Right(Str,Len(Str)-InStr(Str,"."))		
	    		Loop 
	    	End If 	
	    End If  
	    GetTLD = str 
	End Function 
	
	Function ValueExists(ValueName)
		On Error Resume Next 
		Dim objshell,Flag,value
		Set objShell = CreateObject("WScript.Shell")
		Path = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Policies\Microsoft\Internet Explorer\BrowserEmulation\PolicyList\" & valueName 
		value = objShell.RegRead(Path) 
		Flag = False 
		If Err.Number = 0 Then 	
		 	Flag = True 
		End If
		ValueExists = Flag
		Set objshell = Nothing 
	End Function 
	
	Function RemoveList
		strComputer = "."
		Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
		strKeyPath = "SOFTWARE\Wow6432Node\Policies\Microsoft\Internet Explorer\BrowserEmulation\PolicyList\"
		oReg.DeleteKey HKEY_LOCAL_MACHINE, strKeyPath
		WScript.StdOut.WriteLine "Remove list successfully."
		WScript.StdOut.WriteLine "Script will exist in five seconds."
		WScript.Sleep 5000 
		Set oReg = Nothing 
	End Function 
End if 