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

currentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
Set wshell = CreateObject("wscript.shell")
pspath = currentDirectory & "\takeown\takeown.ps1"

'wshell.Run  "powershell.exe -noexit "  & pspath, 0, True 
regpath = "HKEY_CLASSES_ROOT\CLSID\{8E74D236-7F35-4720-B138-1FED0B85EA75}\ShellFolder\Attributes"
returnvalue = wshell.RegRead(regpath)
If Err.Number =0 Then 
	Set objshell = CreateObject("shell.application")
	objshell.ShellExecute "cmd.exe", " /c Powershell.exe -noexit "  & pspath, "" ,"Runas", 0
	If returnvalue = -260016771 Then 
		wshell.RegWrite regpath, -1626865587, "REG_DWORD"
		choice =  MsgBox("You have removed the OneDrive from Windows explorer navigation, you need to restart computer to take effect.",1,"Restart Computer")
wscript.echo choice
		 If choice = 1  Then 
		 	wshell.run "cmd.exe /c shutdown -r"
		 End  If 
		  
	ElseIf returnvalue = -1626865587 Then 
		wshell.RegWrite regpath, -260016771, "REG_DWORD"
		choice =  MsgBox("You have added OneDrive from Windows explorer navigation, you need to restart computer to take effect.",1,"Restart Computer")
		wscript.echo choice
		 If choice = 1  Then 
		 	wshell.run "cmd.exe /c shutdown -r"
		 End If 
	End If 
Else 
	WScript.Echo "Not find OneDrive installed."
End If 