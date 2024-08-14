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
'This script is to change Windows or office product key and try to active it.

'Run script with administrator privilege
If WScript.Arguments.Count = 0 Then
	Dim objshell 
	Set objshell = CreateObject("Shell.Application")
	objshell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "Runas", 1
Else
	ChangeWinKey 
	WScript.StdOut.Write("Press enter to exist.")
   	WScript.StdIn.ReadLine()
End If 
'This function is to change Windows key and active it
Function ChangeWinKey
	Dim objshell, VBpath, SysRoot, Result, Flag, value, str
	'Create wscript.shell object 
	Set objshell = CreateObject("Wscript.shell")
	'Get system root path 
	SysRoot = objshell.ExpandEnvironmentStrings("%SystemRoot%")
	'Get the vbscript path 
	VBpath =  SysRoot & "\System32\slmgr.vbs"
	'Get Windows key 
	value = InputBox("Input Windows key:")
    If Len(value) <> 0 Then 
    	'Import Windows product key 
		Set Result = objshell.Exec("Cscript.exe " & VBPath & " -ipk " & value)
		Flag = False 
		Do While Not result.StdOut.AtEndOfStream 
			str = result.StdOut.ReadLine()
			If InStr(UCase(str),UCase("Error")) <> 0 Then 
				Flag = True 
			End If 
			WScript.Echo str
		Loop 
		'If importing failed, try again.
		If Flag = True  Then 
			WScript.Echo "Please try again."
			ChangeWinKey 			
		Else
		'Try to active Windows 
			WScript.Echo "Try activing Windows product."
			Set Result = objshell.Exec("Cscript.exe " & VBPath & " -ato")
			Flag = False 
			Do While Not result.StdOut.AtEndOfStream 
				str = result.StdOut.ReadLine()
				WScript.Echo str
			Loop 
		End If 
    Else 
    	WScript.Echo "Action cancelled by user."
    End If 
	Set objshell = Nothing 
	Set result = Nothing 
End Function 

