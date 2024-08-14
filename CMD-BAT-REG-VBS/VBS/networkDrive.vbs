
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


If WScript.Arguments.Count = 0 Then
	Dim objshell 
	Set objshell = CreateObject("Shell.Application")
	objshell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "open", 1
Else
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_MappedLogicalDisk",,48)
	Set oShell = CreateObject("Shell.Application")
	For Each objItem in colItems 
	    sDrive =  objItem.Name 
	    oShell.NameSpace(sDrive).Self.Name = "MappedDrive"
	    WScript.StdOut.WriteLine "Hidden " & objItem.Name & "(" & objItem.ProviderName & ") successfully"
	Next
	
	WScript.StdOut.Write("Press ENTER to exist...")
	qiut = WScript.StdIn.ReadLine()
	WScript.Quit
End If 