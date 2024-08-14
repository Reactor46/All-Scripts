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
'Check if the script run with administrator privilege
If WScript.Arguments.Count = 0 Then
	Dim objshell 
	Set objshell = CreateObject("Shell.Application")
	objshell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "Runas", 1
Else
	Dim wshell, UserProfile, ListPath
	'Get the shared folders
	WScript.StdOut.Write("Choose shared folders:")
	InputStr = WScript.StdIn.ReadLine
	'Create wscript.shell object 
	Set wshell = CreateObject("wscript.shell")
	'Get the user profile path
	UserProfile = wshell.ExpandEnvironmentStrings("%UserProfile%")
	ListPath = UserProfile & "\list.txt"
	Set FSObject = CreateObject("Scripting.FileSystemObject")
	'Create list text 
	Set ListTxt = FSObject.CreateTextFile(ListPath, True, False)
	SharedFolders = Split(InputStr, ",")
	'Write the shared folders string to list text file 
	for each x in SharedFolders
	    Folder = Replace(x,"""","")
	    ListTxt.writeline(Folder)
	Next
 	Listtxt.Close
 	'Create batch file 
 	CMDpath = UserProfile & "\setPerm.cmd"
 	CMDStr = "FOR /F %%I IN (" & ListPath & ") DO icacls " &  """%%I""" & " /C /Q /T /grant HomeUsers:(OI)(CI)(F)"
 	Set CMDFile = FSObject.CreateTextFile(CMDpath, True, False)
 	'write string to batch file 
 	CMDFile.WriteLine(CMDStr)
 	CMDFile.Close
 	Dim objshell 
 	Set objShell = WScript.CreateObject("WScript.Shell")
 	'Create a schedule task
 	objShell.Run "SCHTASKS /Create /SC ONLOGON /TN setPermTask /TR " & CMDpath & " /RU System", 1, True 
	WScript.StdOut.WriteLine "Grant HomeGroup permission successfully."
	WScript.StdOut.WriteLine "Script will exist in five seconds."
	WScript.Sleep 5000
	Set FSObject = Nothing 
	Set listTxt = Nothing 
	Set wshell = Nothing 
End If 


