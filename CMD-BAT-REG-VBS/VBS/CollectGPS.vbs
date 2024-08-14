'--------------------------------------------------------------------------------- 
'The sample scripts are not supported under any Microsoft standard support 
'program or service. The sample scripts are provided AS IS without warranty  
'of any kind. Microsoft further disclaims all implied warranties including,  
'without limitation, any implied warranties of merchantability or of fitness for 
'a particular purpose. The entire risk arising out of the use or performance of  
'the sample scripts and documentation remains with you. In no event shall 
'Microsoft, its authors, or anyone else involved in the creation, production, or 
'delivery of the scripts be liable for any damages whatsoever (including, 
'without limitation, damages for loss of business profits, business interruption, 
'loss of business information, or other pecuniary loss) arising out of the use 
'of or inability to use the sample scripts or documentation, even if Microsoft 
'has been advised of the possibility of such damages 
'--------------------------------------------------------------------------------- 
'Run script with administrator privilege
If WScript.Arguments.Count = 0 Then
	Dim objshell 
	Set objshell = CreateObject("Shell.Application")
	objshell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "Runas", 1
Else 
	Const HKLM = &H80000002
	Set objshell = CreateObject("wscript.shell")
	strCom = "."
	CurrentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
	ResultFile = CurrentDirectory & "result.html"
	
	objshell.Run "cmd /c gpresult /h " & ResultFile, 0, True 
	Set objReg =  GetObject("winmgmts:\\" & strCom & "\root\default:StdRegProv")
	strKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Diagnostics"
	
	objReg.CreateKey HKLM, strkey
	objReg.SetDWORDValue HKLM, strKey, "GPSvcDebugLevel", 196610
	
	Set FileObject = CreateObject("Scripting.FileSystemObject")
	Windir = objshell.ExpandEnvironmentStrings("%windir%")
	Usermode = Windir & "\debug\usermode"
	If  FileObject.FolderExists(UserMode) = False  Then 
		FileObject.CreateFolder(UserMode)
	End If 
	GPSVG = usermode & "\gpsvc.log"
	objshell.Run  "cmd /c gpupdate /force", 0, True 
	If FileObject.FileExists(GPSVG) Then 
		FileObject.CopyFile  GPSVG, CurrentDirectory 
	End If 
	WScript.Echo "Generate Result.html and gpsvg.log file successfully."
	WScript.Echo "Press Enter to exist..."
	WScript.StdIn.ReadLine
End If 