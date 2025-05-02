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
	Set objshell = CreateObject("Shell.Application")
	objshell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "runas", 1
Else

Dim strFile
strFile = SelectFile( )
If strFile =  "" Then 
    WScript.Echo "No file selected.Script execution cancelled by user."
    WScript.Sleep 9000
Else
	WScript.Echo "Select 'install.wim' file"
	Dim wshobject
	Set wshobject = CreateObject("wscript.shell")
	Dim FSObject 
	Set FSObject = CreateObject("Scripting.FileSystemobject")
	WScript.Echo """" & strFile & """"
	If FSObject.FolderExists("C:\Refresh") = False Then 
		FSObject.CreateFolder("C:\Refresh")
	End If
	WScript.Echo "Copying 'install.wim' file to 'C:\Refresh'..." 
	If FSObject.FileExists("C:\Refresh\install.wim") = False Then 
		FSObject.CopyFile strFile, "C:\Refresh\" 
	End If 
	count =  0
    set wshobject = createobject("wscript.shell")
	Set objExecObject = wshobject.Exec("cmd /c dism /get-wiminfo /wimFile:" & strFile)
	Dim count,returnValue
	Do While Not objExecObject.StdOut.AtEndOfStream
		line = objExecObject.StdOut.ReadLine()
		returnValue = InStr(line,"Index :")
		count = count + returnValue
	    strText = strText & vbNewLine & line
	Loop
	WScript.Echo strText
	WScript.stdout.Write "Input the index of your OS version(No bigger than" & count & "):"
	If index < count Then 
		index = WScript.StdIn.ReadLine()
		wshobject.Run "cmd /c " & """reagentc /setosimage /path C:\Refresh /target C:\Windows /index " & index & """"
		WScript.Echo "The operation completed successfully.Script will exist in 5 seconds."
		WScript.Sleep 5000
	Else 
		WScript.Echo "Illegal input, script exists."
		WScript.Sleep 3000
	End If 
End If 


Function SelectFile()
    Dim objExec, strMSHTA, wshShell
    SelectFile = ""
    strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
             & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
    Set wshShell = CreateObject("WScript.Shell")
    Set objExec = wshShell.Exec(strMSHTA)
    SelectFile = objExec.StdOut.ReadLine()
    Set objExec = Nothing
    Set wshShell = Nothing
End Function
End If 
