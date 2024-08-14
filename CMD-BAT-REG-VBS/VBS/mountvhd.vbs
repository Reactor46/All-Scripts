
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
If WScript.Arguments.Count = 0 Then
	Dim objShell
	Set objShell = CreateObject("Shell.Application")
	objShell.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "runas", 1
Else
	Dim strFile	
	strFile = SelectFile( )
	If strFile = "" Then 
	    WScript.Echo "No file selected."
	Else
		
		Set fileobject = CreateObject("Scripting.FileSystemObject")
		Set objshell = CreateObject("wscript.shell")
		'get user profile path
		userfolder = objshell.ExpandEnvironmentStrings("%userprofile%")
		'create diskpart configuration file 
		txtfile = userfolder & "\mountVHD.txt"
		Set file = fileobject.CreateTextFile(txtfile,True)
		'write content into file
		file.WriteLine("select vdisk file= """ & strFile & """")
		file.WriteLine("attach vdisk")
		file.Close()
		'create execute command
		command = "schtasks /create /tn ""MountVHD"" /tr ""diskpart.exe /s ""C:\Users\Administrator\Desktop\cmd.txt"""" /sc ONLOGON"
		objshell.Run command,1, True
		WScript.Echo "Operation executed successfully. The specified VHD file will be mounted next logon."
	End If
	
	
	Function SelectFile( )
	    
	    Dim objExec, strMSHTA, wshShell
	
	    SelectFile = ""
	
	    ' For use in HTAs as well as "plain" VBScript:
	    strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
	             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
	             & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
	    Set wshShell = CreateObject( "WScript.Shell" )
	    Set objExec = wshShell.Exec( strMSHTA )
	    SelectFile = objExec.StdOut.ReadLine( )
	    Set objExec = Nothing
	    Set wshShell = Nothing
	End Function

End If 