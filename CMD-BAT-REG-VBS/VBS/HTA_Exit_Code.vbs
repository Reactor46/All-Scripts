On Error Resume Next
Dim Shell
Dim strKeyPath
Dim strValueName
' Configure these options if you don't want my website address spammed in your registry
strKeyPath = "HKCU\SOFTWARE\thelazysa.com"
strValueName = "HTAReturnCode"
Set Shell = CreateObject("WScript.Shell")
Return = Shell.Run(".\HTA_Exit_Code.hta",0,True)
ExitCode = Shell.RegRead (strKeyPath & "\" & strValueName)
Shell.RegDelete (strKeyPath & "\")
Return = Msgbox("The exit code stored in the registry by the HTA was: " &_
	ExitCode & vbCrLf, vbOk, "HTA_Exit_Code.vbs")
Wscript.Quit ExitCode