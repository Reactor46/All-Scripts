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
	Dim objSh
	Set objSh = CreateObject("Shell.Application")
	objSh.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "runas", 1
Else

	'Get operating system version 
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
	For Each os in oss
	    Caption  = os.Caption
	Next
	'verify if office is instaled
	Set colSoft = objWMIService.ExecQuery("SELECT * FROM Win32_Product WHERE Name Like 'Microsoft Office%'")
    If colSoft.Count = 0 Then
      wscript.echo "NO OFFFICE INSTALLED" 
    else
       For Each objItem In colSoft
           version = Left(objItem.Version, InStr(1,objItem.Version,".")-1)
          exit for
       Next
    End If
    WScript.Echo "Try to repair registry key..."
	'verify Office version
	Set objshell = CreateObject("wscript.shell")
	If version  = 12 Then 
		objshell.Run "reg add HKCU\Software\Microsoft\Office\12.0\Word\Options /v NoReReg /t REG_DWORD /d 1",1, True 
	ElseIf version = 14 Then 
		objshell.Run "reg add HKCU\Software\Microsoft\Office\14.0\Word\Options /v NoReReg /t REG_DWORD /d 1",1, True 
	ElseIf version = 15 Then 
		objshell.Run "reg add HKCU\Software\Microsoft\Office\15.0\Word\Options /v NoReReg /t REG_DWORD /d 1",1, True
	Else 
		WScript.Echo "Not support office version"
		WScript.StdOut.Write "Press enter to exit"
		WScript.StdIn.ReadLine()
		WScript.Qui
	End If  
	
	WScript.Echo "Try to repair configuration..."
	If InStr(UCase(Caption),UCase("XP")) Then 
		objshell.Run "secedit /configure /cfg %windir%\repair\secsetup.inf /db secsetup.sdb /verbose",1, True 
	Else 
	    objshell.Run "secedit /configure /cfg %windir%\inf\defltbase.inf /db defltbase.sdb /verbose",1, True 
	End If 
	WScript.Echo "Operation done. If the problem continues to occur, we recommend that you uninstall Office and then install it."
	WScript.StdOut.Write "Press enter to exit"
	WScript.StdIn.ReadLine()

End If 