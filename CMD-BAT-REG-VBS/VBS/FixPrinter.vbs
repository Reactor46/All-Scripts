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
Option Explicit

' ################################################
' The starting point of execution for this script.
' ################################################

Sub main()
	Dim sExecutable
	sExecutable = LCase(Mid(Wscript.FullName, InstrRev(Wscript.FullName,"\")+1))
	If sExecutable <> "cscript.exe" Then
		wscript.echo "Please run this script with cscript.exe"
	Else 
	    Call FixPrinter 	
	End If
End Sub 

Function FixPrinter
'	On Error Resume Next 
	Err.Clear
	Dim objshell,objWMIService,colListOfServices,objService
	Set objshell = CreateObject("wscript.shell")
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	'Stop Service
	WScript.StdOut.Writeline("Stoping 'spooler' service")
	Set colListOfServices = objWMIService.ExecQuery("Select * from Win32_Service Where Name ='spooler'")
	For Each objService in colListOfServices
	    objService.StopService()
	Next
	If Err.Number <> 0 Then 
		WScript.StdOut.WriteLine "Failed to stop service 'spooler'.Script exits."
		WScript.Quit -1 
	End If 
	'Delete files in the folder
	Dim Windir,objFSO
	Const DeleteReadOnly = True
	Windir = objshell.ExpandEnvironmentStrings("%windir%")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	WScript.StdOut.Writeline("Deleting files in foler '" & Windir & "\system32\spool\printers\'")
	objshell.Run("CMD /c del " & Windir & "\system32\spool\printers\*.* /s /f /q")	
	WScript.StdOut.Writeline("Deleting files in foler '" & Windir & "\system32\spool\drivers\w32x86\'")	
	objshell.Run("CMD /c del " & Windir & "\system32\spool\drivers\w32x86\\*.* /s /f /q")	
	If Err.Number <> 0 Then 
		WScript.StdOut.WriteLine "Failed to delete files in spool foler.Script exits."
		WScript.Quit -1 
	End If 
	'Start the service 'spooler'
	WScript.StdOut.Writeline("Starting 'spooler' service")
	Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='spooler'")
	For Each objService in colListOfServices
	    objService.StartService()
	Next
	If Err.Number <> 0 Then 
		WScript.StdOut.WriteLine "Failed to start service 'spooler'.Script exits."
		WScript.Quit -1 
	End If 
	'Back up the registry keys in the current folder
	Dim CurrentDirectory,RegPath1,RegPath2,Key1,Key2
	CurrentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
	RegPath1 = CurrentDirectory & "\NTx86.reg"
	Key1 = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows NT x86"
	RegPath2 = CurrentDirectory & "\Monitors.reg"
	Key2 = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Monitors"
	WScript.StdOut.WriteLine("Back up registry...")
	objshell.Exec "regedit /e /s """ & RegPath1 & """ """ & Key1 & """"
	objshell.Exec "regedit /e /s """ & RegPath2 & """ """ & Key2 & """"
	If Err.Number <> 0 Then 
		WScript.StdOut.WriteLine "Failed to backup regisry keys.Script exits."
		WScript.Quit -1 
	End If 
	'Delete the registry keys
	Dim strComputer,NTx86Path,objRegistry,strSubkey,MonitorPath,NTSubkeys,MonitorSubkeys
	Const HKEY_LOCAL_MACHINE = &H80000002 
	strComputer = "."
	NTx86Path = "SYSTEM\CurrentControlSet\Control\Print\Environments\Windows NT x86" 
	Set objRegistry = GetObject("winmgmts:\\" & _
	strComputer & "\root\default:StdRegProv") 
	WScript.StdOut.WriteLine("Deleting registry...")
	objRegistry.EnumKey HKEY_LOCAL_MACHINE, NTx86Path, NTSubkeys 
    For Each strSubkey In NTSubkeys
    	If InStr(UCase(strSubkey),UCase("Print Processors")) = 0  Then 
    		objRegistry.DeleteKey HKEY_LOCAL_MACHINE, NTx86Path & "\" & strSubkey
    	End If 
    Next 
	MonitorPath = "SYSTEM\CurrentControlSet\Control\Print\Monitors"
	objRegistry.EnumKey HKEY_LOCAL_MACHINE, MonitorPath, MonitorSubkeys
	For Each strSubkey In MonitorSubkeys
		If CheckStr(strSubkey) = False   Then 
			objRegistry.DeleteKey HKEY_LOCAL_MACHINE, MonitorPath & "\" & strSubkey
		End If 
	Next 
	If Err.Number <> 0 Then 
		WScript.StdOut.WriteLine "Failed to delete registry keys.Script exits."
		WScript.Quit -1 
	End If 
	WScript.StdOut.WriteLine("Action done.")
End Function 
	
'This function is to check if the registry key need to be deleted	
Function CheckStr(strSubkey)
	Dim FLag,StrArrs,str
	Flag = False 
	StrArrs= Array("Local Port","Microsoft Document Imaging Writer Monitor",_
				   "Microsoft Shared Fax Monitor","Standard TCP/IP Port","USB Monitor","WSD Port")
	For Each str In StrArrs 
		If InStr(UCase(strSubkey),UCase(str)) = 1 Then 
			Flag = True 
		End If 
	Next 
	CheckStr =  Flag 
End Function 	

Call  main



