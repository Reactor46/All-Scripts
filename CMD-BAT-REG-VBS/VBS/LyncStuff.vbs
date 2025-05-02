Option Explicit

Const HKEY_CURRENT_USER   = &H80000001
Dim StrKeyPath, StrValueName, StrPC, ObjRegistry, ObjLyncFSO
Dim GetThisValue, ObjNet, ObjProcess, ObjItem, ObjShell

On Error Resume Next

Set ObjNet = WScript.CreateObject("WScript.Network")
StrPC = Trim(ObjNet.ComputerName)
Set ObjNet = Nothing

Set ObjProcess = GetObject("WinMgmts:")
For Each ObjItem In ObjProcess.InstancesOf("Win32_Process")
	If Strcomp(Trim(ObjItem.Name), "Communicator.exe", vbTextCompare) = 0 OR Strcomp(Trim(ObjItem.Name), "Lync.exe", vbTextCompare) = 0 Then
		ObjItem.Terminate()
		If Err.Number <> 0 Then
			Err.Clear
		End If
	End If
Next
Set ObjProcess = Nothing

Set ObjLyncFSO = CreateObject("Scripting.FileSystemObject")
If ObjLyncFSO.FileExists("C:\Program Files (x86)\Microsoft Office\Office15\Lync.exe") = True Then
	StrKeyPath = "Software\Microsoft\Office\15.0\Lync"
Else
	StrKeyPath = "Software\Microsoft\Communicator"
End If
Set ObjLyncFSO = Nothing

StrValueName = "AutoRunWhenLogonToWindows"

On Error Resume Next

Set ObjRegistry = GetObject("WinMgmts:\\" & StrPC & "\Root\Default:StdRegProv")
ObjRegistry.GetDWORDValue HKEY_CURRENT_USER, StrKeyPath, StrValueName, GetThisValue
If Err.Number <> 0 Then
	Err.Clear
End If
If Trim(GetThisValue) <> 0 Then
	ObjRegistry.SetDWORDValue HKEY_CURRENT_USER, StrKeyPath, StrValueName, 0
	If Err.Number <> 0 Then
		Err.Clear
	End If
End If
Set ObjRegistry = Nothing

Set ObjShell = CreateObject("WScript.Shell")
ObjShell.Run "LyncMatter.cmd", 0, True
Set ObjShell = Nothing

WScript.Quit