' =============================================================================
' VBScript --- Determine If Your System is 32 Bit Or 64 Bit
' The Reassuring Way of Checking the Processor Details and Architecture
' Usage --- CScript /nologo ProcOSArch.vbs OR CScript ProcOSArch.vbs
' =============================================================================

Option Explicit

Dim ObjWMI, ColSettings, ObjProcessor
Dim StrComputer, ObjNetwork

Set ObjNetwork = WScript.CreateObject("WScript.Network")
StrComputer = Trim(ObjNetwork.ComputerName)
Set ObjNetwork = Nothing
WScript.Echo VbCrLf & "Computer Name: " & StrComputer
WScript.Echo vbNullString
Set ObjWMI = GetObject("WINMGMTS:" & "{ImpersonationLevel=Impersonate,AuthenticationLevel=Pkt}!\\" & StrComputer & "\Root\CIMV2")
Set ColSettings = ObjWMI.ExecQuery ("SELECT * FROM Win32_Processor")
For Each ObjProcessor In ColSettings
	Select Case ObjProcessor.Architecture
		Case 0
			WScript.Echo "Processor Architecture Used by the Platform: x86"
		Case 6
			WScript.Echo "Processor Architecture Used by the Platform: Itanium-Based System"
		Case 9
			WScript.Echo "Processor Architecture Used by the Platform: x64"
	End Select
	Select Case ObjProcessor.ProcessorType
		Case 1
			WScript.Echo "Processor Type: Other. Not in the Known List"	
		Case 2
			WScript.Echo "Processor Type: Unknown Type"
		Case 3
			WScript.Echo "Processor Type: Central Processor (CPU)"
		Case 4
			WScript.Echo "Processor Type: Math Processor"
		Case 5
			WScript.Echo "Processor Type: DSP Processor"
		Case 6
			WScript.Echo "Processor Type: Video Processor"
	End Select
	WScript.Echo "Processor: " & ObjProcessor.DataWidth & "-Bit"
	WScript.Echo "Operating System: " & ObjProcessor.AddressWidth & "-Bit"
	WScript.Echo vbNullString	
	If ObjProcessor.Architecture = 0 AND ObjProcessor.AddressWidth = 32 Then
		WScript.Echo "This Machine has 32 Bit Processor and Running 32 Bit OS"
	End If
	If (ObjProcessor.Architecture = 6 OR ObjProcessor.Architecture = 9) AND ObjProcessor.DataWidth = 64 AND ObjProcessor.AddressWidth = 32 Then
		WScript.Echo "This Machine has 64-Bit Processor and Running 32-Bit OS"
	End If
	If (ObjProcessor.Architecture = 6 OR ObjProcessor.Architecture = 9) AND ObjProcessor.DataWidth = 64 AND ObjProcessor.AddressWidth = 64 Then
		WScript.Echo "This Machine has 64-Bit Processor and Running 64-Bit OS"
	End If
Next
Set ObjProcessor = Nothing:	Set ColSettings = Nothing:	Set ObjWMI = Nothing:	StrComputer = vbNullstring