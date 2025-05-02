On Error Resume Next
Dim objFSO,objFILE,objShell,objNetwork
Dim Shell, WshShell
Set WshShell = CreateObject( "WScript.Shell" )
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")
Set objNetwork = CreateObject("Wscript.Network")

	strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colAdapters = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled=True")

For Each objAdapter In colAdapters
	For Each strAddress In objAdapter.IPAddress
		arrOctets = Split(strAddress, ".")
		If arrOctets(0) <> "" Then
		strSubnet = arrOctets(0) & "." & arrOctets(1) & "." &arrOctets(2)
		x = 1
		Exit For
	End If
		If x = 1 Then
			Exit For
		End If
	Next
Next
	Select Case arrOctets(2)

	Case "96" 
	objNetwork.MapNetworkDrive "X:", "\\BRANCH1DC\ais_data"
	objNetwork.MapNetworkDrive "S:", "\\BRANCH2DC\SHARED"

	Case "94" 
	objNetwork.MapNetworkDrive "X:", "\\BRANCH2DC\ais_data"
	objNetwork.MapNetworkDrive "S:", "\\BRANCH2DC\SHARED"
		
	Case "95" 
	objNetwork.MapNetworkDrive "X:", "\\BRANCH3DC\ais_data"
	objNetwork.MapNetworkDrive "S:", "\\BRANCH2DC\SHARED"
		
	Case "109"
	objNetwork.MapNetworkDrive "X:", "\\branch4dc\ais_data"
	objNetwork.MapNetworkDrive "S:", "\\branch2dc\shared"
	

End Select


mDrive = "X:\"
Set oShell = CreateObject("Shell.Application")
oShell.NameSpace(mDrive).Self.Name = "AIS"

mDrive = "P:\"
Set oShell = CreateObject("Shell.Application")
oShell.NameSpace(mDrive).Self.Name = "PERSONAL"

mDrive = "S:\"
Set oShell = CreateObject("Shell.Application")
oShell.NameSpace(mDrive).Self.Name = "S:DRIVE"