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
'################################################
' The starting point of execution for this script.
'################################################
Sub Main()
	On Error Resume Next 
	Dim arrayNetCards , objShell
	Dim netcard , AdapterDeviceNumber
	Dim Path , value , intAnswer
	Set arrayNetCards =  GetObject( _
		"Winmgmts:").ExecQuery("Select * From Win32_NetworkAdapter Where   PhysicalAdapter = True and  Manufacturer <> 'Microsoft' and ConfigManagerErrorCode = 0 and  ConfigManagerErrorCode <> 22")
	Set objShell = CreateObject("WScript.Shell")		'Create wscript.shell object
	If arrayNetCards.count <> 0 Then 
		For Each netcard In arrayNetCards
			If  Instr(netcard.PNPDeviceID,"ROOT") =0 Then    
				If netcard.DeviceID  < 10 Then 
					AdapterDeviceNumber = "000"+netcard.DeviceID
				Else
					AdapterDeviceNumber = "00"+netcard.DeviceID
				End If 					
			End If 
			Path = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}\"&AdapterDeviceNumber&"\PnPCapabilities"
			value = objShell.RegRead(Path)  'Read the registry
			If  Err = 0  Then 
				If value = 24 Then 'Verify if the model is enable
					wscript.echo  "The 'Allow the computer to turn off this device to save power' has been disabled." 
				Else
					objShell.RegWrite Path, 24, "REG_DWORD"  'Modify the key ,and disable the model
					If Err = 0 Then 
						intAnswer = _
							 MsgBox("The 'Allow the computer to turn off this device to save power' has been disabled. It will take effect after reboot, do you want to reboot right now?", _
							 vbYesNo, "Reboot right now")
						If intAnswer = vbYes Then
							objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 0"
						End If
					Else 
						wscript.echo  "Operation failed.Ensure that you have the administrator permission"
					End If 
				End If 
			Else 
				wscript.echo  "Operation failed.Ensure that you have the administrator permission"
			End If 
		Next
	Else 
		wscript.echo  "Please ensure your network adapter work normally!"
	End If 
End Sub 

Call Main

