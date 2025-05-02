Option Explicit
Dim objWshShell

Set objWshShell = WScript.CreateObject("WScript.Shell")

objWshShell.Run "uninstall_flash_player.exe /silent", 1, True

objWshShell.Run "msiexec /i install_flash_player_active_x.msi REBOOT=REALLYSUPPRESS /qn", 1, True

If ((FuzzyIsInstalled ("Mozilla")) OR _
(FuzzyIsInstalled ("Netscape"))) _
Then
objWshShell.Run "msiexec /i install_flash_player_plugin.msi REBOOT=REALLYSUPPRESS /qn", 1, True
End If


Function FuzzyIsInstalled(strfSearchDisplayName)

' *********************************************************
' Purpose: Tells whether a piece of software is installed.
' Rather than looking for an exact match, this version searches
' for installed software that contains the search string.
' So searching for Mozilla will return true if Mozilla Firefox
' Is installed. 
' Inputs: strfSearchDisplayName: the Display Name of the 
' software to be searched for.
' Returns: A boolean saying whether or not the software
' is installed
' *********************************************************

dim strfComputer, strfKeyPath, fSubKey, strfFoundDisplayName
dim intfRet
dim blnfFoundYet
dim arrfSubKeys
dim objfReg
const HKEY_LOCAL_MACHINE = &H80000002

strfComputer = "."
strfKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
blnfFoundYet = False
Set objfReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strfComputer & "\root\default:StdRegProv")
objfReg.EnumKey HKEY_LOCAL_MACHINE, strfKeyPath, arrfSubKeys

For Each fSubKey in arrfSubKeys
intfRet = objfReg.GetStringValue(HKEY_LOCAL_MACHINE, strfKeyPath & fSubKey, "DisplayName", strfFoundDisplayName) 
If intfRet = 0 Then
objfReg.GetStringValue HKEY_LOCAL_MACHINE, strfKeyPath & fSubkey, "DisplayName", strfFoundDisplayName
If (inStr (strfFoundDisplayName, strfSearchDisplayName) > 0) Then
blnfFoundYet = True
End If
End If
Next

FuzzyIsInstalled = blnfFoundYet
End Function