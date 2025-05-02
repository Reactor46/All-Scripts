'''''BitLocker Status Script. By Oliver Ford''''''
'''''Detects if Bitlocker is active. If so, suspends BitLocker and creates a Scheduled Task for the script to reenable BitLocker on reboot'''''
'Script location. Set the below to the location of the current script. This is so that the script can be run again on reboot
Dim scriptLoc
scriptLoc = "C:\bitlockerstatus.vbs"

'Disable error messages while getting the WMI class in case the OS doesn't support bitlocker
On Error Resume Next
Dim objWMIService, colEncrypt, objShell

'Get drive encryption object
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2\Security\MicrosoftVolumeEncryption")
Set colEncrypt = objWMIService.ExecQuery("SELECT * FROM Win32_EncryptableVolume",,48)

'Quit script if running on a Windows version (like XP or 7 Professional) which doesn't support Bitlocker
IF objWMIService is Nothing THEN
wscript.quit
END IF

On Error GoTo 0 

Dim objEncrypt

'Get drive encryption status
FOR EACH objEncrypt in colEncrypt
Dim EncryptionMethod
Dim ProtectionStatus
objEncrypt.GetEncryptionMethod EncryptionMethod
objEncrypt.GetProtectionStatus ProtectionStatus
EXIT FOR
NEXT


'If ProtectionStatus != 0 and EncryptionMethod != 0, then bitlocker is active. Suspend bitlocker.
IF (ProtectionStatus <> 0) AND (EncryptionMethod <> 0) THEN
objEncrypt.DisableKeyProtectors

'Create a scheduled task to run the script at startup, in order to reenable bitlocker
schtasksCreateCommand = "schtasks /Create  /RU ""NT AUTHORITY\SYSTEM"" /SC ONSTART /TN ""bitlockerscript"" /TR ""C:\Windows\system32\wscript.exe " & scriptLoc & " //Nologo"""
Set objShell = WScript.CreateObject ("WScript.shell")
objShell.run schtasksCreateCommand,0
Set objShell = Nothing
wscript.quit
END IF

'If ProtectionStatus == 0 and EncryptionMethod != 0, then bitlocker is suspended. Reenable bitlocker.
IF (ProtectionStatus = 0) AND (EncryptionMethod <> 0) THEN
objEncrypt.EnableKeyProtectors

'Delete the scheduled task once bitlocker has been reenabled
Set objShell = WScript.CreateObject ("WScript.shell")
objShell.run "schtasks /Delete /TN ""bitlockerscript"" /F",0
Set objShell = Nothing
wscript.quit
END IF

