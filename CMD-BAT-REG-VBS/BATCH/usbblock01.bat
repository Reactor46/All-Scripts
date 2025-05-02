REGEDIT4

; @ECHO OFF
; CLS
; REGEDIT.EXE /S "%~f0"
; EXIT

[HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\USBSTOR]
"Start"=dword:00000004

[HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\CDROM]
"Start"=dword:00000004

takeown /s %computername% /u "SHREK" /p "P@$$w0rd" /f %systemroot%\inf\usbstor.inf
takeown /s %computername% /u "SHREK" /p "P@$$w0rd" /f %systemroot%\inf\usbstor.pnf
icacls usbstor.inf /deny builtin\users:(WDAC)
icacls usbstor.pnf /deny builtin\Administrators:(D,WDAC)

REM Created by John Sam George for Warba Bank