@ECHO ON

Powershell.exe -NoLogo -ExecutionPolicy ByPass

CD C:\Patches
$File = "C:\Patches\Set-ScreenResolution.ps1"
$FileExists = Test-Path $File
If ($FileExists -eq $True) {.\Set-ScreenResolution.ps1 -width 1920 -height 1080 }
Else {copy-item \\Contosocorp\NETLOGON\WIN10\Set-ScreenResolution.ps1 C:\Patches\Set-ScreenResolution.ps1 }

.\Set-ScreenResolution -width 1920 -height 1080

Exit