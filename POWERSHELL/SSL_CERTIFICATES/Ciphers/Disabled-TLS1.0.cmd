@ECHO ON
regedit -s /i "\\lasfs03\software\current Versions\Microsoft\Windows\Windows10\Disable-TLS1.0.reg"

reg query HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols /s
