@ECHO OFF
msiexec /i "\\uson.local\NETLOGON\PrinterLogic\PrinterInstallerClient.msi" HOMEURL=http://msopl01.uson.local AUTHORIZATION_CODE=kdmndjv0 REINSTALLMODE=vomus REBOOT=ReallySupress REINSTALL=ALL /qn /norestart

net start PrinterInstallerLauncher