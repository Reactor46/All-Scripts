@ECHO ON
ECHO JAVA INSTALL TEST

::Installing Current Java

IF NOT "%ProgramFiles(x86)%"=="" (goto ARP64) else (goto ARP86)

:ARP64
start /wait %CD%\Apps\java\jre-7u25-windows-x64.exe /s /v"/norestart AUTOUPDATECHECK=0 JAVAUPDATE=0 JU=0"
:ARP86
start /wait %CD%\Apps\java\jre-7u25-windows-i586.exe /s /v"/norestart AUTOUPDATECHECK=0 JAVAUPDATE=0 JU=0"

:END