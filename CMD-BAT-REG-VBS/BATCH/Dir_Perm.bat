@ECHO ON

icacls C:\Bitmaps /grant EVERYONE:(OI)(CI)(F) /T
icacls C:\Temp /grant EVERYONE:(OI)(CI)(F) /T
icacls C:\WRM /grant EVERYONE:(OI)(CI)(F) /T

IF NOT "%ProgramFiles(x86)%"=="" (goto ARP64) else (goto ARP86)
:ARP64
icacls "%ProgramFiles(x86)%\Open Solutions" /grant EVERYONE:(OI)(CI)(F) /T
icacls "%ProgramFiles(x86)%\Nexus" /grant EVERYONE:(OI)(CI)(F) /T
icacls "%ProgramFiles(x86)%\VerantID" /grant EVERYONE:(OI)(CI)(F) /T
:ARP86
icacls "%ProgramFiles%\Open Solutions" /grant EVERYONE:(OI)(CI)(F) /T
icacls "%ProgramFiles%\Nexus" /grant EVERYONE:(OI)(CI)(F) /T
icacls "%ProgramFiles%\VerantID" /grant EVERYONE:(OI)(CI)(F) /T

