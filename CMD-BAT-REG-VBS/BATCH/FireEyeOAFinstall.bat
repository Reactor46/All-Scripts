
IF EXIST "c:\fireeyehx.txt" GOTO eof
msiexec /i "\\uson.local\NETLOGON\FireEyeHX\IMAGE_HX_AGENT_WIN_29.7.9\xagtSetup_29.7.9_universal.msi" CONFJSONDIR="\\uson.local\NETLOGON\FireEyeHX\IMAGE_HX_AGENT_WIN_29.7.9" /quiet /log c:\fireeyehx.txt
:eof

EXIT