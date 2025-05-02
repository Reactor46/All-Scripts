IF EXIST "c:\snow.txt" GOTO eof

msiexec /i "\\uson.local\netlogon\SNOW\UNITEDHEALTHCARE_3704_uid_x64 - OPTUMMSO.msi" /qn  /log c:\snow.txt

:eof


EXIT