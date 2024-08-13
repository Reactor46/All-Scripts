@ECHO OFF

:: SHUT DOWN CreditOne Application IIS and Windows Services.


echo Do you really want to STOP services? (Y/N)
set INPUT=
set /P INPUT=Type input: %=%
If /I "%INPUT%"=="y" goto RUNSCRIPT 
If /I "%INPUT%"=="n" goto CANCELSCRIPT
echo Incorrect input & goto ENDScript

:RUNSCRIPT

cscript.exe SvcManager.vbs Contosocorp LASCAS01 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCAS02 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCAS03 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCAS04 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCAS05 W3SVC stop 300

cscript.exe SvcManager.vbs Contosocorp LASCOLL01 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL05 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL06 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL07 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL08 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL09 W3SVC stop 300

cscript.exe SvcManager.vbs Contosocorp LASCAPS01 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCAPS02 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCAPS03 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCAPS04 W3SVC stop 300

goto ENDSCRIPT


:CANCELSCRIPT

echo You decided to exit script by entering %input%


:ENDScript


echo end of script
