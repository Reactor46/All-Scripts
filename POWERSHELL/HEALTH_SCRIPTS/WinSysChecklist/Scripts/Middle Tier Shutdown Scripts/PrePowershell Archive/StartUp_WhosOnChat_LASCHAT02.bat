
:: START UP Whos On Chat Application services.


echo Do you really want to START services? (Y/N)
set INPUT=
set /P INPUT=Type input: %=%
If /I "%INPUT%"=="y" goto RUNSCRIPT 
If /I "%INPUT%"=="n" goto CANCELSCRIPT
echo Incorrect input & goto ENDScript

:RUNSCRIPT

@ECHO OFF
:: STARTS UP Whos On Chat Application services on LASCHAT02.

cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOn start 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnGateway start 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnQuery start 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnReports start 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnServiceMonitor start 10

goto ENDSCRIPT

:CANCELSCRIPT

echo You decided to exit script by entering %input%

:ENDScript

echo end of script
