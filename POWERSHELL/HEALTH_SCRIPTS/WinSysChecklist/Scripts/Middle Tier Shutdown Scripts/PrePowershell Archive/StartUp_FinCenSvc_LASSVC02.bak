
:: START UP Fin Cen Service.


echo Do you really want to START services? (Y/N)
set INPUT=
set /P INPUT=Type input: %=%
If /I "%INPUT%"=="y" goto RUNSCRIPT 
If /I "%INPUT%"=="n" goto CANCELSCRIPT
echo Incorrect input & goto ENDScript

:RUNSCRIPT

@ECHO OFF
:: STARTS UP Fin Cen Service on LASSVC02.

cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoFinCenService start 15

goto ENDSCRIPT

:CANCELSCRIPT

echo You decided to exit script by entering %input%

:ENDScript

echo end of script
