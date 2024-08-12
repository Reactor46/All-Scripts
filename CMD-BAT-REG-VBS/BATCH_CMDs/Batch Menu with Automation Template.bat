@ECHO OFF
cls

:START
set input=Y
set install=Z
set back=START
cls

title Install Options
echo Choose what to Install:               Other:
echo ( 1 ) Full Install                    ( A ) IE10 Windows 7
echo ( 2 ) Office                          ( B ) Windows 7 SP1
echo ( 3 ) Adobe Reader                    ( C ) VLC Media Player
echo ( 4 ) Java                            ( D ) AVG Free
echo ( 5 ) Flash                           
echo ( 6 ) Citrix Client                   
echo ( 7 ) Silverlight                     
echo.
echo ( Q ) Quit
echo.

:: Prompt for Install
Set /P input=
set install=%input:~0,1%
echo.

:: Verify
if /I %install% EQU Q goto END
if %install% EQU 1 goto Office2010
if %install% EQU 2 goto Office2010
if %install% EQU 3 goto ADOBE
if %install% EQU 4 goto JAVA
if %install% EQU 5 goto Flash
if %install% EQU 6 goto Citrix
if %install% EQU 7 goto SILVER
:: Optionals
if /I %install% EQU A goto IE10W7
if /I %install% EQU B goto W7SP1
if /I %install% EQU C goto VLC
if /I %install% EQU D goto AVGFREE
goto ERR0

:Office2010
cls
TITLE Installing Office 2010
echo Office 2010 is installing. . .
:: INSERT CODE HERE
if %install% EQU 2 goto START
goto ADOBE

:ADOBE
cls
title Installing Adobe Reader
echo Installing Adobe Reader
:: INSERT CODE HERE
if %install% EQU 3 goto START
goto JAVA

:JAVA
cls
title Installing Java
echo Installing Java
:: INSERT CODE HERE
if %install% EQU 4 goto START
goto FLASH

:FLASH
cls
title Installing Flash Player
echo Installing Adobe Flash Player. . .
:: INSERT CODE HERE
if %install% EQU 5 goto START
goto CITRIX

:CITRIX
cls
title Citrix Presentation Server Clients
echo Installing Citrix Presentation Server Clients
:: INSERT CODE HERE
if %install% EQU 6 goto START
goto SILVER

:SILVER
cls
title Installing Silverlight
echo Installing Microsoft Silverlight
:: INSERT CODE HERE
if %install% EQU 7 goto START
:: goto Next Program if there is another
goto Start

:AVGFREE
cls
Title AVG Free
echo Installing AVG 2013 Free. . .
:: INSERT CODE HERE
goto START

:IE10W7
cls
title Installing IE9
echo Installing IE9. . .
:: INSERT CODE HERE
goto START

:W7SP1
cls
title Installing Windows 7 SP1
echo Installing Windows 7 Service Pack 1. . .
:: INSERT CODE HERE
goto START

:VLC
cls
TITLE VLC Media Player
echo Installing VLC Media Player. . .
:: INSERT CODE HERE
goto START

:ERR0
echo '%input%' is an invalid entry.
echo.
pause
cls
goto %back%


:END
exit