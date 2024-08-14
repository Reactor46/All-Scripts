@ECHO OFF

REM *********************************************************************
REM Script Title : Add_Account_to_LocalAdmins.cmd
REM
REM Author	 : Shawn Gibbs
REM
REM Description  : Add an account to local admins that is provided by arg
REM
REM *********************************************************************

IF "%1"=="" GOTO Syntax

ECHO Adding %1 to local administrators group >> %windir%\temp\Output.log
net localgroup administrators %1 /add >> %windir%\temp\Output.log
ECHO. >> %windir%\temp\Output.log

GOTO :End

:SYNTAX
ECHO.
ECHO Account Name Required as cmd argument >> %windir%\temp\Output.log
ECHO.

REM If arguments doesn't exist batch files exits with code 9
EXIT 9

:End
EXIT 0