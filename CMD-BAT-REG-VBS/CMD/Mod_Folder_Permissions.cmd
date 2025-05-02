@ECHO OFF

REM *********************************************************************
REM Script Title : Mod_Folder_Permissions.cmd
REM
REM Author	 : Shawn Gibbs
REM
REM Description  : Configures Permissions on directories used by IIS.
REM                This script was used to change data directory permisions
REM		   on a deployment of BlogEngine using XML files
REM
REM *********************************************************************

IF "%1"=="" GOTO Syntax

ECHO Stopping IIS to be able to change permissions on folder. Mod_Folder_Permissions.cmd >> %windir%\temp\Output.log
IISRESET /STOP >> %windir%\temp\Output.log

ECHO 'Modify Permissions' >> %windir%\temp\Output.log
cacls c:\inetpub\wwwroot\%1 /t /e /g everyone:f >> %windir%\temp\Output.log

ECHO Starting IIS after permissions on folder folder are changed. >> %windir%\temp\Output.log
IISRESET /START >> %windir%\temp\Output.log
ECHO. >> %windir%\temp\Output.log

GOTO End

:SYNTAX
ECHO.
ECHO Folder name is required for Mod_Folder_Permissions.cmd to run. >> %windir%\temp\Output.log
ECHO.
EXIT 9

:END
EXIT 0