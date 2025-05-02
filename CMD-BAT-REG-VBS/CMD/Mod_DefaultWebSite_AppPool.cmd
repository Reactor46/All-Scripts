@ECHO OFF

REM *********************************************************************
REM Script Title : Mod_DefaultWebSite_AppPool.cmd
REM
REM Author	 : Shawn Gibbs
REM
REM Description  : Configures IIS Application Pools and Deletes 
REM                Default Web Site. Also creates new folder for site as
REM                well as configures asp.net 4.0 for new app pool.
REM Parameter    : Takes in name to use for site, app pool and folder
REM
REM *********************************************************************

IF "%1"=="" GOTO Syntax

ECHO Delete 'Defaul Web Site' >> %windir%\temp\Output.log
%windir%\system32\inetsrv\appcmd delete site "Default Web Site" >> %windir%\temp\Output.log

ECHO Create the App Pool >> %windir%\temp\Output.log
%windir%\system32\inetsrv\appcmd ADD APPPOOL /name:%1 /managedruntimeversion:v4.0 >> %windir%\temp\Output.log

ECHO Creating folder to use for new site at c:\inetpub\wwwroot\%1 >> %windir%\temp\Output.log
mkdir c:\inetpub\wwwroot\%1 >> %windir%\temp\Output.log

ECHO Create the Web Site >> %windir%\temp\Output.log
%windir%\system32\inetsrv\appcmd ADD SITE /name:%1 /id:100 /bindings:http/*:80: /physicalPath:C:\inetpub\wwwroot\%1 /state:started >> %windir%\temp\Output.log

ECHO Assign app pool to site >> %windir%\temp\Output.log
%windir%\system32\inetsrv\appcmd set app "%1/" /applicationPool:"%1" >> %windir%\temp\Output.log
ECHO. >> %windir%\temp\Output.log

GOTO :End

:SYNTAX
ECHO.
ECHO Site Name Required as parameter of Mod_DefaultWebSite_AppPool.cmd >> %windir%\temp\Output.log
ECHO.
EXIT 9

:END
EXIT 0