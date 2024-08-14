@ECHO OFF

REM *********************************************************************
REM Script Title : Mod_DefaultWebSite.cmd
REM
REM Author	 : Shawn Gibbs
REM
REM Description  : Configures IIS Application Pools, Delete Default Web Site
REM                and assign default app pool to site
REM
REM *********************************************************************

IF "%1"=="" GOTO Syntax

ECHO Delete 'Defaul Web Site' >> %windir%\temp\Output.log
%windir%\system32\inetsrv\appcmd delete site "Default Web Site" >> %windir%\temp\Output.log

ECHO Creating folder to use for new site at c:\inetpub\wwwroot\%1 >> %windir%\temp\Output.log
mkdir c:\inetpub\wwwroot\%1 >> %windir%\temp\Output.log

ECHO Create the Web Site named %1 >> %windir%\temp\Output.log
%windir%\system32\inetsrv\appcmd ADD SITE /name:%1 /id:100 /bindings:http/*:80: /physicalPath:C:\inetpub\wwwroot\%1 /state:started >> %windir%\temp\Output.log

ECHO Assign app pool to site %1 >> %windir%\temp\Output.log
%windir%\system32\inetsrv\appcmd set app "%1/" /applicationPool:"DefaultAppPool" >> %windir%\temp\Output.log

GOTO :End

:SYNTAX
ECHO.
ECHO Site Name Required is required by Mod_DefaultWebSite.cmd >> %windir%\temp\Output.log
ECHO.
EXIT 9

:END
EXIT 0