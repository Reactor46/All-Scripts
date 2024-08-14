@ECHO OFF
:: ### FSMORoles Transfer-Script
::
:: ### Moves Active Directory FSMO Roles of one Domaincontroller to another, using the NTDSUTIL.EXE
:: ### Transfer Schema master
:: ### Transfer domain naming master
:: ### Transfer infrastructure master
:: ### Transfer PDC
:: ### Transfer RID master
::
:: Author: Roland Schoen
::   Mail: <roland.schoen@atp-autoteile.de>
::   Create Date: 2006-03-21
::   Change Date: 2014-01-24
::

SETLOCAL
:START
SET viewdomain=%1

::Catch emptry runtime parameters
If Not Defined viewdomain (
  ECHO.
  ECHO Move Active Directory FSMO Roles using NTDSUTIL.EXE
  ECHO.
  ECHO %~n0 [DNSDOMAIN]
  ECHO.
  ECHO   DNSDOMAIN  Must be a Fully Qualified Domain name ^(FQDN^) of a 
  ECHO              existing Microsoft Active Directory Domain
  ECHO.
  ECHO   The script works only if you are logged on as a administrator
  ECHO.	
  GOTO END
)


GOTO CHOICES

:CHOICES
ECHO **************************************************
ECHO ----------- FSMORoles Transfer Script ------------
ECHO **************************************************
ECHO #                                                #
ECHO #########  What do you want to do today? #########
ECHO #                                                #
ECHO #########       Make your choices        #########
ECHO #                                                #
ECHO #         A) Transfer Schema Master Role         #
ECHO #         B) Transfer Domain Naming Master Role  #
ECHO #         C) Transfer Infrastructure Master Role #
ECHO #         D) Transfer PDC Master Role            #
ECHO #         E) Transfer RID Master Role            #
ECHO #                                                #
ECHO #         O) Overview - current role holders     #
ECHO #         Q) Quit the Script                     #
ECHO #                                                #
ECHO ##################################################

CHOICE /C:abcdeoq /M "Choose your task: "
CLS

IF ERRORLEVEL 7 GOTO END
IF ERRORLEVEL 6 GOTO OVERVIEW
IF ERRORLEVEL 5 GOTO RID
IF ERRORLEVEL 4 GOTO PDC
IF ERRORLEVEL 3 GOTO INFRASTRUCTURE
IF ERRORLEVEL 2 GOTO DOMAIN
IF ERRORLEVEL 1 GOTO SCHEMA

::Create and Display Overview
:OVERVIEW

IF EXIST %TEMP%\temp.log DEL %TEMP%\temp.log 
IF EXIST %TEMP%\fsmoroles.log DEL %TEMP%\fsmoroles.log

ntdsutil "roles" "Connections" "Connect to domain %viewdomain%" "Quit" "Select operation target" "List roles for connected server" "Quit" "Quit" "Quit" >> %TEMP%\temp.log
FOR /F "tokens=*" %%a IN (%TEMP%\temp.log) DO Echo %%a | find /I "Schema -" >> %TEMP%\fsmoroles.log
FOR /F "tokens=*" %%a IN (%TEMP%\temp.log) DO Echo %%a | find /I "Naming Master -" >> %TEMP%\fsmoroles.log
FOR /F "tokens=*" %%a IN (%TEMP%\temp.log) DO Echo %%a | find /I "Infrastructure -" >> %TEMP%\fsmoroles.log
FOR /F "tokens=*" %%a IN (%TEMP%\temp.log) DO Echo %%a | find /I "PDC -" >> %TEMP%\fsmoroles.log
FOR /F "tokens=*" %%a IN (%TEMP%\temp.log) DO Echo %%a | find /I "RID -" >> %TEMP%\fsmoroles.log


ECHO **************************************************
ECHO ----------- Current FSMO Role Holders ------------
ECHO **************************************************
TYPE %TEMP%\fsmoroles.log

:: DELETE TEMP Logfiles
IF EXIST %TEMP%\temp.log DEL %TEMP%\temp.log 
IF EXIST %TEMP%\fsmoroles.log DEL %TEMP%\fsmoroles.log
ECHO.
ECHO.
GOTO CHOICES

::Transfer of Schema master
:SCHEMA
ECHO Wich Domaincontroller should become SCHEMA MASTER?
SET /P SCHEMAMASTER="Enter Servername: "

ECHO Transfering the Schema Master Role to %SCHEMAMASTER%
ECHO Please Wait...
ntdsutil "roles" "Connection" "Connect to Server %SCHEMAMASTER%" "Quit" "Transfer Schema master" "Quit" "Quit"
ECHO.
ECHO.
GOTO CHOICES

::Transfer of domain naming master
:DOMAIN
ECHO Wich Domaincontroller should become DOMAIN NAMING MASTER?
SET /P DOMAINMASTER="Enter Servername: "

ECHO Transfering the Domain Naming Master Role to %DOMAINMASTER%
ECHO Please Wait...
ntdsutil "roles" "Connection" "Connect to Server %DOMAINMASTER%" "Quit" "Transfer domain naming master" "Quit" "Quit"
ECHO.
ECHO.
GOTO CHOICES

::Transfer of infrastructure master
:INFRASTRUCTURE
ECHO Wich Domaincontroller should become INFRASTRUCTURE MASTER?
SET /P INFRASTRUCTUREMASTER="Enter Servername: "

ECHO Transfering the Infrastructure Master Role to %INFRASTRUCTUREMASTER%
ECHO Please Wait...
ntdsutil "roles" "Connection" "Connect to Server %INFRASTRUCTUREMASTER%" "Quit" "Transfer infrastructure master" "Quit" "Quit"
ECHO.
ECHO.
GOTO CHOICES

::Transfer of pdc master
:PDC
ECHO Wich Domaincontroller should become PDC MASTER?
SET /P PDCMASTER="Enter Servername: "

ECHO Transfering the PDC Master Role to %PDCMASTER%
ECHO Please Wait...
ntdsutil "roles" "Connection" "Connect to Server %PDCMASTER%" "Quit" "Transfer PDC" "Quit" "Quit"
ECHO.
ECHO.
GOTO CHOICES

::Transfer of rid master
:RID
ECHO Wich Domaincontroller should become RID MASTER?
SET /P RIDMASTER="Enter Servername: "

ECHO Transfering the RID Master Role to %RIDMASTER%
ECHO Please Wait...
ntdsutil "roles" "Connection" "Connect to Server %RIDMASTER%" "Quit" "Transfer RID master" "Quit" "Quit"
ECHO.
ECHO.
GOTO CHOICES

:END
EcHO. 
ECHO Done!

ENDLOCAL