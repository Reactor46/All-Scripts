@ECHO off

REM Help Desk Menu By Haim Cohen 2014
REM Feedback: http://il.linkedin.com/in/haimc
REM Blog: http://blogs.microsoft.co.il/skepper/
REM Last edited 16/08/14 by Daniel Petri (www.petri.com)

REM External tools location – set it to match your own path
SET TOOLS=C:HelpDesk_MenuTools

REM

################################################
:start
title Helpdesk Tools Script
color 0a
CLS
ECHO. Helpdesk Tools v3.4, By Haim Cohen 2014
ECHO. Last edited 18/08/14 by Daniel Petri (www.petri.com)
ECHO. %DATE% %TIME:~0,2%:%TIME:~3,2%
ECHO. All Commands Running as:%USERDOMAIN%%USERNAME%
ECHO. Last targetted IP/Hostname was: %IP%
ECHO.
ECHO. ===================================================
ECHO 1. Ping
ECHO 2. Nslookup
ECHO 3. RDP (MSTSC)
ECHO 4. Port Query
ECHO 5. Remote logged users, sessions query and logoff remote sessions
ECHO 6. Reset user’s pwd, unlock account, change pwd at next logon
ECHO 7. Display remote tasks, kill remote tasks
ECHO 8. Services on remote computer
ECHO 9. Computer Manager on remote computer
ECHO 10. CMD on remote computer
ECHO 11. Gpupdate /force on remote computer
ECHO 12. Get serial from remote computer
ECHO 13. VNC Viewer
ECHO 14. PuTTY (SHH or Telnet) on remote computer
ECHO 15. Documenting remote computer – (Microsoft Word must be installed)
ECHO 16. Rename computer name remotely
ECHO 96. Open PowerShell
ECHO 97. Open CMD
ECHO 98. Restart remote computer
ECHO X. Exit
ECHO.
SET /p choice=Please enter command number:

IF “%choice%”==”1″ GOTO step1
IF “%choice%”==”2″ GOTO step2
IF “%choice%”==”3″ GOTO step3
IF “%choice%”==”4″ GOTO step4
IF “%choice%”==”5″ GOTO step5
IF “%choice%”==”6″ GOTO step6
IF “%choice%”==”7″ GOTO step7
IF “%choice%”==”8″ GOTO step8
IF “%choice%”==”9″ GOTO step9
IF “%choice%”==”10″ GOTO step10
IF “%choice%”==”11″ GOTO step11
IF “%choice%”==”12″ GOTO step12
IF “%choice%”==”13″ GOTO step13
IF “%choice%”==”14″ GOTO step14
IF “%choice%”==”15″ GOTO step15
IF “%choice%”==”16″ GOTO step16
IF “%choice%”==”96″ GOTO step96
IF “%choice%”==”97″ GOTO step97
IF “%choice%”==”98″ GOTO step98
IF /I “%choice%”==”x” GOTO stepx

ECHO.
GOTO start

REM

################################################
:step1
REM Ping
CLS
ECHO Selected Command: Ping
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO nextstep)
PAUSE
GOTO step1

:nextstep
ECHO(
SET /P LIMIT4=Limit to IPv4? (-4) ([y]/n)
IF /I “%LIMIT4%” NEQ “n” (SET PINGSWITCH=-a -4) ELSE (SET PINGSWITCH=-a)
ECHO(
SET /P PINGCONT=Use continuous Ping? (-t) (y/[n])
IF /I “%PINGCONT%” NEQ “y” (SET CONTSWITCH=) ELSE (SET CONTSWITCH=-t)
ping %PINGSWITCH% %IP% %CONTSWITCH%
GOTO start

REM

################################################
:step2
REM Nslookup
CLS
ECHO Selected Command: Nslookup
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO nextstep)
PAUSE
GOTO step2

:nextstep
nslookup %IP%
PAUSE
GOTO start

REM

################################################
:step3
REM RDP
CLS
ECHO Selected Command: RDP
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO nextstep)
PAUSE
GOTO step3

:nextstep
ECHO.
ECHO Connecting to %IP%
START mstsc /v %IP%
GOTO start

REM

################################################
:step4
REM Port Query
CLS
ECHO(
ECHO Selected Command: Port Query
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO nextstep)
PAUSE
GOTO step4

:nextstep
ECHO Enter single port to query (valid range: 1-65535):
ECHO(
SET /P PORT=
%tools%portqry.exe -n %IP% -e %PORT% | findstr TCP*
PAUSE
GOTO start

REM

################################################
:step5
REM Query remote sessions + Logoff remote sessions
CLS
ECHO Selected Command: Query remote sessions + Log off remote sessions
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO LISTSESSIONS)
PAUSE
GOTO step5

:LISTSESSIONS
ECHO(
ECHO Review the active sessions and the session IDs:
ECHO(
quser /server:%IP%
ECHO(

:SESSIONACTION
ECHO(
ECHO Do you want to [L]og of a session, or [C]ancel? (l/c)
SET /p choice=
ECHO(
IF /I “%choice%”==”l” GOTO KILLSESSION
IF /I “%choice%”==”c” GOTO start
GOTO SESSIONACTION

:KILLSESSION
ECHO(
ECHO Enter session ID to log off
SET /P ID=
IF “%ID%”==”” (ECHO You must enter an active session ID) ELSE (GOTO KILLSESSIONACTION)
PAUSE
GOTO KILLSESSION

:KILLSESSIONACTION
CLS
ECHO Logging of session ID %ID%
ECHO(
logoff /server:%IP% %ID%
ECHO(
ECHO Review the active sessions:
ECHO(
quser /server:%IP%
ECHO(
GOTO SESSIONACTION

REM

################################################

:step6
REM Reset user’s password, unlock account, and configure user to change their password at next logon
CLS
ECHO Reset user’s password, unlock account,
ECHO and configure user to change their password at next logon
ECHO You must specifiy an existing user in SAMID format (i.e. user1)
ECHO(
ECHO Enter username to be modified:
SET /p targetusername=
ECHO(

REM Verify if user exists in AD
dsquery user domainroot -samid %targetusername% | findstr /i /c:%targetusername% 1>NUL 2>NUL
IF %ERRORLEVEL% EQU 1 (ECHO User not found!) ELSE (GOTO VERIFYUSER)
ECHO(
ECHO You must enter an existing username. Try again.
PAUSE
GOTO step6

:VERIFYUSER
ECHO(
ECHO Please verify your selected user account:
dsquery user domainroot -samid %targetusername%
ECHO(

:CHOOSEACTION6
ECHO Select action to perform: [U]nlock user, [R]eset password, or [C]ancel (u/r/c)
SET /p choice=
ECHO(
IF /I “%choice%”==”u” GOTO UNLOCKUSER
IF /I “%choice%”==”r” GOTO RESETPWD
IF /I “%choice%”==”c” GOTO start
GOTO CHOOSEACTION6

:UNLOCKUSER
CLS
ECHO(
ECHO You chose to unlock the user’s account.
ECHO(
PAUSE
ECHO(
dsquery user domainroot -samid %targetusername% | dsmod user -disabled no
ECHO(
ECHO User’s account was unlocked.
ECHO(
PAUSE
GOTO start

:RESETPWD
CLS
ECHO(
ECHO You chose to reset the user’s password.
ECHO(
ECHO Before you reset a password, please note:
ECHO 1. Make sure pwd meets complexity and minimum length requirements
ECHO 2. Password cannot be blank
ECHO 3. Typed password will be visible on the screen
ECHO 4. If user has any EFS-encrypted files, they may not be accessible
ECHO unless decrypted by the domain’s Recovery Agent user account
ECHO(
PAUSE
ECHO(
ECHO Enter new password:
SET /p userpwd=
IF “%userpwd%”==”” (ECHO Password cannot be blank) ELSE (GOTO RESETPWDACTION)
PAUSE
GOTO RESETPWD

:RESETPWDACTION
ECHO(
dsquery user domainroot -samid %targetusername% | dsmod user -pwd %userpwd% -mustchpwd yes
ECHO(
ECHO User’s password was reset. User must change password at next logon.
ECHO(
PAUSE
GOTO start

REM

################################################

:step7
REM Display and kill processes on a remote computer
CLS
ECHO Selected Command: Display and kill processes on a remote computer
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Note: You must enter a hostname only (not an IP address)
SET /P IP=
IF “%IP%”==”” (ECHO You must enter a hostname) ELSE (GOTO listprocess)
PAUSE
GOTO step7

:listprocess
tasklist /s %IP%
ECHO(
ECHO Review the running tasks and PID numbers
ECHO(
ECHO(

:processaction
ECHO(
ECHO Do you want to [K]ill a remote process, [R]efresh list, or [C]ancel? (k/r/c)
SET /p choice=
ECHO(
IF /I “%choice%”==”k” GOTO killprocess
IF /I “%choice%”==”r” GOTO listprocess
IF /I “%choice%”==”c” GOTO start
GOTO processaction

:killprocess
ECHO(
ECHO Enter PID ID to kill
SET /P PID=
ECHO Kill Remote Machine, spinning process, please wait…
taskkill /s %IP% /PID %PID%
PAUSE
GOTO start

REM

################################################

:step8
REM Remote Services
CLS
ECHO Selected Command: Services on a remote computer
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO nextstep)
PAUSE
GOTO step8

:nextstep
services.msc /computer:%IP%
PAUSE
GOTO start

REM

################################################

:step9
REM Remote Computer Management
CLS
ECHO Selected Command: Computer Management on a remote computer
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO To manage a Windows machine remotely,
ECHO Windows Firewall rules must be enabled on the remote computer:
ECHO 1. COM+ Network Access (DCOM-In)
ECHO 2. All rules in the Remote Event Log Management group
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO nextstep)
PAUSE
GOTO step9

:nextstep
compmgmt.msc /computer:%IP%
PAUSE
GOTO start

REM

################################################

:step10
REM Remote Command Prompt
CLS
ECHO Selected Command: Open Command Prompt on a remote computer
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO nextstep)
PAUSE
GOTO step10

:nextstep
%tools%psexec.exe \%IP% cmd.exe
PAUSE
GOTO start

REM

####################################################

:step11
REM Remote gpupdate /force
CLS
ECHO Selected Command: Run gpupdate /force on a remote computer
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO nextstep)
PAUSE
GOTO step11

:nextstep
%tools%psexec.exe \%IP% gpupdate /force
PAUSE
GOTO start

REM

##################################################

:step12
REM Get Serial
CLS
ECHO Selected Command: Get serial from a remote computer
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Only IP addresses can be used for this command!
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address) ELSE (GOTO nextstep)
PAUSE
GOTO step12

:nextstep
ECHO.
wmic /node:%IP% bios get serialnumber
wmic /node:%IP% bios get serialnumber >%tmp%%IP%_Serial.txt
notepad %tmp%%IP%_Serial.txt
ECHO.
PAUSE
GOTO start

REM

###################################################

:step13
REM VNC
CLS
ECHO Selected Command: VNC Viewer
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO nextstep)
PAUSE
GOTO step13

:nextstep
START %tools%vncviewer.exe %IP%
GOTO start

REM

####################################################

:step14
REM PuTTY – SSH and Telnet Client
CLS
ECHO(
ECHO Selected Command: PuTTY
ECHO(
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO choseaction14)
PAUSE
GOTO step14

:choseaction14
ECHO(
ECHO Select: [S]SH, [T]elnet or [C]ancel (t/s/c)
SET /p choice=
ECHO(
IF /I “%choice%”==”s” GOTO SSH
IF /I “%choice%”==”t” GOTO TELNET
IF /I “%choice%”==”c” GOTO start
GOTO choseaction14

:SSH
ECHO(
ECHO Opening SSH to %IP%
ECHO(
START %tools%putty.exe -ssh %IP%
GOTO start

:TELNET
ECHO(
ECHO Opening Telnet to %IP%
ECHO(
START %tools%putty.exe -telnet %IP%
GOTO start

REM

###################################################

:step15
REM Inventory and Documenting Remote Computer – Full Report
CLS
ECHO Selected Command: Inventory and Documenting Remote Computer – Full Report
ECHO Microsoft Word must be installed on the current computer.
ECHO(
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO DOCUMENTINGOPTION)
PAUSE
GOTO step15

:DOCUMENTINGOPTION
ECHO Do you want [F]ull reporting or [M]inimal reporting? ([f]/m)
SET /P DOCUMENTINGOPTIONQ=
IF /I “%DOCUMENTINGOPTIONQ%” NEQ “m” SET DOCUMENTINGOPTIONA=-wabefghipPqrsu -racdklp
IF /I “%DOCUMENTINGOPTIONQ%”==”m” SET DOCUMENTINGOPTIONA=-w -r

cscript %tools%sydisydi-server.vbs %DOCUMENTINGOPTIONA% -ew -f10 -d -t%IP%
GOTO start

REM

#####################################################

:step16
REM Rename Computer Name
CLS
ECHO Selected Command: Rename Computer Name
SET /P POLDC= Type in the OLD computer name:
SET /P PNEWC= Type in the NEW computer name:
SET /P UID= Type in the DOMAINUSER:
REM SET /P PAS= Type in the PASSWORD:
ECHO(
ECHO Shall I also reboot the remote computer? ([y]/n)
SET /P REBOOTCOMPQ=
IF /I “%REBOOTCOMPQ%” NEQ “n” SET REBOOTCOMP=/reboot

:nextstep
setlocal
color 1c
ECHO(
ECHO(
ECHO Computer %POLDC% will be renamed
IF /I “%REBOOTCOMPQ%” NEQ “n” ECHO and rebooted, all unsaved work may be lost.
ECHO(
ECHO Are you sure (y/[n])?
SET /P AREYOUSURE=
IF /I “%AREYOUSURE%” NEQ “y” GOTO end
ECHO(
netdom renamecomputer %POLDC% /newname:%PNEWC% /userd:%UID% /passwordd:* /force %REBOOTCOMP%
PAUSE

:end
endlocal
GOTO start

REM

###########################################################

:step96
REM Open Powershell Console
CLS
start powershell
GOTO start

REM

############################################################

:step97
REM Open New CMD Console
CLS
start
GOTO start

REM

##########################################################

:step98
REM Restart Remote Computer
CLS
ECHO Selected Command: Restart Remote Computer
ECHO Last targetted IP/Hostname was: %IP%
ECHO(
ECHO Enter IP or Hostname
SET /P IP=
IF “%IP%”==”” (ECHO You must enter an IP address or hostname) ELSE (GOTO nextstep)
PAUSE
GOTO step98

:nextstep
setlocal
color 1c
ECHO(
ECHO(
ECHO Computer %IP% will reboot in 10 seconds, all unsaved work may be lost.
ECHO(
ECHO Are you sure (y/[n])?
SET /P AREYOUSURE=
IF /I “%AREYOUSURE%” NEQ “y” GOTO end
ECHO(
ECHO Restarting %IP% in 10 seconds…
ECHO(
shutdown /r /f /t 10 /m %IP%
PAUSE

:end
endlocal
GOTO start

REM

######################################################

:stepx
REM Exit
msg * /TIME:3 “Thank you for using this tool, Haim.”
exit

REM

######################################################