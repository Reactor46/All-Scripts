@echo off
rem --------------------------------------------------------------------------------------------------
rem Script: spol.bat (Security Policy batch file)
rem Author: Jeff Mason / aka bitdoctor / aka tnjman
rem Date:   08/28/2013
rem Purpose: automate copying of local security policy/rights settings from a source computer, to a central 
rem network location, so the settings can be copied from that network location to one or more remote computers. 
rem   This is for cases where you may have workgroups that are not able to use domain/GPO.
rem   Notes: One-time creation of flag \\server\scripts\spol\copynet.txt - if flag file not 
rem     there, create it & then copy the GroupPolicy folder to \\server\scripts\spol\
rem     Uses "secedit" to import the security policy to the local sdb on the Target computer
rem  *Warning: This version allows for passing of admin credentials "in the clear", to do the import
rem   It can be used in 'batch mode,' by removing all "rem" statements and creating a "wrapper bat file" 
rem   containing a "call" to this script; i.e. spol-servers.bat wrapper file might look like this:
rem call \\server\scripts\spol.bat net pc2 %1 %2
rem call \\server\scripts\spol.bat net pc3 %1 %2
rem call \\server\scripts\spol.bat net pc3 %1 %2
rem Then, just execute like: spol-severs.bat mydom\myUserID mypassword
rem   Assumptions:
rem     1) You have local admin permissions on the source pc and target pc
rem     2) You have network permissions to  copy and create the "copynet" location and file
rem     3) You have "psexec" utility in c:\tools folder
rem     4) You've created a \\server\scripts share and an "spol" sub-folder within that share,
rem        and you store & access your scripts on "\\server\scripts" (fyi, 'server' can be your workstation)
rem     NOTE: Replace "your-script-server-goes-here" with the name of your own server; i.e. my-server1
rem     WARNING - You may get a WARNING like "No mapping between account names and security IDs was done."
rem        In many cases, this can be ignored, because it may be a domain-based ASP account that is different
rem        between the two computers and is not needed on the destination, check %windir%\security\logs\scesrv.log
rem        just to be sure, if you are concerned.
rem --------------------------------------------------------------------------------------------------
rem - Variables and Aliases and Parameters, oh my! -
rem NOTE: Replace "your-script-server-goes-here" with the name of your own server; i.e. my-server1
set svr=your-script-server-goes-here
set uid=%3
set pwd=%4
set rx=c:\tools\psexec
set flag=\\%svr%\scripts\spol\copynet.txt
set nloc=\\%svr%\scripts\spol
rem
rem [ Param 1 = source pc or net, Param 2 = dest pc, Param 3 = domain\UserID Param 4 = password ]
echo.
echo "Make sure you execute like this: "
echo "   spol pc1 pc2 mydom\myUserID mypassword"
echo "     where "pc1 is Source, and pc2 is Target
echo "     this will copy baseline policy from PC1 to Network and to PC2"
echo "or: 
echo "   spol net pc2 mydom\myUserID mypassword 
echo     if you already copied a baseline policy to the network" 
echo.
pause "If okay, press Enter to proceed, else ctrl-c to abort"
echo.
if %1% == "net" or %1% == "NET" or %1% == "Net" (
  if exist %flag% (
     echo.
     echo "Network Flag already exists, will process from NETWORK: %nloc%\GroupPolicy"
     echo.
     echo "Note: if you want specific Source Computer policy re-copied, 
     echo " just delete the %flag% file and then rerun this script"
     echo.
     pause "If okay, press Enter to proceed, else ctrl-c to abort, delete flag file, then re-run script with Source & Dest"
     goto processnet
     ) else (
    echo.
    echo "WARNING: NO Source computer given; and no flag file exists!"
    echo "Please re-run with Source Computer [and] Target Computer, 
    echo "   to copy baseline to network & target;"
    echo "i.e., spol pc1 pc2 mydom\myUserID mypassword"
    echo.
    goto badfin
  )
)
rem --------------------------------------------------------------------------------------------------
:copynet
rem --------------------------------------------------------------------------------------------------
echo.
echo "Doing one-time copy of group policy folder "
echo "  from source PC %1 to network location: %nloc%"
xcopy \\%1\c$\windows\system32\GroupPolicy %nloc%\GroupPolicy /s /e /h /i /y
attrib -h %nloc%\GroupPolicy
echo.
pause "Pause after copy of policy to the network..."
echo.
echo "Extracing policy from source PC %1 to network location: %nloc%"
echo.
%rx%  \\%1 secedit /export /cfg c:\security.inf
pause "Pause after local extract, next, copy security.inf to %nloc%"
copy  \\%1\c$\security.inf %nloc%\security.inf /Y
echo  "Creating copynet flag file..."
echo  "copynet completed" > %nloc%\copynet.txt
echo.
echo  "Done copying security.inf & local security policy to central network location."
echo  "Done creating copynet flag file"
rem
echo  "Listing %nloc% area, so you can verify if anything is missing..."
dir   %nloc%
rem
pause "Enter to proceed"
rem --------------------------------------------------------------------------------------------------
:processnet
rem --------------------------------------------------------------------------------------------------
echo.
echo   "Unhiding Destination computer grouppolicy folder on TARGET PC %2"
attrib -h \\%2\c$\windows\system32\GroupPolicy
echo.
pause "Pausing after unhide..."
echo  "Copying group policy from source Network location %nloc% to destination computer %2"
xcopy %nloc%\GroupPolicy \\%2\c$\Windows\GroupPolicy /s /e /h /i /y
echo.
pause "Pause before rehiding destination computer's group policy folder "
echo  "Re-hiding destination computer's grouppolicy folder on computer %2"
attrib +h \\%2\c$\windows\system32\GroupPolicy
echo.
echo  "Copy source: %nloc%\security.inf file to dest: \\%2%\c$\windows\inf folder"
copy  %nloc%\security.inf \\%2\c$\windows\inf /Y
pause "Pause after copying security.inf file to target computer %2"
echo.
echo  "Finished copying grouppolicy & security.inf from Source to Destination"
pause "Enter to proceed"
echo.
echo  "You may be able to ignore warnings about no security mapping between accounts, see WARNING at top of this script"
echo.
echo  "Importing security settings onto the Dest PC"
%rx%  \\%2 -u %uid% -p %pwd% secedit /configure /cfg c:\windows\inf\security.inf /db defltbase.sdb /verbose
:fin
 echo.
 echo "Done. Be sure and double-check any user rights, etc. expected from this policy, 
 echo "  and you may need to reboot the dest pc %2"
:badfin

