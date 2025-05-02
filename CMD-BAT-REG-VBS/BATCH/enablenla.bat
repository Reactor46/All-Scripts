@echo off
rem ==========================================================================================================================
rem - Script: enablenla.bat
rem - 
rem - Author: Jeff Mason aka TNJMAN aka bitdoctor
rem - Date: Sep 12, 2013
rem - Purpose: To enable the remote desktop protocol (RDP) NetworK-Level Authentication (NLA) feature on a list of remote 
rem - (target) computers, so that the target computers will require any connecting (source) computer to utilize 
rem -  NLA-capable RDP - i.e. Windows 7 or higher and Windows Server 2008 or higher (or older O/S with NLA-compliant RDP)
rem - Notes:
rem -  After making this setting change, the target system needs to be rebooted, since this is an HKLM (Local Machine) key.
rem -  After making this setting change, you no longer will be able to connect to the target system via systems that run
rem -  Windows XP or older, or Windows Server 2003 or older, unless the source systems have upgraded to NLA-capable RDP versions
rem -  How to Run this script:
rem -  1) Save this script as "c:\scripts\enablenla.bat"
rem -     a) To run against a single computer
rem -        Execute this script, passing the remote computer name as the only parameter, redirecting output/errors to a log file
rem -         c:\scripts\enablenla.bat Server1 >> c:\scripts\nla-log.txt 2>&1
rem -           where Server1 is the remote computer needing RDP-NLA to be enabled
rem -     b) To run against multiple computers:
rem -        i. Make a wrapper "bat" file to 'call' this (enablenla.bat) script - 
rem -           There is a "wrapper file" example in the "rem" statements at the end of this script
rem -       ii. In the "wrapper" file, place a series of "call" statements, each on a single line, for each remote computer
rem -           i.e. "call c:\scripts\enablenla Server1" (that would run this script against to modify the remote computer "Server1")
rem -      iii. Once you've entered all the lines containing all the target remote computer names, 
rem -           save that "wrapper" script as "c:\scripts\callnla.bat"
rem -       iv. Execute the "wrapper" file, redirecting output & errors to a log file:
rem -           c:\scripts\callnla.bat >> c:\scripts\nla-log.txt 2>&1
rem -     c) Examine log for successes and to troubleshoot any errors ("notepad  c:\scripts\nla-log.txt")
rem -     
rem - Assumptions: 
rem - 1) You can create/save this script and a wrapper script to a c:\scripts folder
rem - 2) You have the necessary privs/rights to modify the HKLM key on the targeted remote computers
rem -
rem ==========================================================================================================================
rem
echo.
echo "Adding NLA-ONLY key to remote computer %1"
echo.
reg add "\\%1\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" /v UserAuthentication /t REG_DWORD /d 1 /f
echo.
echo "Finished adding NLA key to remote computer %1"
rem
rem [EXAMPLE OF A CALLING WRAPPER SCRIPT: "callnla.bat"]
rem -
rem - This script would contain the names of the remote computers on which you wish to enable RDP-NLA
rem -
rem - Create a script similar to below (save as "c:\scripts\callnla.bat"):
rem -
rem rem callnla.bat (wrapper script to make the call to "enablenla.bat"
rem rem
rem call c:\scripts\enablenla.bat Workstation1
rem call c:\scripts\enablenla.bat Workstation2
rem call c:\scripts\enablenla.bat Workstation3
rem call c:\scripts\enablenla.bat Server1
rem call c:\scripts\enablenla.bat Server2
rem call c:\scripts\enablenla.bat Server3



