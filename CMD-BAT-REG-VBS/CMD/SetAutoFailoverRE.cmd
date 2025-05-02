@ECHO OFF

REM Description:
REM     Ensures failover/recovery environment is setup
REM     1) Update the BCD store
REM	2) Setting auto failover configuration for Windows Recovery Environment
REM	3) Setting WinRE partition to be hidden

echo ----------------------------
echo Start script: %~n0
date /t
time /t
echo ----------------------------

:INIT
	set BCDEDIT=X:\windows\system32\bcdedit.exe
	set DISKPART=X:\windows\system32\diskpart.exe
	set DISKPART_SCRIPT=x:\sources\recovery\tools\DiskPart_script.txt
        set RECOVERY_DRIVE=NONE
Title Setting auto failover configuration for Windows Recovery Environment
:MAIN

Title  Update the BCD Store
REM 1) Update the BCD Store. 

	W:\Windows\system32\bcdboot.exe W:\windows

REM 2) Setting auto failover configuration for Windows Recovery Environment

for %%a in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do if exist %%a:\WinRE.TAG set RECOVERY_DRIVE=%%a
        REM Configure the path to Windows RE image using REAgentC.exe tool
         W:\Windows\System32\Reagentc.exe /setreimage /path R:\Recovery\WindowsRE /target W:\Windows
        REM Configure path to recovery image using REAgentC.exe tool
	 W:\Windows\System32\reagentc.exe /SetOSImage /Path R:\Recovery\WindowsRE /Target W:\Windows
         W:\Windows\System32\Reagentc.exe /setosimage /customtool /target W:\Windows
Title  Hide WinRE partition	
REM 3) Hide WinRE partition. 

	echo sel volume %RECOVERY_DRIVE% > %DISKPART_SCRIPT%
	echo remove >> %DISKPART_SCRIPT%
	echo set id=27 override >> %DISKPART_SCRIPT%
	%DISKPART% /s %DISKPART_SCRIPT%
  	del %DISKPART_SCRIPT%
echo ----------------------------
echo End script: %~n0
date /t
time /t
echo ----------------------------
