Rem Variable names of output listed below Goto:[Stage] %RecoveryImagePath%
ECHO %0 %1 %2
ECHO ==============
Set Scriptroot=%~dp0
Set REimagepath=%2
%1

ECHO =======
ECHO | The step below configures the custom WinRE settings, also another identifier file is being created here. (OS.Tag)
ECHO =======
:EnableRE
cscript "%Scriptroot%\Assign-RE-DriveLetter.vbs"
MD R:\Recovery\WindowsRE\ 
ECHO Windows Recovery Partition > R:\WinRE.Tag
IF %PROCESSOR_ARCHITECTURE% EQU x86 	Copy "%REimagepath%\WinREx86.wim" R:\Recovery\WindowsRE\WinRE.wim
IF %PROCESSOR_ARCHITECTURE% EQU AMD64 	Copy "%REimagepath%\WinREx64.wim" R:\Recovery\WindowsRE\WinRE.wim
IF %ErrorLevel% NEQ 0 Exit %ErrorLevel%
MD C:\Windows\Setup\Scripts
ECHO OS Volume > C:\Windows\OS.Tag
ECHO RD C:\MININT /Q /S > C:\Windows\Setup\Scripts\SetupComplete.cmd
REM ECHO [Replace this with any command you want to execute additionally in the SetupComplete.cmd, also remove "REM" at the start of this line] >> %WinDir%\Setup\Scripts\SetupComplete.cmd
C:\Windows\System32\ReAgentC.exe /Disable
C:\Windows\System32\ReAgentC.exe /SetREImage /Path R:\Recovery\WindowsRE /Target C:\Windows
C:\Windows\System32\ReAgentC.exe /SetOSImage /Customtool
C:\Windows\System32\ReAgentC.exe /Enable
Goto :EOF

ECHO =======
ECHO | The step below does some final configuration and cleanup files and folders that are not needed.
ECHO =======
:Finalize
cscript "%Scriptroot%\Assign-RE-DriveLetter.vbs"
ECHO Select Vol R			>> %Temp%\SetPartitionID.txt
ECHO Set ID 27 Override		>> %Temp%\SetPartitionID.txt
Diskpart /s %Temp%\SetPartitionID.txt

For %%a in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do if exist %%a:\Windows\OS.Tag set WinVol=%%a
Del %WinVol%:\Minint\Unattend.xml /Q
Attrib R:\Bootmgr -h -s -r -a
Del R:\Bootmgr /Q
Copy "%REimagepath%\Windows 7\Bootmgr" R:\Bootmgr
Attrib R:\Bootmgr +h +s +r +a
IF EXIST R:\Sources RD R:\Sources /Q /S
IF EXIST %WinVol%:\Recovery RD %WinVol%:\Recovery /Q /S
Goto :EOF