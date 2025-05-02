REM UpdateBAM.cmd

if "%1"=="" goto Usage

"%programfiles(x86)%\Microsoft BizTalk Server 2013\Tracking\bm.exe" update-config -FileName:%1

:Usage
echo "Re-configures BAM using a provided xml answer file"
echo "Usage: UpdateBAM.cmd <BamConfig.xml>"