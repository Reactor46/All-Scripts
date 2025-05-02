@ECHO OFF

REM *********************************************************************
REM Script Title : EnablePowerShellScripts.cmd
REM
REM Author	 : Shawn Gibbs
REM
REM Description  : Enable powershell ExecutionPolicy on system. 
REM                This Allows PowerShell scripts to run locally.
REM
REM *********************************************************************

ECHO Enabling PowerShell Script Execution. Set-ExecutionPolicy RemoteSigned >> %windir%\temp\Output.log
%windir%\system32\windowspowershell\v1.0\powershell.exe -noprofile Set-ExecutionPolicy RemoteSigned >> %windir%\temp\Output.log
ECHO. >> %windir%\temp\Output.log

EXIT 0