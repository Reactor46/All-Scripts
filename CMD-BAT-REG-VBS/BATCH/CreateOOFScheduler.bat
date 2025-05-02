@echo off

goto start
---------------------------------------------------------------------------------
 The sample scripts are not supported under any Microsoft standard support
 program or service. The sample scripts are provided AS IS without warranty
 of any kind. Microsoft further disclaims all implied warranties including,
 without limitation, any implied warranties of merchantability or of fitness for
 a particular purpose. The entire risk arising out of the use or performance of
 the sample scripts and documentation remains with you. In no event shall
 Microsoft, its authors, or anyone else involved in the creation, production, or
 delivery of the scripts be liable for any damages whatsoever (including,
 without limitation, damages for loss of business profits, business interruption,
 loss of business information, or other pecuniary loss) arising out of the use
 of or inability to use the sample scripts or documentation, even if Microsoft
 has been advised of the possibility of such damages.
---------------------------------------------------------------------------------
:start

rem ------------------------------------------
rem Make sure the Schedule service is started.
rem ------------------------------------------
sc config Schedule start= auto
net start Schedule

rem ---------------------------------------------
rem Add a time range for Out of Office Assistant.
rem ---------------------------------------------
schtasks /create /tn OOFStartTime /tr "cscript //nologo D:\OOF.vbs True" /sc weekly /d MON,TUE,WED,THU,FRI /st 12:00:00
schtasks /create /tn OOFEndTime /tr "cscript //nologo D:\OOF.vbs False" /sc weekly /d MON,TUE,WED,THU,FRI /st 12:30:00

pause