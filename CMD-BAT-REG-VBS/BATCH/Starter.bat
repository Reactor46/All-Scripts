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

@echo off
rem ***set the background color.***
color 2e
rem ***set the title name.***
title Remove Notes Pages in PowerPoint
cls

rem ***set the welcome information.***
@echo ¨X©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤¨[
@echo ©¦  Welcome to One Script   ©¦
@echo ¨^©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤¨a
@echo ¨X©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤¨[
@echo ©¦ Warning: Removing Notes Pages using script is irreversible and user©¦
@echo ©¦ should take a back of presentation files before running script.    ©¦
@echo ¨^©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤¨a

rem ***press any key to continue.***
pause

rem ***refresh the command line.***
cls

rem ***set the welcome information.***
@echo ¨X©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤¨[
@echo ©¦  Welcome to One Script   ©¦
@echo ¨^©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤¨a

rem ***execute the VBScript file.***
cscript //nologo RemoveNotes.vbs

rem ***press any key to continue.***
pause