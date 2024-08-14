@REM =======================================================================================
@REM Read FIRST! 
@REM 	The Script can be used standalone, but to fulfill it´s real purpose, you should
@REM 	download the first part too. The first part reads out the defaults Printer name and 
@REM	should run as a logoff script, the second one (this one) sets a new printer as a 
@REM	default printer and should run as a logon Script. You can donwload the first one here:
@REM	https://skydrive.live.com/redir?resid=4195153C2B1240A8!387
@REM	http://code.google.com/p/cool-scripts/downloads/list
@REM SYNOPSIS
@REM 	Sets the default printer which is stored in a txt file.
@REM DESCRIPTION
@REM 	This script was coded for the printer migration in a large environment. The aim was 
@REM 	to set the default printer for the currently logged in user. But this would be to easy, 
@REM	the difficulty was, that the printers were renamed as part of the migration and so in 
@REM	the first script we read out the defaultprinter, saved the name to a txt file, which 
@REM	is used here, to set up the new defaultprinter. 
@REM Param pfad
@REM	defines the path, where the file will be created
@REM Notes
@REM	ScriptName: set_standard.bat
@REM	Version:    1.1
@REM     Created By: Armin Hilbert
@REM     Date Coded: November 12, 2012
@REM =======================================================================================
@echo off

@REM =======================================================================================
@REM Param to set the working path
@REM =======================================================================================
SET pfad="C:\Users\%Username%\Desktop\"

@REM =======================================================================================
@REM Param to read out the name of the new defaultprinter, which is stored in this file
@REM =======================================================================================
SET /p newDrucker=<%pfad%%USERNAME%_new_defaultprinter.txt

setlocal enabledelayedexpansion

@REM =======================================================================================
@REM search installed printing devices, if something is found with the wanted name, this one
@REM will be the new default printer.
@REM =======================================================================================
for /f "delims=" %%i in ('reg query "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Devices" ^| findstr "%newDrucker%"') do (
rundll32 printui.dll,PrintUIEntry /y /n "%newDrucker%"
)

