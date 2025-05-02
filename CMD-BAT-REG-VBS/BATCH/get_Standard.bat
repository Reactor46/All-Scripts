@REM =======================================================================================
@REM Read FIRST! 
@REM 	The Script can be used standalone, but to fulfill it´s real purpose, you should
@REM 	download the second part. The first part reads out the defaults Printer name and 
@REM	should run as a logoff script, the second one sets a new printer as a default 
@REM	printer and should run as a logon Script. You can donwload it here:
@REM	https://skydrive.live.com/redir?resid=4195153C2B1240A8!387
@REM	http://code.google.com/p/cool-scripts/downloads/list
@REM SYNOPSIS
@REM 	Read out the default printer name and save it to a txt File.
@REM DESCRIPTION
@REM 	This script was coded for the printer migration in a large environment. The aim was 
@REM 	to read out the default printer name of the User who is logged in and save it 
@REM 	modified to a txt File. The modification is necessary, cause as part of the 
@REM	migration all printer got a new name and the Users required a valid default printer. 
@REM 	After this, a second txt file will be created to save the old default Printername, 
@REM	just to have a backup.
@REM Param pfad
@REM	defines the path, where the file will be created
@REM Notes
@REM	ScriptName: get_standard.bat
@REM	Version:    1.1
@REM     Created By: Armin Hilbert
@REM     Date Coded: November 12, 2012
@REM =======================================================================================
@echo off

@REM =======================================================================================
@REM Param to set the working path
@REM =======================================================================================
SET pfad="C:\Users\%Username%\Desktop\"

wmic printer where "Default = 'True'" get Name > %pfad%old_defaultprinter.txt

setlocal enabledelayedexpansion

@REM =======================================================================================
@REM Enter here the String which should be replaced with a new one.
@REM =======================================================================================
Set "FindString=Canon"

@REM =======================================================================================
@REM Enter here the String which replace the old one.
@REM =======================================================================================
Set "ReplaceWith=Brother"

@REM =======================================================================================
@REM Replacing Process
@REM =======================================================================================
For /f "delims=" %%a in ( 
   'find.exe /n /v ""^<"%pfad%old_defaultprinter.txt"'
   ) Do (
   (set Line="''%%a")
   (Set Line=!Line:^<=^^^<!) & (Set Line=!Line:^>=^^^>!)
   call:replace
   (echo\!Line!)>>"%pfad%new.txt"
)

@REM =======================================================================================
@REM Creation of the File with the new name in it
@REM =======================================================================================
type %pfad%new.txt|find /v "Name" > %pfad%%Username%_new_defaultprinter.txt

wmic printer where "Default = 'True'" get Name > %pfad%old.txt

@REM =======================================================================================
@REM Creation of the File with the old name in it
@REM =======================================================================================
type %pfad%old.txt|find /v "Name" > %pfad%%Username%_old_defaultprinter.txt

@REM =======================================================================================
@REM deletion of the temporary Files
@REM =======================================================================================
del %pfad%old_defaultprinter.txt
del %pfad%new.txt
del %pfad%old.txt

@REM =======================================================================================
@REM Replacing algorithm
@REM =======================================================================================
goto:eof ------------
:replace  subroutine
Set/a i = %i% + 1
(Set Line=!Line:''[%i%]=!)
Set "FindString2="
Set "FindString2=%FindString:<=^^^<%"
Set "FindString2=%FindString2:>=^^^>%"
Set "ReplaceWith2="
Set "ReplaceWith2=%ReplaceWith:<=^^^<%"
Set "ReplaceWith2=%ReplaceWith2:>=^^^>%"
(Set Line=!Line:%FindString2%=%ReplaceWith2%!)
(Set Line=%Line:~1,-1%)
goto:eof ------------