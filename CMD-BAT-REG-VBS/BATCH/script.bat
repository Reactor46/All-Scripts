@echo OFF

@setlocal enableextensions
@cd /d "%~dp0"

set outputRoot=%cd%\Output
md %outputRoot% > nul 2>&1
set templatesRoot=%cd%\Templates

:templateLabel
set /p template="Enter template (eg. 'Application'): "
if not exist %templatesRoot%\%template% (
	echo Template does not exist, try again!
	goto templateLabel
) 
:durationLabel
set /p duration="Enter duration (sec): "
echo %duration%| findstr /r "^[1-9][0-9]*$">nul
if errorlevel 1 (
	echo Duration value is not valid, try again!
	goto durationLabel
)
:sampleIntervalLabel
set /p sampleInterval="Enter sample interval (sec): "
echo %sampleInterval%| findstr /r "^[1-9][0-9]*$">nul
if errorlevel 1 (
	echo Sample interval value is not valid, try again!
	goto sampleIntervalLabel
)

set folderName=%duration%_%sampleInterval%

md %outputRoot%\%template%_%folderName% > nul 2>&1

set startFile=%outputRoot%\%template%_%folderName%\start.bat
break > %startFile%
set stopFile=%outputRoot%\%template%_%folderName%\stop.bat
break > %stopFile%
set deleteFile=%outputRoot%\%template%_%folderName%\delete_groups.bat
break > %deleteFile%
set importFile=%outputRoot%\%template%_%folderName%\import_groups.bat
echo @setlocal enableextensions > %importFile%
echo @cd /d "%outputRoot%\%template%_%folderName%\" >> %importFile%

setlocal enabledelayedexpansion

for /f %%f in ('dir /b %templatesRoot%\%template%') do (

	set outputFile=%outputRoot%\%template%_%folderName%\%%f
	break > !outputFile!
	echo !outputFile! created
	
	for /F "delims=" %%G in (%templatesRoot%\%template%\%%f) do (
		set LINE=%%G
		set LINE=!LINE:^<Duration^>#####^</Duration^>=^<Duration^>%duration%^</Duration^>!
		set LINE=!LINE:^<SampleInterval^>#####^</SampleInterval^>=^<SampleInterval^>%sampleInterval%^</SampleInterval^>!
		echo !LINE! >> !outputFile!
	) 
	
	logman delete %folderName%.%%~nf > nul
	logman import -name %folderName%.%%~nf -xml !outputFile!
	
	echo logman start %folderName%.%%~nf >> !startFile!
	echo logman stop %folderName%.%%~nf >> !stopFile!
	echo logman delete %folderName%.%%~nf >> !deleteFile!
	echo logman import -name %folderName%.%%~nf -xml %%f >> !importFile!
)

endlocal