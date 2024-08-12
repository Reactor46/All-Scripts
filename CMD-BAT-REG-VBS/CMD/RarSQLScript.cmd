@ECHO OFF
::ECHO ON
SETLOCAL ENABLEDELAYEDEXPANSION
SET MYDIR="D:\Backup\"
FOR /r %MYDIR% %%A IN (*.sql) DO (
    ECHO %%A
	SET MYVAR=%%~nA.rar
	CD /D "%%~dpA" 				  
	RAR A "!MYVAR!" "%%~nxA"      
	IF EXIST !MYVAR! DEL "%%~nxA" 
)
PAUSE