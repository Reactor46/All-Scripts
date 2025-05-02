@ECHO ON
D:
CD D:\D-Drive\User

REM FOR /F "tokens=*" %G IN ('dir /b') DO icacls %G /setowner %G /T
REM FOR /F "tokens=*" %G IN ('dir /b') DO icacls %G /inheritance:E /T
REM FOR /F "tokens=*" %G IN ('dir /b') DO icacls %G /grant %G:(OI)(CI)F /T

