@ECHO ON
D:
CD D:\D-Drive\User

FOR /F "tokens=*" %%G IN ('dir /b') DO takeown /f "BUILTIN\Administrators" /r /d y 
FOR /F "tokens=*" %%G IN ('dir /b') DO icacls %%G /grant administrators:F /T
FOR /F "tokens=*" %%G IN ('dir /b') DO icacls %%G /setowner %G /T
FOR /F "tokens=*" %%G IN ('dir /b') DO icacls %%G /inheritance:E /T
FOR /F "tokens=*" %%G IN ('dir /b') DO icacls %%G /grant %%G:(OI)(CI)F /T