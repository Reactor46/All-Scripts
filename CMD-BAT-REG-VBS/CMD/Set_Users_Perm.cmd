@echo on

D:

REM 
CD D:\D-Drive\

REM icacls D:\D-Drive\User /inheritance:E /T
REM icacls D:\D-Drive\User /grant "BUILTIN\Administrators":(F) /T
REM icacls D:\D-Drive\User /grant "FNBMCORP\Domain Admins":(OI)(CI)(F) /T
REM icacls D:\D-Drive\User /grant "NT AUTHORITY\SYSTEM":(OI)(CI)(F) /T
REM icacls D:\D-Drive\User /grant "CREATOR OWNER":(OI)(CI)(IO)(F) /T
REM icacls D:\D-Drive\User /grant "FNBMCORP\svc_uniflow":(I)(OI)(CI)(M) /T

CD D:\D-Drive\User

REM for /d %%d in (*.*) do icacls %%d /setowner "FNBMCORP\%%d" /T


REM for /d %%d in (*.*) do icacls %%d /reset /t
