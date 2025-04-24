@echo on

D:

REM 


REM icacls D:\D-Drive\User /inheritance:E /T
REM icacls D:\D-Drive\User /grant "BUILTIN\Administrators":(F) /T
REM icacls D:\D-Drive\User /grant "FNBMCORP\Domain Admins":(OI)(CI)(F) /T
REM icacls D:\D-Drive\User /grant "NT AUTHORITY\SYSTEM":(OI)(CI)(F) /T
REM icacls D:\D-Drive\User /grant "CREATOR OWNER":(OI)(CI)(IO)(F) /T
REM icacls D:\D-Drive\User /grant "FNBMCORP\svc_uniflow":(I)(OI)(CI)(M) /T

REM for /d %%d in (*.*) do icacls %%d /setowner "FNBMCORP\%%d" /T

REM for /d %%d in (*.*) do icacls %%d /reset /t

icacls D:\WebApplications /remove "S-1-5-21-789336058-1085031214-725345543-23869" /T /C



C:
CD C:\

icacls C:\inetpub /setowner "NT Service\TrustedInstaller" /T
icacls 