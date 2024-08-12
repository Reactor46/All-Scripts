
icacls F:\ /inheritance:E /T
icacls F:\ /grant "BUILTIN\Administrators":(F) /T
icacls F:\ /grant "NT AUTHORITY\SYSTEM":(OI)(CI)(F) /T
icacls F:\ /grant "CREATOR OWNER":(OI)(CI)(IO)(F) /T

for /d %%d in (*.*) do icacls %%d /setowner "jbatt\%%d" /T


