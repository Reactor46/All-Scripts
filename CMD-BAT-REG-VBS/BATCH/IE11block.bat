for /f %%C in (C:\Temp\computers.txt) do  C:\pstools\psexec \\%%C REG ADD "HKLM\SOFTWARE\Microsoft\Internet Explorer\Setup\11.0" /v DoNotAllowIE11 /t REG_DWORD /d 1 /f