REG ADD "HKEY_USERS\.DEFAULT\Control Panel\Desktop" /V HungAppTimeout /T REG_SZ /D 1000 /F
taskkill /f /im explorer.exe
start explorer.exe