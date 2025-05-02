:: Created by: Shawn Brink
:: http://www.tenforums.com
:: Tutorial: http://www.tenforums.com/tutorials/20616-select-items-using-check-boxes-turn-off-windows-10-a.html


REG ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /V AutoCheckSelect /T REG_DWORD /D 1 /F

taskkill /f /im explorer.exe
start explorer.exe