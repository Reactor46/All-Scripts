@echo off

:: Created by: Shawn Brink
:: Created on: August 27th 2016
:: Tutorial: http://www.tenforums.com/tutorials/61731-network-icon-lock-sign-screen-add-remove-windows-10-a.html


REG Delete "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /V DontDisplayNetworkSelectionUI /F

taskkill /f /im explorer.exe
start explorer.exe