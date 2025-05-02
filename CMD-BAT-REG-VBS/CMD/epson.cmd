@ECHO ON
:EPSON NETWORK
cls
title Epson Network Printer
echo Installing Epson Printer
ipconfig | find /i "192.168.94"
if NOT errorlevel 1
start /wait %CD%\Apps\epson\epson.exe /s /f1 "%CD%\Apps\epson\94.35.inf"
	ELSE
ipconfig | find /i "192.168.95"
if NOT errorlevel 1
start /wait %CD%\Apps\epson\epson.exe /s /f1 "%CD%\Apps\epson\95.35.inf"
	ELSE
ipconfig | find /i "192.168.96"
if NOT errorlevel 1
start /wait %CD%\Apps\epson\epson.exe /s /f1 "%CD%\Apps\epson\96.35.inf"
	ELSE
ipconfig | find /i "192.168.97"
if NOT errorlevel 1
start /wait %CD%\Apps\epson\epson.exe /s /f1 "%CD%\Apps\epson\97.35.inf"
	ELSE
ipconfig | find /i "192.168.109"
if NOT errorlevel 1
start /wait %CD%\Apps\epson\epson.exe /s /f1 "%CD%\Apps\epson\109.35.inf"
	ELSE
ipconfig | find /i "192.168.205"
if NOT errorlevel 1
start /wait %CD%\Apps\epson\epson.exe /s /f1 "%CD%\Apps\epson\205.35.inf"
	ELSE