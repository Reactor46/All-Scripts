net stop spooler
del %SYSTEMROOT%\system32\spool\PRINTERS\*.* /q /s
net start spooler