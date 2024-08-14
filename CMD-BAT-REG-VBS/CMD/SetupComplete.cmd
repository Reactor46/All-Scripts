REM Add the local Administrator account as sysadmin
SqlCmd -s . -d master -A -Q "DROP LOGIN [$(OLDCOMPUTERNAME)\Administrator]"
SqlCmd -s . -d master -A -Q "CREATE LOGIN [%computername%\Administrator] FROM Windows"
SqlCmd -s . -A -Q "EXEC sp_addsrvrolemember @loginame = N'%computername%\Administrator', @rolename = N'sysadmin'"