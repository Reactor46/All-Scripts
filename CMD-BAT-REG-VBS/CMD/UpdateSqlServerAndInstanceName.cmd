if "%1"=="" goto Usage

SqlCmd -s . -d Master -A -Q "select @@servername as 'Before Sql Server Name Change'" >> C:\Scripts\SqlServerNameChange.log

SqlCmd -s . -d Master -A -Q "EXEC sp_dropserver '%1'"

SqlCmd -s . -d Master -A -Q "EXEC sp_dropserver '%1\SQLEXPRESS'"

SqlCmd -s . -d Master -A -Q "EXEC sp_addserver '$(NEWCOMPUTERNAME)', local"

SqlCmd -s . -d Master -A -Q "EXEC sp_addserver '$(NEWCOMPUTERNAME)\SQLEXPRESS', local"

ping localhost -n 30

net stop BTSSvc$BizTalkServerApplication

REM - Unregister Bam Alerts (SQL Server 2008 R2)
REM net stop NS$BAMAlerts
REM taskkill /f /im nsservice.exe
REM cd "%ProgramFiles%\Microsoft SQL Server\90\NotificationServices\9.0.242\bin\"
REM nscontrol unregister -name BamAlerts



REM - Stop BAMAlerts (SQL Server 2012)

net stop BAMAlerts

net stop ENTSSO
taskkill /f /im entsso.exe

net stop SQLAgent$SQLEXPRESS
net stop sqlserveragent

net stop ReportServer$SQLEXPRESS
net stop ReportServer
taskkill /f /im ReportingServicesService*

net stop MSOLAP$SQLEXPRESS
net stop MSSQLServerOLAPService

net stop MSSQL$SQLEXPRESS
net stop mssqlserver

net start mssqlserver
net start MSSQL$SQLEXPRESS

net start ReportServer$SQLEXPRESS
net start ReportServer

net start MSOLAP$SQLEXPRESS
net start MSSQLServerOLAPService

net start sqlserveragent
net start SQLAgent$SQLEXPRESS

net start ENTSSO

REM - Reregister Bam Alerts service (SQL Server 2008 R2)
REM nscontrol register -name BamAlerts -server $(NEWCOMPUTERNAME) -service -serviceusername "s$(BIZTALKSERVERSERVICEACCOUNT)" -servicepassword "$(BIZTALKSERVERSERVICEACCOUNTPASSWORD)"
REM net start NS$BAMAlerts



REM Start BAMAlerts (SQL Server 2012)

net start BAMAlerts

net start BTSSvc$BizTalkServerApplication

REM Drop and Re-create the BAMPrimaryImportSrv Linked Server if it exists
SqlCmd -s . -d Master -A -Q "IF  EXISTS (SELECT srv.name FROM sys.servers srv WHERE srv.server_id != 0 AND srv.name = N'BAMPrimaryImportSrv') BEGIN EXEC master.dbo.sp_dropserver @server=N'BAMPrimaryImportSrv', @droplogins='droplogins'; EXEC sp_addlinkedserver @server=N'BAMPrimaryImportSrv', @srvproduct=N'', @provider=N'SQLNCLI', @datasrc=N'(local)'; END" >> C:\Scripts\SqlServerUpdateLinkedServer.log

SqlCmd -s . -d Master -A -Q "select @@servername as 'After Sql Server Name Change'" >> C:\Scripts\SqlServerNameChange.log

:Usage
echo "Updates SQL after a machine name change and cycles a number of dependent services"
echo "Usage: UpdateSqlServerAndInstanceName <oldDBServerName>"
