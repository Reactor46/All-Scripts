REM UpdateSqlSecurity.cmd

if "%1"=="" goto Usage

SET oldcomputername=%1

REM Add the local Administrator account as sysadmin
REM SqlCmd -s . -d master -A -Q "DROP LOGIN [%oldcomputername%\Administrator]"
REM SqlCmd -s . -d master -A -Q "CREATE LOGIN [$(NEWCOMPUTERNAME)\Administrator] FROM Windows"
REM SqlCmd -s . -A -Q "EXEC sp_addsrvrolemember @loginame = N'$(NEWCOMPUTERNAME)\Administrator', @rolename = N'sysadmin'"

REM Re-create the logins with the new computer name
SqlCmd -s . -d master -A -Q "DROP LOGIN [%oldcomputername%\BizTalk Application Users]"
SqlCmd -s . -d master -A -Q "CREATE LOGIN [$(NEWCOMPUTERNAME)\BizTalk Application Users] FROM Windows"

SqlCmd -s . -d master -A -Q "DROP LOGIN [%oldcomputername%\BizTalk Isolated Host Users]"
SqlCmd -s . -d master -A -Q "CREATE LOGIN [$(NEWCOMPUTERNAME)\BizTalk Isolated Host Users] FROM Windows"

SqlCmd -s . -d master -A -Q "DROP LOGIN [%oldcomputername%\BizTalk Server Administrators]"
SqlCmd -s . -d master -A -Q "CREATE LOGIN [$(NEWCOMPUTERNAME)\BizTalk Server Administrators] FROM Windows"

SqlCmd -s . -d master -A -Q "DROP LOGIN [%oldcomputername%\BizTalk Server B2B Operators]"
SqlCmd -s . -d master -A -Q "CREATE LOGIN [$(NEWCOMPUTERNAME)\BizTalk Server B2B Operators] FROM Windows"

SqlCmd -s . -d master -A -Q "DROP LOGIN [%oldcomputername%\BizTalk Server Operators]"
SqlCmd -s . -d master -A -Q "CREATE LOGIN [$(NEWCOMPUTERNAME)\BizTalk Server Operators] FROM Windows"

SqlCmd -s . -d master -A -Q "DROP LOGIN [%oldcomputername%\$(BIZTALKSERVERSERVICEACCOUNT)]"
SqlCmd -s . -d master -A -Q "CREATE LOGIN [$(NEWCOMPUTERNAME)\$(BIZTALKSERVERSERVICEACCOUNT)] FROM Windows"

SqlCmd -s . -d master -A -Q "DROP LOGIN [%oldcomputername%\$(SQLSERVERSERVICEACCOUNT)]"
SqlCmd -s . -d master -A -Q "CREATE LOGIN [$(NEWCOMPUTERNAME)\$(SQLSERVERSERVICEACCOUNT)] FROM Windows"

SqlCmd -s . -d master -A -Q "DROP LOGIN [%oldcomputername%\SSO Administrators]"
SqlCmd -s . -d master -A -Q "CREATE LOGIN [$(NEWCOMPUTERNAME)\SSO Administrators] FROM Windows"

REM Re-map the users to the new logins

REM BAMAlertsApplication
SqlCmd -s . -d BAMAlertsApplication -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Administrators]"
SqlCmd -s . -d BAMAlertsApplication -A -Q "ALTER USER [%oldcomputername%\$(BIZTALKSERVERSERVICEACCOUNT)] WITH LOGIN = [$(NEWCOMPUTERNAME)\$(BIZTALKSERVERSERVICEACCOUNT)]"

REM BAMAlertsNSMain
SqlCmd -s . -d BAMAlertsNSMain -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Administrators]"
SqlCmd -s . -d BAMAlertsNSMain -A -Q "ALTER USER [%oldcomputername%\$(BIZTALKSERVERSERVICEACCOUNT)] WITH LOGIN = [$(NEWCOMPUTERNAME)\$(BIZTALKSERVERSERVICEACCOUNT)]"

REM BAMArchive
SqlCmd -s . -d BAMArchive -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Administrators]"

REM BAMPrimaryImport
SqlCmd -s . -d BAMPrimaryImport -A -Q "ALTER USER [%oldcomputername%\BizTalk Application Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Application Users]"
SqlCmd -s . -d BAMPrimaryImport -A -Q "ALTER USER [%oldcomputername%\BizTalk Isolated Host Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Isolated Host Users]"
SqlCmd -s . -d BAMPrimaryImport -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Administrators]"
SqlCmd -s . -d BAMPrimaryImport -A -Q "ALTER USER [%oldcomputername%\BizTalk Server B2B Operators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server B2B Operators]"
SqlCmd -s . -d BAMPrimaryImport -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Operators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Operators]"
SqlCmd -s . -d BAMPrimaryImport -A -Q "ALTER USER [%oldcomputername%\$(BIZTALKSERVERSERVICEACCOUNT)] WITH LOGIN = [$(NEWCOMPUTERNAME)\$(BIZTALKSERVERSERVICEACCOUNT)]"
SqlCmd -s . -d BAMPrimaryImport -A -Q "sp_changedbowner '$(NEWCOMPUTERNAME)\Administrator'"

REM BAMStarSchema
SqlCmd -s . -d BAMStarSchema -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Administrators]"
SqlCmd -s . -d BAMStarSchema -A -Q "ALTER USER [%oldcomputername%\$(SQLSERVERSERVICEACCOUNT)] WITH LOGIN = [$(NEWCOMPUTERNAME)\$(SQLSERVERSERVICEACCOUNT)]"

REM BizTalkDTADb
SqlCmd -s . -d BizTalkDTADb -A -Q "ALTER USER [%oldcomputername%\BizTalk Application Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Application Users]"
SqlCmd -s . -d BizTalkDTADb -A -Q "ALTER USER [%oldcomputername%\BizTalk Isolated Host Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Isolated Host Users]"
SqlCmd -s . -d BizTalkDTADb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Administrators]"
SqlCmd -s . -d BizTalkDTADb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server B2B Operators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server B2B Operators]"
SqlCmd -s . -d BizTalkDTADb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Operators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Operators]"

REM BizTalkMgmtDb
SqlCmd -s . -d BizTalkMgmtDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Application Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Application Users]"
SqlCmd -s . -d BizTalkMgmtDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Isolated Host Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Isolated Host Users]"
SqlCmd -s . -d BizTalkMgmtDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Administrators]"
SqlCmd -s . -d BizTalkMgmtDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server B2B Operators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server B2B Operators]"
SqlCmd -s . -d BizTalkMgmtDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Operators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Operators]"

REM BizTalkMsgBoxDb
SqlCmd -s . -d BizTalkMsgBoxDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Application Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Application Users]"
SqlCmd -s . -d BizTalkMsgBoxDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Isolated Host Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Isolated Host Users]"
SqlCmd -s . -d BizTalkMsgBoxDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Administrators]"
SqlCmd -s . -d BizTalkMsgBoxDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server B2B Operators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server B2B Operators]"
SqlCmd -s . -d BizTalkMsgBoxDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Operators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Operators]"

REM BizTalkRuleEngineDb
SqlCmd -s . -d BizTalkRuleEngineDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Application Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Application Users]"
SqlCmd -s . -d BizTalkRuleEngineDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Isolated Host Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Isolated Host Users]"
SqlCmd -s . -d BizTalkRuleEngineDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Administrators]"
SqlCmd -s . -d BizTalkRuleEngineDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server B2B Operators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server B2B Operators]"
SqlCmd -s . -d BizTalkRuleEngineDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Operators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Operators]"
SqlCmd -s . -d BizTalkRuleEngineDb -A -Q "ALTER USER [%oldcomputername%\$(BIZTALKSERVERSERVICEACCOUNT)] WITH LOGIN = [$(NEWCOMPUTERNAME)\$(BIZTALKSERVERSERVICEACCOUNT)]"

REM ESBAdmin
SqlCmd -s . -d ESBAdmin -A -Q "ALTER USER [%oldcomputername%\BizTalk Application Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Application Users]"
SqlCmd -s . -d ESBAdmin -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Administrators]"

REM EsbExceptionDb
SqlCmd -s . -d EsbExceptionDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Application Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Application Users]"
SqlCmd -s . -d EsbExceptionDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Administrators]"

REM EsbItineraryDb
SqlCmd -s . -d EsbItineraryDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Application Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Application Users]"
SqlCmd -s . -d EsbItineraryDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Isolated Host Users] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Isolated Host Users]"
SqlCmd -s . -d EsbItineraryDb -A -Q "ALTER USER [%oldcomputername%\BizTalk Server Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\BizTalk Server Administrators]"

REM ReportServer
SqlCmd -s . -d ReportServer -A -Q "ALTER USER [%oldcomputername%\$(SQLSERVERSERVICEACCOUNT)] WITH LOGIN = [$(NEWCOMPUTERNAME)\$(SQLSERVERSERVICEACCOUNT)]"

REM ReportServerTempDB
SqlCmd -s . -d ReportServerTempDB -A -Q "ALTER USER [%oldcomputername%\$(SQLSERVERSERVICEACCOUNT)] WITH LOGIN = [$(NEWCOMPUTERNAME)\$(SQLSERVERSERVICEACCOUNT)]"

REM SSODB
SqlCmd -s . -d SSODB -A -Q "ALTER USER [%oldcomputername%\SSO Administrators] WITH LOGIN = [$(NEWCOMPUTERNAME)\SSO Administrators]"

:Usage
echo "Updates SQL Logins and Users after a machine name change"
echo "Usage: UpdateSqlSecurity <oldDBServerName>"