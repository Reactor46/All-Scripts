SET OU2=SharePoint Managed Accounts 
SET OU1=Development 
SET DOMAIN=CORP 
SET DOMAINSUFFIX=com 
SET PWD=dev@word1 
 
 
rem THIS WILL ADD OU:s TO THE ACTIVE DIRECTORY 
 
dsadd ou "ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" 
dsadd ou "ou=User Accounts, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" 
dsadd ou "ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" 
dsadd ou "ou=SharePocom Servers, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" 
 
 
rem DEFINE USER NAMES 
 
SET ACC01=SVC_SP_Farm 
SET ACC02=SVC_SP_DefAppPool 
SET ACC03=SVC_SP_Crawl 
SET ACC04=SVC_SP_Search 
SET ACC05=SVC_SP_MySitePool 
SET ACC06=SVC_SP_Admin 
SET ACC07=SVC_SP_SAExcel 
SET ACC08=SVC_SP_SADefault 
SET ACC09=SVC_SP_SASearchAdm 
SET ACC10=SVC_SP_SASearchQuery 
SET ACC11=SVC_SP_SASecureStore 
SET ACC12=SVC_SP_SAPerfPoint 
SET ACC13=SVC_SP_SAPerfP_Un 
SET ACC14=SVC_SP_SAUserProfile 
SET ACC15=SVC_SP_SA_UP_Sync 
SET ACC16=SVC_SQL_Service 
SET ACC17=SVC_SQL_Report 
SET ACC18=SVC_SQL_Agent 
SET ACC19=SVC_SQL_Analysis 
SET ACC20=SVC_SQL_Integration 
SET ACC21=SVC_SP_SuperUser 
SET ACC22=SVC_SP_SuperReader 
SET ACC25=SVC_SP_VisioService 
 
 
rem ADD USERS 
 
dsadd user "cn=%ACC01%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC01%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint farm account" >>UserCreationLog.txt 
dsadd user "cn=%ACC02%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC02%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint default application pool account" >>UserCreationLog.txt 
dsadd user "cn=%ACC03%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC03%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint crawling account" >>UserCreationLog.txt 
dsadd user "cn=%ACC04%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC04%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint search account" >>UserCreationLog.txt 
dsadd user "cn=%ACC05%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC05%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint MySite application pool account" >>UserCreationLog.txt 
dsadd user "cn=%ACC06%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC06%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint administration account" >>UserCreationLog.txt 
dsadd user "cn=%ACC07%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC07%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint service application account" >>UserCreationLog.txt 
dsadd user "cn=%ACC08%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC08%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint SA default application pool account" >>UserCreationLog.txt 
dsadd user "cn=%ACC09%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC09%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint SA Search Admin application pool account" >>UserCreationLog.txt 
dsadd user "cn=%ACC10%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC10%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint SA Search Query application pool account" >>UserCreationLog.txt 
dsadd user "cn=%ACC11%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC11%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint SA Secure Store application pool account" >>UserCreationLog.txt 
dsadd user "cn=%ACC12%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC12%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint SA Performance Point application pool account" >>UserCreationLog.txt 
dsadd user "cn=%ACC13%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC13%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint SA Performance Point unattended  account (external services)" >>UserCreationLog.txt 
dsadd user "cn=%ACC14%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC14%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint SA User Profile application pool account" >>UserCreationLog.txt 
dsadd user "cn=%ACC15%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC15%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint SA User Profile sync account" >>UserCreationLog.txt 
dsadd user "cn=%ACC16%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC16%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SQL Server account" >>UserCreationLog.txt 
dsadd user "cn=%ACC17%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC17%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SQL Server Reporting Services account" >>UserCreationLog.txt 
dsadd user "cn=%ACC18%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC18%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SQL Server Agent account" >>UserCreationLog.txt 
dsadd user "cn=%ACC19%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC19%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SQL Server Analysis account" >>UserCreationLog.txt 
dsadd user "cn=%ACC20%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC20%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SQL Server Integration Services account" >>UserCreationLog.txt 
dsadd user "cn=%ACC21%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC21%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint super user account (object cache)" >>UserCreationLog.txt 
dsadd user "cn=%ACC22%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC22%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint super reader account (object cache)" >>UserCreationLog.txt 
dsadd user "cn=%ACC25%, ou=%OU2%, ou=%OU1%, dc=%DOMAIN%, dc=%DOMAINSUFFIX%" -upn %ACC25%@%DOMAIN%.%DOMAINSUFFIX% -pwd %PWD% -pwdneverexpires yes -mustchpwd no -desc "SharePoint Visio services" >>UserCreationLog.txt