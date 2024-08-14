REM net use \\%1 <password> /USER:<username>
shutdown.exe /s /m \\%1 /t 0


REM call this .bat file using the command .\alertScripts\triggerShutdown.bat %client_ip%

REM %1 will be replaced by the ip address of the computer where the event was triggered
REM if the "net use" command is not used, credentials of the current login/service account will be used
REM to run this script with specific user credentials, include the "net use" command as a part of the script and do the following - 
REM replace <password> by the password of user with privileges
REM replace <username> with the name of the respective user
