@ECHO ON
SETLOCAL
SET WinSCP=E:\Scripts\WinSCP\WinSCP.COM
SET Scripts=E:\Scripts\WinSCP\UPDOWN
SET Logs=E:\Scripts\PHP_SCRIPTS\WinSCP\LOGS

%WinSCP% /script=%Scripts%\downloadWaystarERA.txt /log="%Logs%\downloadERA.log"