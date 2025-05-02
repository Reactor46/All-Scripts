@ECHO ON
SETLOCAL
SET PHPEXE=C:\wamp\bin\php\php5.3.8\php.exe
SET PHPINI=C:\wamp\bin\apache\Apache2.2.21\bin\php.ini
SET Logs=E:\Scripts\PHP_SCRIPTS\LOGS
SET Scripts=E:\Scripts\PHP_SCRIPTS

%PHPEXE% -c %PHPINI% -f %Scripts%\referrals.php > %Logs%\referrals.log