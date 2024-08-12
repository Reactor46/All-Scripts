@ECHO ON
SETLOCAL
SET PHPEXE=C:\wamp\bin\php\php5.3.8\php.exe
SET PHPINI=C:\wamp\bin\apache\Apache2.2.21\bin\php.ini
SET REPORTS=C:\wamp\www\reports
SET Logs=E:\Scripts\PHP_SCRIPTS\LOGS

%PHPEXE% -c %PHPINI% -f "%REPORTS%\Purple Box\errors.php"
