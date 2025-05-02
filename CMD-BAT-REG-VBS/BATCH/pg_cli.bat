@ECHO ON
SETLOCAL
SET PHPEXE=C:\wamp\bin\php\php5.3.8\php.exe
SET PHPINI=C:\wamp\bin\apache\Apache2.2.21\bin\php.ini
SET Logs=E:\Scripts\PHP_SCRIPTS\LOGS
SET Scripts=E:\Scripts\PHP_SCRIPTS

%PHPEXE% -c %PHPINI% -f %Scripts%\PressGaney.php > %Logs%\PressGaney.Log
%PHPEXE% -c %PHPINI% -f %Scripts%\PressGaneyOMP.php > %Logs%\PressGaneyOMP.log
%PHPEXE% -c %PHPINI% -f %Scripts%\PressGaneyONC_ORTH.php > %Logs%\PressGaneyONC_ORTH.log
%PHPEXE% -c %PHPINI% -f %Scripts%\pg_nofile.php > %Logs%\pg_nofile.log