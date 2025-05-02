@ECHO ON
SETLOCAL
SET PHPEXE=C:\wamp\bin\php\php5.3.8\php.exe
SET PHPINI=C:\wamp\bin\apache\Apache2.2.21\bin\php.ini
SET SCRIPTS=E:\Scripts\PHP_SCRIPTS
SET Logs=E:\Scripts\PHP_SCRIPTS\LOGS

%PHPEXE% -c %PHPINI% -f %REPORTS%\apcreport.php > %Logs%\uson_apc.log
%PHPEXE% -c %PHPINI% -f %REPORTS%\apcreport2.php > %Logs%\omp_apc.log
%PHPEXE% -c %PHPINI% -f %REPORTS%\apcreportONC.php > %Logs%\onc_apc.log
%PHPEXE% -c %PHPINI% -f %REPORTS%\apcreportOrtho.php > %Logs%\ortho_apc.log