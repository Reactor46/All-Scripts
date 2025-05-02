@ECHO ON
SETLOCAL
SET WinSCP=E:\Scripts\WinSCP\WinSCP.COM
SET Scripts=E:\Scripts\WinSCP\UPDOWN
SET Logs=E:\Scripts\WinSCP\LOGS

%WinSCP% %Scripts%\pg_upload.txt > %Logs%\PG2.log
if ERRORLEVEL 1  GOTO :ERROR
move E:\Scripts\REPORTS\PG\*.csv E:\Scripts\REPORTS\PG\Archive

%WinSCP% %Scripts%\pg_omp_upload.txt > %Logs%\PG_OMP2.log
if ERRORLEVEL 1  GOTO :ERROR2
move E:\Scripts\REPORTS\PG_OMP\*.csv E:\Scripts\REPORTS\PG_OMP\Archive

%WinSCP% %Scripts%\pg_az_upload.txt > %Logs%\PG_AZ.log
if ERRORLEVEL 1  GOTO :ERROR2
move E:\Scripts\REPORTS\PG_AZ\*.csv E:\Scripts\REPORTS\PG_AZ\Archive

%WinSCP% %Scripts%\pg_Allscripts.txt > %Logs%\pg_Allscripts.log
if ERRORLEVEL 1  GOTO :ERROR2
move E:\Scripts\REPORTSPG_AllScripts\*.csv E:\Scripts\REPORTS\PG_AllScripts\Archive

%WinSCP% %Scripts%\pg_ONC_ORTH_upload.txt > %Logs%\PG_ONC_ORTH.log
if ERRORLEVEL 1  GOTO :ERROR2
move E:\Scripts\REPORTS\PG_ONC_ORTH\*.csv E:\Scripts\REPORTS\PG_ONC_ORTH\Archive

ENDLOCAL
GOTO :EOF
:ERROR
	GOTO :EOF

:ERROR2
	GOTO :EOF

EXIT /B 1 