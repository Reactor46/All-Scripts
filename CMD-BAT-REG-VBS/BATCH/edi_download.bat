@ECHO ON
SETLOCAL
SET WinSCP=E:\Scripts\WinSCP\WinSCP.COM
SET Scripts=E:\Scripts\WinSCP\UPDOWN
SET Logs=E:\Scripts\WinSCP\LOGS

%WinSCP% /script=%Scripts%\edi_omg_download.txt > %Logs%\edi_omg.log
%WinSCP% /script=%Scripts%\edi_pac_download.txt > %Logs%\edi_pac.log
%WinSCP% /script=%Scripts%\navicure_download.txt > %Logs%\navicure_down.log