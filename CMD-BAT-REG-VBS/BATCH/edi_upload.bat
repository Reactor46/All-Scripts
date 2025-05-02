@ECHO ON
SETLOCAL
SET WinSCP=E:\Scripts\WinSCP\WinSCP.COM
SET Scripts=E:\Scripts\WinSCP\UPDOWN
SET Logs=E:\Scripts\WinSCP\LOGS

%WinSCP% /script=%Scripts%\edi_omg_upload.txt > %Logs%\edi_up_omg.log
%WinSCP% /script=%Scripts%\edi_pac_upload.txt > %Logs%\edi_up_pac.log
%WinSCP% /script=%Scripts%\navicure_upload.txt > %Logs%\navicure_up.log
%WinSCP% /script=%Scripts%\edi_ortho_upload.txt > %Logs%\edi_up_ort.log
%WinSCP% /script=%Scripts%\edi_onc_upload.txt > %Logs%\edi_up_onc.log
%WinSCP% /script=%Scripts%\edi_pulm_upload.txt > %Logs%\edi_up_pulm.log
%WinSCP% /script=%Scripts%\edi_RAD_upload.txt > %Logs%\edi_up_RAD.log