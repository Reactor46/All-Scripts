:: Installing Printer Ports
CScript //H:CScript //S
setlocal

set PrintScripts="C:\Windows\system32\Printing_Admin_Scripts\en-US"
set DEPLOYSCRIPT=\\branch1dc\deploy$
REM set DEPLOYSCRIPT=\\branch2dc\deploy$
REM set DEPLOYSCRIPT=\\branch3dc\deploy$
REM set DEPLOYSCRIPT=\\branch4dc\deploy$
REM set DEPLOYSCRIPT=\\branch5dc\deploy$
REM set DEPLOYSCRIPT=\\branch6dc\deploy$

:PrinterDriver
%PrintScripts%\prndrvr.vbs -a -m "Lexmark Universal" -v 3 -e "Windows NT x86" -i %DEPLOYSCRIPT%\Apps\print\lex\UDO\LMUD0640.INF
%PrintScripts%\prndrvr.vbs -a -m "Lexmark MS310 Series XL" -v 3 -e "Windows NT x86" -i %DEPLOYSCRIPT%\Apps\print\lex\MS310\LMADSP40.inf
%PrintScripts%\prndrvr.vbs -a -m "HP Universal Printing PCL 5" -v 3 -e "Windows NT x86" -i %DEPLOYSCRIPT%\Apps\print\hp\PCL5\pcl5-x32-5.6.5.15717\hpcu150b.inf
%PrintScripts%\prndrvr.vbs -a -m "HP LaserJet 2300 Series PCL 6" -v 3 -e "Windows NT x86" -i %DEPLOYSCRIPT%\Apps\print\hp\2300\hp2300c.inf
%PrintScripts%\prndrvr.vbs -a -m "Canon iR3225 PCL6" -v 3 -e "Windows NT x86" -i %DEPLOYSCRIPT%\Apps\print\can\iR3225\pcl6\P62KUSAL.INF
%PrintScripts%\prndrvr.vbs -a -m "Canon iR1730/1740/1750 PCL6" -v 3 -e "Windows NT x86" -i %DEPLOYSCRIPT%\Apps\print\can\iF1730\CNP60U.INF
%PrintScripts%\prndrvr.vbs -a -m "Canon iR-ADV 4045/4051 PCL6" -v 3 -e "Windows NT x86" -i %DEPLOYSCRIPT%\Apps\print\can\4045\CNP60U.INF


:eReceipts
%PrintScripts%\prndrvr.vbs -a -m "Generic / Text Only" -v 3 -e "Windows NT x86"
%PrintScripts%\prnport.vbs -a -r eReceipts -h 127.0.0.1 -o RAW -n 9100
%PrintScripts%\prnmngr.vbs -a -p "eReceipts" -m "Generic / Text Only" -r eReceipts

:BRANCH1
%PrintScripts%\prnport.vbs -a -r CANON-B1 -h 192.168.96.30 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r SR-PRINT-B1 -h 192.168.96.24 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r TLLR-B1 -h 192.168.96.80 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r FX-DRAFT-B1 -h 192.168.96.81 -o RAW -n 9100

:BRANCH2
%PrintScripts%\prnport.vbs -a -r CANON-B2 -h 192.168.94.30 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r SR-PRINT-B2 -h 192.168.94.52 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r TLLR-B2 -h 192.168.94.80 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r FX-DRAFT-B2 -h 192.168.94.81 -o RAW -n 9100

:BRANCH3
%PrintScripts%\prnport.vbs -a -r CANON-B3 -h 192.168.95.24 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r SR-PRINT-B3 -h 192.168.95.88 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r TLLR-B3 -h 192.168.95.80 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r FX-DRAFT-B3 -h 192.168.95.81 -o RAW -n 9100

:BRANCH4
%PrintScripts%\prnport.vbs -a -r CANON-B4 -h 192.168.109.30 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r SR-PRINT-B4 -h 192.168.109.22 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r TLLR-B4 -h 192.168.109.80 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r FX-DRAFT-B4 -h 192.168.109.81 -o RAW -n 9100

:BRANCH5
%PrintScripts%\prnport.vbs -a -r CANON-B5 -h 192.168.205.30 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r SR-PRINT-B5 -h 192.168.205.8 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r TLLR-B5 -h 192.168.205.80 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r FX-DRAFT-B5 -h 192.168.205.81 -o RAW -n 9100

:BRANCH6
%PrintScripts%\prnport.vbs -a -r CANON-B6 -h 192.168.97.30 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r SR-PRINT-B6 -h 192.168.97.34 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r TLLR-B6 -h 192.168.97.80 -o RAW -n 9100
%PrintScripts%\prnport.vbs -a -r FX-DRAFT-B6 -h 192.168.97.81 -o RAW -n 9100


