@echo off
CScript //H:CScript //S
net use w: \\branch1dc\printers /user:teller password /persistent:no
PATH="C:\Windows\system32\Printing_Admin_Scripts\en-US"

Prndrvr.vbs -a -m "Lexmark Universal" -v 3 -e "Windows NT x86" -i w:\scripted\lex\LMUD0640.INF
prnport.vbs -a -r SR-PRINT -h 192.168.96.24 -o RAW -n 9100
prnmngr.vbs -a -p "SR-PRINT" -m "Lexmark Universal" -r SR-PRINT
prnmngr.vbs -t -p "SR-PRINT"

Prndrvr.vbs -a -m "Lexmark Universal" -v 3 -e "Windows NT x86" -i w:\scripted\lex\LMUD0640INF
prnport.vbs -a -r  CSR1 -h 192.168.96.24 -o RAW -n 9100
prnmngr.vbs -a -p "CSR1" -m "Lexmark Universal" -r CSR1

Prndrvr.vbs -a -m "HP Universal Printing PCL 6" -v 3 -e "Windows NT x86" -i w:\scripted\hp\hpcu115c.inf
prnport.vbs -a -r TLLR -h 192.168.96.80 -o RAW -n 9100
prnmngr.vbs -a -p "TLLR" -m "HP Universal Printing PCL 6" -r TLLR

Prndrvr.vbs -a -m "HP Universal Printing PCL 6" -v 3 -e "Windows NT x86" -i w:\scripted\hp\hpcu115c.inf
prnport.vbs -a -r FX-DRAFT -h 192.168.96.81 -o RAW -n 9100
prnmngr.vbs -a -p "FX-DRAFT" -m "HP Universal Printing PCL 6" -r FX-DRAFT
