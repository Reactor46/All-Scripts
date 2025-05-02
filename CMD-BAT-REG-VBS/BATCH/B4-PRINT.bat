@echo off
CScript //H:CScript //S
net use w: \\branch1dc\printers /user:teller password /persistent:no
PATH="C:\Windows\system32\Printing_Admin_Scripts\en-US"

Prndrvr.vbs -a -m "Canon Generic PCL6 Driver" -v 3 -e "Windows NT x86" -i w:\scripted\canon\CNP60U.INF
prnport.vbs -a -r IP_192.168.109.30 -h 192.168.109.30 -o RAW -n 9100
prnmngr.vbs -a -p "Canon" -m "Canon Generic PCL6 Driver" -r IP_192.168.109.30
prnmngr.vbs -t -p "Canon"

Prndrvr.vbs -a -m "Canon Generic PCL6 Driver" -v 3 -e "Windows NT x86" -i w:\scripted\canon\CNP60U.INF
prnport.vbs -a -r IP_192.168.109.30 -h 192.168.109.30 -o RAW -n 9100
prnmngr.vbs -a -p "CSR1" -m "Canon Generic PCL6 Driver" -r IP_192.168.109.30

Prndrvr.vbs -a -m "HP Universal Printing PCL 6" -v 3 -e "Windows NT x86" -i w:\scripted\hp\hpcu115c.inf
prnport.vbs -a -r IP_192.168.109.80 -h 192.168.109.80 -o RAW -n 9100
prnmngr.vbs -a -p "TLLR" -m "HP Universal Printing PCL 6" -r IP_192.168.109.80

Prndrvr.vbs -a -m "HP Universal Printing PCL 6" -v 3 -e "Windows NT x86" -i w:\scripted\hp\hpcu115c.inf
prnport.vbs -a -r IP_192.168.109.81 -h 192.168.109.81 -o RAW -n 9100
prnmngr.vbs -a -p "FX-DRAFT" -m "HP Universal Printing PCL 6" -r IP_192.168.109.81