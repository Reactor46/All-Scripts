:: Installing Printer Ports
CScript //H:CScript //S
setlocal

set PATH="C:\Windows\system32\Printing_Admin_Scripts\en-US\"
:: eReceipts
prndrvr.vbs -a -m "Generic / Text Only" -v 3 -e "Windows NT x86"
prnport.vbs -a -r eReceipts -h 127.0.0.1 -o RAW -n 9100
prnmngr.vbs -a -p "eReceipts" -m "Generic / Text Only" -r eReceipts

::BRANCH 1
prnport.vbs -a -r CANON-B1 -h 192.168.96.30 -o RAW -n 9100
prnport.vbs -a -r SR-PRINT-B1 -h 192.168.96.24 -o RAW -n 9100
prnport.vbs -a -r TLLR-B1 -h 192.168.96.80 -o RAW -n 9100
prnport.vbs -a -r FX-DRAFT-B1 -h 192.168.96.81 -o RAW -n 9100
prnport.vbs -a -r EPSON-B1 -h 192.168.96.35 -o RAW -n 9100


::BRANCH 2
prnport.vbs -a -r CANON-B2 -h 192.168.94.30 -o RAW -n 9100
prnport.vbs -a -r SR-PRINT-B2 -h 192.168.94.52 -o RAW -n 9100
prnport.vbs -a -r TLLR-B2 -h 192.168.94.80 -o RAW -n 9100
prnport.vbs -a -r FX-DRAFT-B2 -h 192.168.94.81 -o RAW -n 9100
prnport.vbs -a -r EPSON-B2 -h 192.168.94.35 -o RAW -n 9100

::BRANCH 3
prnport.vbs -a -r CANON-B3 -h 192.168.95.24 -o RAW -n 9100
prnport.vbs -a -r SR-PRINT-B3 -h 192.168.95.88 -o RAW -n 9100
prnport.vbs -a -r TLLR-B3 -h 192.168.95.80 -o RAW -n 9100
prnport.vbs -a -r FX-DRAFT-B3 -h 192.168.95.81 -o RAW -n 9100
prnport.vbs -a -r EPSON-B3 -h 192.168.95.35 -o RAW -n 9100

::BRANCH 4
prnport.vbs -a -r CANON-B4 -h 192.168.109.30 -o RAW -n 9100
prnport.vbs -a -r SR-PRINT-B4 -h 192.168.109.22 -o RAW -n 9100
prnport.vbs -a -r TLLR-B4 -h 192.168.109.80 -o RAW -n 9100
prnport.vbs -a -r FX-DRAFT-B4 -h 192.168.109.81 -o RAW -n 9100
prnport.vbs -a -r EPSON-B4 -h 192.168.205.35 -o RAW -n 9100

::BRANCH 5
prnport.vbs -a -r CANON-B5 -h 192.168.205.30 -o RAW -n 9100
prnport.vbs -a -r SR-PRINT-B5 -h 192.168.205.8 -o RAW -n 9100
prnport.vbs -a -r TLLR-B5 -h 192.168.205.80 -o RAW -n 9100
prnport.vbs -a -r FX-DRAFT-B5 -h 192.168.205.81 -o RAW -n 9100
prnport.vbs -a -r EPSON-B5 -h 192.168.205.35 -o RAW -n 9100

::BRANCH 6
prnport.vbs -a -r CANON-B6 -h 192.168.97.30 -o RAW -n 9100
prnport.vbs -a -r SR-PRINT-B6 -h 192.168.97.34 -o RAW -n 9100
prnport.vbs -a -r TLLR-B6 -h 192.168.97.80 -o RAW -n 9100
prnport.vbs -a -r FX-DRAFT-B6 -h 192.168.97.81 -o RAW -n 9100
prnport.vbs -a -r EPSON-B6 -h 192.168.97.35 -o RAW -n 9100