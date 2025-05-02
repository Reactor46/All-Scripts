@Echo Off

set INPUT=
set /P INPUT=Enter Username: %=%
runas /user:uson.local\%INPUT% "CMD /k \"\\usonvsvrfs01\MSO IT\Scripts\BAT\printers\print spooler clear new.bat\""