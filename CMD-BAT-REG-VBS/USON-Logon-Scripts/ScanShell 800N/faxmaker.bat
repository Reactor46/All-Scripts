IF EXIST "C:\Program Files (x86)\GFI\FaxMaker Client\SendFax.exe" GOTO eof


msiexec /i \\uson.local\netlogon\Faxmaker\faxclient.msi /quiet SILENTPRINTERDRIVER=1 USEOUTLOOKFORM=0 FAXSERVER=usonvsvrfax01 DETAILSTYPE=3 /log c:\fmlog.txt

:eof

END && EXIT


