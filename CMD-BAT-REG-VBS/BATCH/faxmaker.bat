IF EXIST "c:\fmlog.txt" GOTO eof

msiexec /i \\uson.local\netlogon\Faxmaker\faxclient.msi /quiet SILENTPRINTERDRIVER=1 USEOUTLOOKFORM=0 FAXSERVER=usonvsvrfax01 DETAILSTYPE=3 /log c:\fmlog.txt

:eof

cscript //nologo \\uson.local\netlogon\GFI.vbs

END && EXIT




