IF EXIST H:\MappedPrinters.txt goto end else goto MapPrinters 
:MapPrinters
H:\ > MappedPrinters.txt
cscript //nologo \\usonvsvrdc\sysvol\USON.local\scripts\MapPrinters\RemovenMapNetPrinters.vbs
:end



