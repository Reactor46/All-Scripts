@ECHO OFF
:: SHUTS DOWN CAS IIS AND COLL IIS SERVICES

::CAS
cscript.exe SvcManager.vbs Contosocorp LASCAS03 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCAS04 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCAS05 W3SVC stop 300
::COLLS
cscript.exe SvcManager.vbs Contosocorp LASCOLL01 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 W3SVC stop 300
cscript.exe SvcManager.vbs creditoneapp LASAUTH01.CREDITONEAPP.BIZ W3SVC stop 300
cscript.exe SvcManager.vbs creditoneapp LASAUTH02.CREDITONEAPP.BIZ W3SVC stop 300