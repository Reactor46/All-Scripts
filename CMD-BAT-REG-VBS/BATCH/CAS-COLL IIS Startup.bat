@ECHO OFF
:: START CAS IIS AND COLL IIS SERVICES

::CAS
cscript.exe SvcManager.vbs Contosocorp LASCAS03 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCAS04 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCAS05 W3SVC start 300
::COLLS
cscript.exe SvcManager.vbs Contosocorp LASCOLL01 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 W3SVC start 300
cscript.exe SvcManager.vbs CREDITONEAPP LASAUTH01.CREDITONEAPP.BIZ W3SVC start 300
cscript.exe SvcManager.vbs CREDITONEAPP LASAUTH02.CREDITONEAPP.BIZ W3SVC start 300