@ECHO OFF
cscript.exe SvcManager.vbs Contosocorp lascsmt01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp lascasmt02 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp lascasmt01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 ContosoDataLayerService stop 10
