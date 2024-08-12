@ECHO OFF
cscript.exe SvcManager.vbs Contosocorp lascsmt01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp lascasmt02 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp lascasmt01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 ContosoDataLayerService start 10
