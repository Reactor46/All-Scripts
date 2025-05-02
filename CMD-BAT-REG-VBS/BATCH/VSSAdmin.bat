Ren " Batch file to resolve Error 0x80042304: The volume Shadow copy provider is not registered in the system"

cd /d %windir%\system32
Net stop vss
Net stop swprv
regsvr32 ole32.dll
regsvr32 oleaut32.dll
regsvr32 vss_ps.dll
vssvc /register
regsvr32 /i swprv.dll
regsvr32 /i eventcls.dll
regsvr32 es.dll
regsvr32 stdprov.dll
regsvr32 vssui.dll
regsvr32 msxml.dll
regsvr32 msxml3.dll
regsvr32 msxml4.dll
Net start vss
Net start swprv
