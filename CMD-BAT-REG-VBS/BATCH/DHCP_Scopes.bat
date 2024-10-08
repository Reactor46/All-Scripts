#  File Name: DHCP_Scopes.BAT 
# +=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+= 
netsh dhcp server show scope > ALL_scopes.txt 
netsh dhcp server show mibinfo > mibinfo.txt 
# +=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+= 
 
 
# File Name: DHCP_details.BAT 
# +=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+= 
@echo off 
for /F %%i in (scopes.txt) do ( 
    echo Processing %%i... 
    netsh dhcp server scope %%i show clients >> show_clients.txt 
    netsh dhcp server scope %%i show clientsv5 >> show_clientsv5.txt 
    netsh dhcp server scope %%i show excluderange >> show_excludeRange.txt 
    netsh dhcp server scope %%i show iprange >> show_iprange.txt 
    netsh dhcp server scope %%i show optionvalue >> show_OptionValue.txt 
    netsh dhcp server scope %%i show reservedip >> show_ReservedIP.txt 
    netsh dhcp server scope %%i show reservedoptionvalue >> reservedoptionvalue.txt 
) 
Echo�. Done ;) 
# +=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=