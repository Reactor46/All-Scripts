netsh advfirewall firewall add rule name="HP WJA TFTP port 69" dir=in action=allow protocol=UDP localport=69
netsh advfirewall firewall add rule name="HP WJA Remote Control Panel of EWS" dir=out action=allow protocol=TCP localport=80
netsh advfirewall firewall add rule name="HP WJA SNMP" dir=out action=allow protocol=UDP localport=161
netsh advfirewall firewall add rule name="HP WJA Discovery: SLP Listen" dir=in action=allow protocol=UDP localport=427
netsh advfirewall firewall add rule name="HP WJA https" dir=out action=allow protocol=TCP localport=443
netsh advfirewall firewall add rule name="HP WJA fax/scan configuration" dir=out action=allow protocol=TCP localport=843
netsh advfirewall firewall add rule name="HP WJA Remote SQL server" dir=in action=allow protocol=UDP localport=1433
netsh advfirewall firewall add rule name="HP WJA Remote SQL server" dir=out action=allow protocol=UDP localport=1433
netsh advfirewall firewall add rule name="HP WJA Discovery: other HP WJA servers" dir=in action=allow protocol=UDP localport=2493
netsh advfirewall firewall add rule name="HP WJA Discovery: other HP WJA servers" dir=out action=allow protocol=UDP localport=2493
netsh advfirewall firewall add rule name="HP WJA Discovery: WS Discovery" dir=out action=allow protocol=UDP localport=3702
netsh advfirewall firewall add rule name="HP WJA Print Request status" dir=out action=allow protocol=TCP localport=3910
netsh advfirewall firewall add rule name="HP WJA Printer Status" dir=out action=allow protocol=TCP localport=3911
netsh advfirewall firewall add rule name="HP WJA client communication" dir=in action=allow protocol=TCP localport=4088
netsh advfirewall firewall add rule name="HP WJA client" dir=in action=allow protocol=TCP localport=4089
netsh advfirewall firewall add rule name="HP WJA Web Services" dir=out action=allow protocol=TCP localport=7627
netsh advfirewall firewall add rule name="HP WJA Discovery Listen" dir=out action=allow protocol=UDP localport=8000
netsh advfirewall firewall add rule name="HP WJA client UI and WJA Help (http)" dir=in action=allow protocol=TCP localport=8000
netsh advfirewall firewall add rule name="HP WJA Pro Adapter" dir=out action=allow protocol=TCP localport=8080
netsh advfirewall firewall add rule name="HP WJA Device communication: WS eventing" dir=in action=allow protocol=TCP localport=8050
netsh advfirewall firewall add rule name="HP WJA OXPm Web Services (http)" dir=in action=allow protocol=TCP localport=8140
netsh advfirewall firewall add rule name="HP WJA OXPm Web Services (https)" dir=in action=allow protocol=TCP localport=8143
netsh advfirewall firewall add rule name="HP WJA Client UI and WJA help (https)" dir=in action=allow protocol=TCP localport=8443
netsh advfirewall firewall add rule name="HP WJA file transfer to printers" dir=out action=allow protocol=TCP localport=9100
netsh advfirewall firewall add rule name="HP WJA SNMP Trap Listener" dir=in action=allow protocol=UDP localport=27892
netsh advfirewall firewall add rule name="HP WJA Communication with WS Proxy Agent" dir=in action=allow protocol=UDP localport=27893
netsh advfirewall firewall add rule name="HP WJA Communication with remote SQL server" dir=out action=allow protocol=TCP localport=59113
netsh advfirewall firewall add rule name="HP WJA Instant On Listener for Security Manager" dir=in action=allow protocol=TCP localport 3329
netsh advfirewall firewall set rule group="windows management instrumentation (wmi)" new enable=yes