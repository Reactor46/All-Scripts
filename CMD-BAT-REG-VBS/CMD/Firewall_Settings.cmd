
::Adding Firewall Exceptions
netsh advfirewall firewall set rule group="windows management instrumentation (WMI)" new enable=Yes
netsh advfirewall firewall set rule group="Windows Remote Management" new enable=yes
netsh advfirewall firewall set rule group="remote desktop" new enable=Yes 
netsh advfirewall firewall set rule group="remote administration" new enable=yes
netsh advfirewall firewall add rule name="ICMP Allow incoming V4 echo request" protocol=icmpv4:8,any dir=in action=allow
netsh advfirewall firewall add rule name="COWWWThreads" dir=in action=allow program="%PROGRAMFILES%\Open Solutions\eReceipts\COWWWReceiptThreads.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="eReceipts" dir=in action=allow program="%PROGRAMFILES%\Open Solutions\eReceipts\eReceipts.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="VerantID" dir=in action=allow program="%PROGRAMFILES%\VerantID\PIVS\ScanID.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="Nexus ECU Remote" dir=in action=allow program="%PROGRAMFILES%\Nexus\Involve\Device Services\EcuRemote.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="Nexus WOSA/XFS Service" dir=in action=allow program="%PROGRAMFILES%\Nexus\Involve\Device Services\WrmServ.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="Nexus Trace Facility" dir=in action=allow program="%PROGRAMFILES%\Nexus\Involve\Device Services\ntfsvc.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="DameWare Mini Remote Control" dir=in action=allow program="%WINDIR%\dwrcs\DWRCS.EXE" enable=yes profile=domain
netsh advfirewall firewall add rule name="DameWare NT Utilities" dir=in action=allow program="%WINDIR%\dwrcs\DNTUS26.EXE" enable=yes profile=domain
netsh advfirewall firewall add rule name="Dell KACE Agent" dir=in action=allow program="%PROGRAMFILES%\Dell\KACE\AMPAgent.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="ESET Service" dir=in action=allow program="%PROGRAMFILES%\ESET\ESET Endpoint Antivirus\ekrn.exe" enable=yes profile=domain
netsh advfirewall set currentprofile logging filename %systemroot%\system32\LogFiles\Firewall\pfirewall.log
netsh advfirewall set currentprofile logging maxfilesize 4096
netsh advfirewall set currentprofile logging droppedconnections enable
netsh advfirewall set currentprofile logging allowedconnections enable
