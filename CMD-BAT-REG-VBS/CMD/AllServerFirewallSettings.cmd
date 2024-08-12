REM Firewall Settings All Servers
netsh.exe firewall show notifications 
netsh.exe firewall show logging 
netsh.exe firewall show portopening
netsh.exe firewall show allowedprogram
netsh.exe firewall show currentprofile
REM Reset firewall to default
netsh.exe firewall reset
REM Enable the remote admin feature and limit its scope:
netsh.exe firewall set service type = remoteadmin mode = enable profile = all scope = custom addresses = 10.0.0.0/8
REM Disable the remote administration feature:
pause
netsh.exe firewall set service type = remoteadmin mode = disable profile = all
REM Securing Windows Track :: SANS Institute :: Jason Fossen
REM mode = [ENABLE|DISABLE]
REM profile = [CURRENT|DOMAIN|STANDARD|ALL]
netsh.exe firewall set notifications mode = ENABLE profile = ALL
REM Enable Logging
netsh.exe firewall set logging filelocation = %WinDir%\pfirewall.log maxfilesize = 10053 droppedpackets = ENABLE connections = DISABLE 
REM Dumps most of the configuration settings for the firewall

netsh.exe firewall show config verbose = enable

REM For Windows Vista/2008 and later, start here:

netsh.exe advfirewall show currentprofile
REM Disable Exceptions in Firewall
netsh.exe firewall set opmode mode = enable exceptions = disable
REM   Disables the firewall for all profiles.

netsh.exe advfirewall set allprofiles state off

REM   Enables the firewall for all profiles.

netsh.exe advfirewall set allprofiles state on
REM Add a program to the Exceptions tab and configure its scope.
REM You can also set the scope to ALL or SUBNET.
netsh.exe firewall add allowedprogram program = "%WinDir%\system32\notepad.exe" name = Notepad mode = enable profile = current scope = custom 10.0.0.0/255.0.0.0,192.168.0.0/255.255.0.0 

REM Now delete that excepted program.
pause
netsh.exe firewall delete allowedprogram program = "%WinDir%\system32\notepad.exe" profile = current

REM Create exceptions for ports on the Exceptions tab, along with custom scopes.
REM Protocol can be TCP, UDP or ALL (ALL option adds two exceptions, one for UDP, one for TCP).
REM Custom scope can include subnet mask in either dotted-decimal or CIDR format.

netsh.exe firewall add portopening protocol = all port = 53 name = DNS mode = enable profile = current scope = custom addresses = 10.7.7.7,10.1.1.7,192.168.0.0/255.255.0.0
netsh.exe firewall add portopening protocol = tcp port = 22 name = SSH mode = enable profile = current scope = custom addresses = 52.23.113.3,10.0.0.0/8
REM Now remove the excepted ports.
pause
netsh.exe firewall delete portopening protocol = tcp port = 53 profile = current
netsh.exe firewall delete portopening protocol = udp port = 53 profile = current
netsh.exe firewall delete portopening protocol = tcp port = 22 profile = current
