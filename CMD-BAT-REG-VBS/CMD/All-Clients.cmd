@ECHO ON
ECHO Updating LAS Clients
psexec @LAS-Clients "C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\12.1.7004.6500.105\Bin\Smc.exe" -updateconfig
psexec @LAS-Clients.txt -c -s -f sylink.cmd >> SEPM-Clients.log 2>&1
ECHO Updating PHX Clients
psexec @PHX-Clients.txt -c -s -f phx-sylink.cmd >> SEPM-Clients.log 2>&1