setlocal
If .%1 EQU . (set hostname=%computername%)&(set name=)&(Goto :Start)
Set hostname=%1
set name=-r:%1
 
:Start
echo . >%hostname%.txt
echo . >%hostname%.Error.txt
 
:Ping
echo ping %hostname% >>%hostname%.txt
echo ping %hostname% >>%hostname%.Error.txt
echo +++++++++++++++++++++++++++++++++++++++++++++  >>%hostname%.txt
echo +++++++++++++++++++++++++++++++++++++++++++++  >>%hostname%.Error.txt
ping %hostname% >>%hostname%.txt 2>>%hostname%.Error.txt
echo ================================================================= >>%hostname%.txt
echo ================================================================= >>%hostname%.Error.txt
 
:VMMAgent
:echo winrm invoke GetVersion wmi/root/scvmm/AgentManagement %name% @{}  >>%hostname%.txt
:echo winrm invoke GetVersion wmi/root/scvmm/AgentManagement %name% @{}  >>%hostname%.Error.txt
:echo +++++++++++++++++++++++++++++++++++++++++++++  >>%hostname%.txt
:echo +++++++++++++++++++++++++++++++++++++++++++++  >>%hostname%.Error.txt
:call winrm invoke GetVersion wmi/root/scvmm/AgentManagement %name% @{} >>%hostname%.txt 2>>%hostname%.Error.txt
:echo ================================================================= >>%hostname%.txt
:echo ================================================================= >>%hostname%.Error.txt
 
:Cluster
:: SCVMM cluster information pulled
:Call :WinRM wmi/root/mscluster/MSCluster_ClusterToNode
:Call :WinRM wmi/root/mscluster/MSCluster_Cluster
:Call :WinRM wmi/root/mscluster/MSCluster_ResourceGroup
:Call :WinRM wmi/root/mscluster/MSCluster_ResourceGroupToResource
:Call :WinRM wmi/root/mscluster/MSCluster_Service
:: During host refresh, for example
:Call :WinRM wmi/root/mscluster/MSCluster_ResourceToDisk
 
:System
:: Initial system scan during P2V, for example
Call :WinRM wmi/root/cimv2/Win32_ComputerSystem
Call :WinRM wmi/root/cimv2/win32_OperatingSystem
:: Collects avail memory during host refresh, for example
Call :WinRM wmi/root/cimv2/Win32_PerfRawData_PerfOS_Memory
 
:End
exit /b
:: ----------
 
:WinRM
echo winrm enum %1 %name% >>%hostname%.txt
echo winrm enum %1 %name%>>%hostname%.Error.txt
echo +++++++++++++++++++++++++++++++++++++++++++++  >>%hostname%.txt
echo +++++++++++++++++++++++++++++++++++++++++++++  >>%hostname%.Error.txt
Call winrm enum %1 %name% >>%hostname%.txt 2>>%hostname%.Error.txt
echo ================================================================= >>%hostname%.txt
echo ================================================================= >>%hostname%.Error.txt
echo.  >>%hostname%.txt
echo.  >>%hostname%.Error.txt
Goto :eof

Identify the Problem

Depending on how WinRM is being used, there should be either output errors or event log entries to review. WinRM errors are typically specific and should point in the direction of the problem.

Console
If using a command prompt or PowerShell the error will output to the console.

C:\>winrm id -r:cssvmm-n1
WSManFault
    Message = The WinRM client cannot complete the operation within the time specified. Check if the
 machine name is valid and is reachable over the network and firewall exception for Windows Remote M
anagement service is enabled.
 
Error number:  -2144108250 0x80338126
The WinRM client cannot complete the operation within the time specified. Check if the machine name
is valid and is reachable over the network and firewall exception for Windows Remote Management serv
ice is enabled.