#################################################################################  
#  
# The sample scripts are not supported under any Microsoft standard support  
# program or service. The sample scripts are provided AS IS without warranty 
# of any kind. Microsoft further disclaims all implied warranties including, 
# without limitation, any implied warranties of merchantability or of fitness for  
# a particular purpose. The entire risk arising out of the use or performance of 
# the sample scripts and documentation remains with you. In no event shall
# Microsoft, its authors, or anyone else involved in the creation, production, or
# delivery of the scripts be liable for any damages whatsoever (including,
# without limitation, damages for loss of business profits, business interruption,  
# loss of business information, or other pecuniary loss) arising out of the use
# of or inability to use the sample scripts or documentation, even if Microsoft  
# has been advised of the possibility of such damages  
#  
################################################################################# 


#################################################################################  
# Script name: UpdateSQLAlwaysOnVIP-Azure.ps1
# Script usage: Can be run or can be scheduled on any cluster node. 
#	This cluster should have the Availability Group in IaaS.
# Pre-requisite 1: In a Powershell window on each cluster node, run "Enable-PSRemoting" (without quotes). This is a one-time activity.
# Pre-requisite 2: Update values of variables $ServiceName (cloud service name), $AvailabilityGroupName, $Nodes. This is a one-time activity.
# Windows Task Scheduler: Below to schedule script in scheduler. Change user account, time and path as appropriate.: 
#              schtasks /CREATE /TN "UpdateSQLAlwaysOnVIP-Azure" /RU "corp\vijayrod" /ST 03:49:00  /SC DAILY /RL HIGHEST /TR "powershell -f 'C:\UpdateSQLAlwaysOnVIP-Azure.ps1' -ExecutionPolicy Unrestricted" /F 
################################################################################# 

$ServiceName="SQLAA"                        
$AvailabilityGroupName="AG1"
$Nodes="ContosoSQL1","ContosoSQL2"

$serviceIP = (Resolve-DnsName  "$ServiceName.cloudapp.net").IPAddress
Write-Host "Updating SQL Always On to use Service IP: $serviceIP" -ForegroundColor Green

$probePort = "59999"
Get-ClusterResource | Where { $_.OwnerGroup -eq "$AvailabilityGroupName" -and $_.ResourceType -eq "IP Address" } | Set-ClusterParameter -Multiple @{"Address"="$serviceIP";"ProbePort"="$probePort";SubnetMask="255.255.255.255";"Network"=(Get-ClusterNetwork)[0].Name;"OverrideAddressMatch"=1;"EnableDhcp"=0} -ErrorAction Continue;
foreach($node in $nodes)
    {
        Write-Host "Restarting MSSQLSERVER on $node" -ForegroundColor Green
        Invoke-Command -ComputerName $node -ScriptBlock {
            Restart-Service -Name "MSSQLSERVER" -Force
        }
    }