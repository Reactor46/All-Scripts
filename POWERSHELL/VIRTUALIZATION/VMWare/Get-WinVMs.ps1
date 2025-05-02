$vCenterServers = @(
    'AUCKLAND-VC01', 
    'BRISBANE-VC01', 
    'CAPETOWN-VC02', 
    'CHENNAI-VC02', 
    'LIVINGSTON-VC02', 
    'NANJING-VC01'
)


#Loop through the vCenter servers above and get the data required
foreach($vCenter in $vCenterServers)
{
    #Connect to the vCenter servers above
    Connect-VIServer -Server $vCenter
    
    #Count the VMs and write the total on the screen
    #write-host (Get-VMguest -Server $vCenter -VM * | Where-Object {$_.OSFullName -like "*Windows Server*" }).Count
    
    #Get a list of the VMs that have Windows Server in them and output to a csv file
    Get-VMguest -Server $vCenter -VM * | Where-Object {$_.OSFullName -like "*Windows Server*" } | Export-Csv -Path c:\Scripts\VMsAndHostsInfo-$vCenter.csv -NoTypeInformation -UseCulture
}

# Disconnect from all vCenter servers after script has run
if ($global:DefaultVIServers.Count -gt 0) {
    Disconnect-VIServer * -Confirm:$false
}