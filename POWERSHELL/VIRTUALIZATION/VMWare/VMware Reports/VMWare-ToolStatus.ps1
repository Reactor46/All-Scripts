Connect-VIServer -Server corp-vmware.vegas.com
$VMs = Get-VM | Select PowerState, Name,@{Name="ToolsVersion";Expression={$_.ExtensionData.Guest.ToolsVersion}},@{Name="ToolsStatus";Expression={$_.ExtensionData.Guest.ToolsVersionStatus}}

$VMs | Select Name, ToolsVersion, ToolsStatus, PowerState | Sort-Object -Property ToolsStatus | Export-Csv -Path .\Corp-VMWare-Tools_All.csv -NoTypeInformation

