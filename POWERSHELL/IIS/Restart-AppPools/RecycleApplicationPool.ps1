param( 
[String]$ApplicationPoolComputerName="",
[String]$WebAppPool=""
)

Write-Host "Recycling application pool" $WebAppPool "on" $ApplicationPoolComputerName;
Invoke-Command -ComputerName $ApplicationPoolComputerName -ScriptBlock { Restart-WebAppPool -Name $WebAppPool }
    
$appPoolStatus = Invoke-Command -ComputerName $ApplicationPoolComputerName { param($apn) Import-Module WebAdministration; Get-WebAppPoolState $apn} -Args $WebAppPool

If($appPoolStatus.Value -eq "Stopped")
{
    Write-Host "ERROR - Failed to recycle application pool"
}
else
{
    Write-Host "Recycled application pool" $WebAppPool "on" $ApplicationPoolComputerName;
}