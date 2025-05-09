Clear-Host

Set-PowerCLIConfiguration -InvalidCertificateAction:Ignore -DefaultVIServerMode:Single -ParticipateInCEIP:$false -Confirm:$false -Scope:Session | Out-Null

."$PSScriptRoot\Settings.ps1"
."$PSScriptRoot\Get-vCenterInfo.ps1"
."$PSScriptRoot\Get-HostInfo.ps1"
."$PSScriptRoot\Get-HostUplinkInfo.ps1"
."$PSScriptRoot\Get-StorageInfo.ps1"
."$PSScriptRoot\Get-SnapshotInfo.ps1"
."$PSScriptRoot\Get-VMInfo.ps1"

$DataStorePath = (Get-Item $PSScriptRoot).Parent.FullName + "\DataStore"

$CurrentDateTime = (Get-Date -uformat "%Y%m%d%H%M%S").ToString()

$vCenterInfoArr = @()
$HostInfoArr = @()
$HostUplinkInfoArr = @()
$StorageInfoArr = @()
$SnapshotInfoArr = @()
$VMInfoArr = @()

$vCenterServers | % {
	Write-Host "$(Get-Date) Connect-VIServer $_"
	$vc = Connect-VIServer -Server $_ -User "administrator@vsphere.local" -Password "J@bb3rJ4w"
	
	Write-Host "$(Get-Date) Get-vCenterInfo"
	$vCenterInfoArr += Get-vCenterInfo $vc
	
	Write-Host "$(Get-Date) Get-Datacenter"
	Get-Datacenter | % {
		$vcName = $vc.Name
		$dcName = $_.Name
		
		Write-Host "$(Get-Date) Get-VMHost"
		$_ | Get-VMHost | sort Parent, Name | % {
			Write-Host "$(Get-Date) Get-HostInfo $($_.Name)"
			$HostInfoArr += Get-HostInfo $_ $vcName $dcName
			
			Write-Host "$(Get-Date) Get-HostUplinkInfo $($_.Name)"
			$HostUplinkInfoArr += Get-HostUplinkInfo $_ $vcName $dcName
		}
		
		Write-Host "$(Get-Date) Get-StorageInfo"
		$StorageInfoArr += Get-StorageInfo $_ $vcName
		
		Write-Host "$(Get-Date) Get-VM"
		$_ | Get-VM | % {
			Write-Host "$(Get-Date) Get-SnapshotInfo $($_.Name)"
			$SnapshotInfoArr += Get-SnapshotInfo $_ $vcName $dcName
			
			Write-Host "$(Get-Date) Get-VMInfo $($_.Name)"
			$VMInfoArr += Get-VMInfo $_ $vcName $dcName
		}
	}
	
	Write-Host "$(Get-Date) Disconnect-VIServer $_"
	Disconnect-VIServer -Server $vc -Force -Confirm:$false
}

$delimiter = "`t"
$encoding = "Unicode"

$vCenterInfoArr | Export-Csv -Path ($DataStorePath + "\vCenterInfo\" + $CurrentDateTime + ".csv") -Delimiter $delimiter -Encoding $encoding -NoTypeInformation
$HostInfoArr | Export-Csv -Path ($DataStorePath + "\HostInfo\" + $CurrentDateTime + ".csv") -Delimiter $delimiter -Encoding $encoding -NoTypeInformation
$HostUplinkInfoArr | Export-Csv -Path ($DataStorePath + "\HostUplinkInfo\" + $CurrentDateTime + ".csv") -Delimiter $delimiter -Encoding $encoding -NoTypeInformation -ErrorAction SilentlyContinue
$StorageInfoArr | Export-Csv -Path ($DataStorePath + "\StorageInfo\" + $CurrentDateTime + ".csv") -Delimiter $delimiter -Encoding $encoding -NoTypeInformation
$SnapshotInfoArr | sort vCenter, Datacenter, VM, Created | Export-Csv -Path ($DataStorePath + "\SnapshotInfo\" + $CurrentDateTime + ".csv") -Delimiter $delimiter -Encoding $encoding -NoTypeInformation
$VMInfoArr | sort vCenter, Datacenter, "VM Path", "VM Name" | Export-Csv -Path ($DataStorePath + "\VMInfo\" + $CurrentDateTime + ".csv") -Delimiter $delimiter -Encoding $encoding -NoTypeInformation