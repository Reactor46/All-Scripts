# Функция для получения информации о vCenter.
# Параметры: сервер vCenter.
#
# Любимов Роман, 2015-2017

function Get-vCenterInfo
{
	Param(
		[Parameter(Mandatory = $True)]
		[VMware.VimAutomation.Types.VIServer] $vc
	)

	$report = New-Object psobject
	$report | Add-Member -type noteproperty -name vCenter -Value $vc.Name
	$report | Add-Member -type noteproperty -name Version -Value $vc.Version
	$report | Add-Member -type noteproperty -name Build -Value	$vc.Build
	
	return $report
}