Clear-Host
Get-Job | Stop-Job
Get-Job | Remove-Job

if ($PSScriptRoot -eq $null)
{
  $Location = $PSCommandPath
}
else
{
  $Location = $PSScriptRoot
}

##region GeneralSettings
$moduleName = 'SharePoint.PowerShell.Report'
$exportPath = '{0}\Export' -f $Location
$reportFilePath = '{0}\Report\Report.html' -f $Location
#endregion GeneralSettings

if ($env:PSModulePath -notlike ('*{0}*' -f $PSScriptRoot))
{
  #  $env:PSModulePath += ';{0}\Modules' -f $PSScriptRoot
  $env:PSModulePath += ';{0}' -f $PSScriptRoot
}

#region reload module

#$moduleLocation = '{0}\Modules\{1}\{1}.psd1' -f $Location, $moduleName
if (Get-Module -Name $moduleName)
{
  Remove-Module -Name $moduleName
  #Import-Module -Name $moduleName
}
#endregion reload module

#Export-SPRObjects -Path $exportPath
Get-SPReport -ExportPath $exportPath -FromExportedFiles -ReportFilePath $reportFilePath
#Invoke-Item $reportFilePath























































































































































































































