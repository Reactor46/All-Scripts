function Get-SPRComputerNameByAlias
{
  param
  (
    [String][Parameter(Mandatory)]
    $Alias
  )
	
  $registryKey = Get-ItemProperty -Path HKLM:\software\Microsoft\MSSQLServer\Client\ConnectTo
	
  $array = ($registryKey.$Alias) -split ','
  $output = $array[1]
	
  Write-Output -InputObject $output
}