function ConvertTo-SPRLoginName
{
  param
  (
    [Parameter(Mandatory = $true)]
    [string]$Login
  )
	
  $array = $Login.Split('|')
  [array]::Reverse($array)
	
  $output = $array[0]
	
  Write-Output -InputObject $output
}