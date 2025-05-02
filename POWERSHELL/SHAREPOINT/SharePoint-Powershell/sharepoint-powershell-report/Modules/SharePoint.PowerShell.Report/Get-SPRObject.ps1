function Get-SPRObject
{
  param
  (
    [Parameter(ParameterSetName = 'Offline',
    Mandatory = $true)]
    [String]$Path,
    [Parameter(ParameterSetName = 'Online')]
    [switch]$Online
  )
	
  $output = @{ }
	
  if (!($Online))
  {
    # we get our objects from files
    $files = Get-ChildItem -Path $Path
		
    foreach ($file in $files)
    {
      $key = $file.BaseName
      $value = Import-Clixml -LiteralPath $file.FullName
			
      $output.Add($key, $value)
    }
  }
	
  Write-Output -InputObject $output
}