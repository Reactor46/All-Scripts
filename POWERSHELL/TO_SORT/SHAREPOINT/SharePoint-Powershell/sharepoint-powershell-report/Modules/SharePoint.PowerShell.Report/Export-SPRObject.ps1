function Export-SPRObject
{
  param
  (
    [Parameter(Mandatory = $true)]
    [string]$Name,
    [Parameter(Mandatory = $true)]
    [object]$Object,
    [Parameter(Mandatory = $true)]
    [string]$Path
  )
	

  if (!(Test-Path -PathType Container -Path $Path))
  {
    mkdir -Path $Path
  }
  $fileName = '{0}\{1}.xml' -f $Path, $Name
	
  Write-Host -Object ('Exporting {0} object(s) to {1}.' -f $Name, $fileName )
  Start-Job -ScriptBlock {
    $args[0] | Export-Clixml -Path $args[1]
  } -ArgumentList $Object, $fileName -Name $Name | Out-Null
}
