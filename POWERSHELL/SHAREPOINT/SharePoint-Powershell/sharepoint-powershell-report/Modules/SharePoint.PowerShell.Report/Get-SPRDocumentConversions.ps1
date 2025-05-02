function Get-SPRDocumentConversions
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPWebApplication
  )

  foreach ($WebApp in $SPWebApplication)
  {
  $DConversionEnabled=$WebApp.DocumentConversionsEnabled

  if ($DConversionEnabled) {

    foreach ($converter in $WebApp.DocumentConverters)
    {
    
       $properties = [ordered]@{
      'WebApplication' = $WebApp.DisplayName
      'Converter'    = $converter.DisplayName
      'Type'         = [string]$converter.ConvertFrom + '>' + [string]$converter.ConvertTo
      'Timeout'      = $converter.Timeout
      'MaxRetries'   = $converter.Retries
      'MaxFileSize'  = $converter.MaxFileSize
    }
     $output = New-Object -TypeName PSObject -Property $properties	
     Write-Output -InputObject $output
    
    }

   }
  

  }
}