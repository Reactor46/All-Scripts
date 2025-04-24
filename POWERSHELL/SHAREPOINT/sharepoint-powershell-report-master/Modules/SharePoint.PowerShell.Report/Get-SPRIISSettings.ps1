function Get-SPRIISSettings
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPWebApplication,
    [Parameter(Mandatory = $true)]
    [object[]]$SPAlternateURL,
    [Parameter(Mandatory = $true)]
    [object[]]$SPAuthenticationProvider
  )
	
  foreach ($webApplication in $SPWebApplication)
  {
    $alternateUrls = $SPAlternateURL | Where-Object -FilterScript {
      $_.WebApplicationUrl -eq $webApplication.Url
    } 
    
    foreach ($alternateUrl in $alternateUrls)
    {
      $authenticationProvider = $SPAuthenticationProvider | Where-Object -FilterScript {
        $_. WebApplicationUrl -eq $webApplication.Url -and $_.Zone -eq $alternateUrl.Zone
      } 
      $ssl = $alternateUrl.PublicUrl -like 'https*'
      if ($authenticationProvider.DisableKerberos -eq $true -and $authenticationProvider.UseWindowsIntegratedAuthentication -eq $true)
      {
        $authentication = 'NTLM'
      }
      elseif ($authenticationProvider.DisableKerberos -eq $false -and $authenticationProvider.UseWindowsIntegratedAuthentication -eq $true)
      {
        $authentication = 'Kerberos'
      }
      
      $properties = [ordered]@{
        'DisplayName'           = $webApplication.DisplayName
        'Url'                   = $alternateUrl.IncomingUrl
        'Zone'                  = $alternateUrl.Zone
        'Authentication'        = $authentication
        'ApplicationPoolName'   = $webApplication.ApplicationPool.Name
        'ApplicationPoolIdentity' = $webApplication.ApplicationPool.Username
        'SSL'                   = $ssl
        'ClaimsAuth'            = $webApplication.UseClaimsAuthentication
        'CEIP'                  = $webApplication.BrowserCEIPEnabled
      }
      
      $output = New-Object -TypeName PSObject -Property $properties
      Write-Output -InputObject $output
    }
  }
}
