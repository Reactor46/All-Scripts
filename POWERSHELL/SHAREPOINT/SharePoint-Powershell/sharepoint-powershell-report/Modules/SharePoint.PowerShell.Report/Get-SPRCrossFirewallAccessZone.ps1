function Get-SPRCrossFirewallAccessZone
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPWebApplication
  )
	
  foreach ($WebApp in $SPWebApplication)
  {
    $crossFirewallAccessZone = $WebApp.ExternalURLZone
    if($crossFirewallAccessZone -eq $null){$crossFirewallAccessZone = 'none'}

    $properties = [ordered]@{
      'WebApplication'        = $WebApp.DisplayName
     
      'CrossFirewallAccessZone' = $crossFirewallAccessZone
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}