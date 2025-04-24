#requires -Version 3.0
function Get-SPRWebApplicationUserPolicy
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPWebApplication,
    [Parameter(Mandatory = $true)]
    [object[]]$SPManagedAccount
  )
	
  foreach ($webApplication in $SPWebApplication)
  {
    $policies = $webApplication.Policies
    $managedAccountUsernames = $SPManagedAccount | Select-Object -ExpandProperty UserName
    
    foreach ($policy in $policies)
    {
      $systemUser = $managedAccountUsernames.Contains($policy.UserName)
      $properties = [ordered]@{
        'DisplayName'      = $webApplication.DisplayName
        'UserDisplayName'  = $policy.DisplayName
        'Username'         = $policy.UserName
        'PolicyRoleBindings' = $policy.PolicyRoleBindings.Name
        'SystemUser'       = $systemUser
      }
      $output = New-Object -TypeName PSObject -Property $properties
		
      Write-Output -InputObject $output
    }
  }
}
