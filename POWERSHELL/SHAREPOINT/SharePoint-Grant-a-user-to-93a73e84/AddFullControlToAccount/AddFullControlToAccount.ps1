$webApplicationUrl = "http://server"
$account = "domain\user"
$roleName = "FullControl"

$webApplication = Get-SPWebApplication -Identity $webApplicationUrl
$account = (New-SPClaimsPrincipal -identity $account -identitytype 1).ToEncodedString()

$zonePolicies = $webApplication.ZonePolicies("Default")
$policy = $zonePolicies.Add($account, $account)
$role = $webApplication.PolicyRoles.GetSpecialRole($roleName)
$policy.PolicyRoleBindings.Add($role)
$webApplication.Update()