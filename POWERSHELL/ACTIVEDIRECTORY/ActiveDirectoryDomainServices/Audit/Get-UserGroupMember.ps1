Function Get-UserPrincipal($cName, $cContainer, $userName)
{
 $dsam = "System.DirectoryServices.AccountManagement" 
 $rtn = [reflection.assembly]::LoadWithPartialName($dsam)
 $cType = "domain" #context type
 $iType = "SamAccountName"
 $dsamUserPrincipal = "$dsam.userPrincipal" -as [type]
 $principalContext = new-object "$dsam.PrincipalContext"($cType,$cName,$cContainer)
 $dsamUserPrincipal::FindByIdentity($principalContext,$iType,$userName)
} # end Get-UserPrincipal


$cName = "fnbm.corp"
$cContainer = "DC=fnbm,DC=corp"

$userName = jbattista #Get-Content -Path 'C:\LazyWinAdmin\HPU.txt' -raw

$userPrincipal = Get-UserPrincipal -userName $userName -cName $cName -cContainer $cContainer

New-UnderLine -Text "Direct Group MemberShip:"
$userPrincipal.getGroups() | foreach-object { $_.name }

New-UnderLine -Text "Indirect Group Membership:"
$userPrincipal.GetAuthorizationGroups()  | foreach-object { $_.name }