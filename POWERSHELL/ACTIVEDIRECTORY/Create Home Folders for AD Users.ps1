$ADServer = 'DOMAIN' #change name to your DC


#Get Admin accountb credential

$GetAdminact = Get-Credential 

#Import Active Directory Module

Import-Module ActiveDirectory


$ADUsers = Get-ADUser -server $ADServer -Filter * -Credential $GetAdminact -Properties *

#modify display name of all users in AD (based on search criteria) to the format "LastName, FirstName Initials"

ForEach ($ADUser in $ADUsers) 
{


New-Item -ItemType Directory -Path "\\REMOTE\Users\$($ADUser.sAMAccountname)"

$Domain = 'COMPANYDOMAIN'

$UsersAm = "$Domain\$($ADUser.sAMAccountname)"

$FileSystemAccessRights = [System.Security.AccessControl.FileSystemRights]"FullControl"

$InheritanceFlags = [System.Security.AccessControl.InheritanceFlags]::"ContainerInherit", "ObjectInherit"

$PropagationFlags = [System.Security.AccessControl.PropagationFlags]::None

$AccessControl =[System.Security.AccessControl.AccessControlType]::Allow 

$NewAccessrule = New-Object System.Security.AccessControl.FileSystemAccessRule `
    ($UsersAm, $FileSystemAccessRights, $InheritanceFlags, $PropagationFlags, $AccessControl) 


$userfolder = "\\DATASERVER\Users\$($ADUser.sAMAccountname)"

$currentACL = Get-ACL -path $userfolder
$currentACL.SetAccessRule($NewAccessrule)

Set-ACL -path $userfolder -AclObject $currentACL


$homeDirectory = "\\DATASERVER\Users\$($ADUser.sAMAccountname)" #This maps the folder for each user 

$homeDrive = "U" # Drive Letter

Set-ADUser -server $ADServer -Credential $GetAdminact -Identity $ADUser.sAMAccountname -Replace @{HomeDirectory=$homeDirectory}
Set-ADUser -server $ADServer -Credential $GetAdminact -Identity $ADUser.sAMAccountname -Replace @{HomeDrive=$homeDrive}

}
  