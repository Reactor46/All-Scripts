# clear-host

# The section below will be an input from Service Manager once the record has been approved for deletion

$AccountToBeDeleted = Read-host -prompt "Account to be deleted...."
$deluser = Get-ADuser $AccountToBeDeleted 
$deleteuser = $deluser.name
# Display the user name and confirm that this is the account to be deleted
write-host
write-host "This is the user to be deleted" 
write-host
write-host $deleteuser -foregroundcolor "yellow"

Do {
Write-Host "
---------- Delete User: ----------
Y: Yes 	- Delete User'
N: No 	- Exit'	
----------------------------------"
$deleteuser = read-host -prompt "Select (Y/N) & press enter"
} until ($deleteuser -eq "Y" -or $deleteuser -eq "N")

# End section of manual section

if ($deleteuser -eq "Y")
{
$getdomaincontroller = get-addomaincontroller
$domaincontrollername = $getdomaincontroller.name
$profilescript = ('\\' + $domaincontrollername + '\' + 'NETLOGON\userscript' + '\') # You may need to change this dependent on the location of profiles in your organisation

$userdatadirectory = "\\Fileshare\folder" 
$userdirtodelete = ($userdatadirectory + $AccountToBeDeleted)
remove-item -Recurse -Force $userdirtodelete -ErrorAction SilentlyContinue

$userdatadirectory = "\\Fileshare\folder"
$userdirtodelete = ($userdatadirectory + $AccountToBeDeleted)
remove-item -Recurse -Force $userdirtodelete -ErrorAction SilentlyContinue

# XML file to be deleted

$xmlfiletodelete = ($profilescript + $AccountToBeDeleted + '.xml')
remove-item $xmlfiletodelete

#Mailbox section to be deleted

disable-mailbox -identity $AccountToBeDeleted -confirm:$false
remove-aduser -identity $AccountToBeDeleted -confirm:$false

write-host "AD account, Profile folder, Profile deleted & Exchange mailbox disabled"

}
elseif ($deleteuser -eq "N")
{
exit
}


