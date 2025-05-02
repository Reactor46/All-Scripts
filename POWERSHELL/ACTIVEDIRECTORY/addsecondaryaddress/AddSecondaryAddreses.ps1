######################################################################
#             Author: Vikas Sukhija
#             Date: 02/27/2013
#             Description: Adding secondary email address
#####################################################################

#----ADD Exchange Shell-------

If ((Get-PSSnapin | where {$_.Name -match "Exchange.Management"}) -eq $null)
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
}


$now=Get-Date -format "dd-MMM-yyyy HH:mm"

# replace : by -

$now = $now.ToString().Replace(“:”, “-”)

$Log1 = ".\AddSecondAddress" + $now + ".log"


#-----Import CSV



$data = import-csv $args[0]

if ($? -like "False")

{

exit

}

# loop thru the malboxes now

foreach ($i in $data)

{
$secondaryemail = $i.emailaddress

$mailbox = get-mailbox $i.userid

if ($? -like "False")

{

exit

}


Write-host "$secondaryemail email address will be added to $mailbox"


#get in to multivalued attribute & add one more value

$mailbox.emailaddresses+= $secondaryemail

$mailbox | set-mailbox

Add-content  $Log1 "$secondaryemail email address added to $mailbox"

}

############################################################################