$errorView = "CategoryView"
## Logging setup
function write-log([string]$info){            
    if($loginitialized -eq $false){            
        $FileHeader > $logfile            
        $script:loginitialized = $True            
    }            
    $info >> $logfile            
}
## Logfile Info           
$script:logfile = ".\update-phone-$(get-date -format MMddyyHHmmss).log"            
$script:Seperator = @"

"@            
$script:loginitialized = $false            
$script:FileHeader = @"
$("-" * 50)
***Application Information***
Filename:update-phone.ps1
Created by:  Shon Miles
Last Modified:  $(Get-Date -Date (get-item .\update-phone.ps1).LastWriteTime -f MM/dd/yyyy)
$("-" * 50)
"@

## Import data from CSV file and store it in variable 'data' array
$data = import-csv $args[0]

## Loop through data array and perform action
foreach ($i in $data)
	{
$error.clear()## Setup some variables
$ipPhone = $i.EXTENSION.Trim()
$lastName = $i.LASTNAME.Trim()
$firstName = $i.FIRSTNAME.Trim()
$displayName = $lastName + ", " + $firstName
$uAlias = $i.Alias.Trim()
$filter ="(&(sn=$lastName)(givenName=$firstName))"
$Props = 'Name', 'samAccountName','telephoneNumber','ipPhone'

		Write-Host "Searching for ""$displayName""...and updating IP Phone to" $ipPhone -foregroundcolor yellow
		$count = Get-QADUser -includeAllProperties -LdapFilter $filter | Measure-Object
		$x = Get-QADUser -includeAllProperties -LdapFilter $filter
			if ($count.count -eq 0)
				{
				write-host "I could not find ""$displayname"" in Active Directory!" -foregroundcolor magenta
				write-log "I could not find ""$($displayname)"" in Active Directory! A Manual phone update may be required for this user."
				}
			Else
				{
					if ($count.count -gt 1)
						{
							write-Host "I found more than one record for ""$displayName!""" -foregroundcolor red
							write-log "I found more than one record for ""$($DisplayName)"", skipping record update. Please update phone record manually for this user."
						}
					Else
						{
							#Set-QADUser $x.samAccountName -ObjectAttributes @{ipPhone=$ipPhone}
							#write-log "$($x.sn),$($x.givenName),$($x.samAccountname),$($x.telephoneNumber),$($x.ipPhone),$($ext),$($uAlias)"
							If ($error.count -gt 0)
								{
									write-log $("*" * 50)
									write-log "There was an error setting the IP Phone data for ""$($DisplayName)"" The error was..."
									write-log $error
									write-log $("*" * 50)
								}
							Else
								{
									write-log "Updated IP Phone for ""$($DisplayName)"" to $($ipPhone)."
								}
		
						}
				}
}