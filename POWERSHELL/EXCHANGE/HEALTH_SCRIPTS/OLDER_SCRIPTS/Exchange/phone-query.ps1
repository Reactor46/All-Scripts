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
$script:logfile = ".\phone-query-$(get-date -format MMddyyHHmmss).log"            
$script:Seperator = @"

"@            
$script:loginitialized = $false            
$script:FileHeader = @"
LastName,FirstName,ADUser,ADTelephone,ADipPhone,UnityEXT,UnityAlias
"@

## Import data from CSV file and store it in variable 'data' array
$data = import-csv $args[0]

## Output Function ###################################################################################
$csvFile = "phone.query.csv"                                                                        ##
function out-CSV ( $csvFile, $Append = $true) {                                                     ##
	Foreach ($item in $input){                                                                      ##
		# Get all the Object Properties                                                             ##
		$Properties = $item.PsObject.get_properties()                                               ##
		# Create Empty Strings                                                                      ##
		$Headers = ""                                                                               ##
		$Values = ""                                                                                ##
		# Go over each Property and get it's Name and value                                         ##
		$Properties | %{                                                                            ##
			$Headers += $_.Name+"`t"                                                                ##
			$Values += $_.Value+"`t"                                                                ##
		}                                                                                           ##
		# Output the Object Values and Headers to the Log file                                      ##
		If($Append -and (Test-Path $csvFile)) {                                                     ##
			$Values | Out-File -Append -FilePath $csvFile -Encoding Unicode                         ##
		}                                                                                           ##
		else {                                                                                      ##
			# Used to mark it as an Powershell Custum object - you can Import it later and use it   ##
			# "#TYPE System.Management.Automation.PSCustomObject" | Out-File -FilePath $LogFile     ##
			$Headers | Out-File -FilePath $csvFile -Encoding Unicode                                ##
			$Values | Out-File -Append -FilePath $csvFile -Encoding Unicode                         ##
		}                                                                                           ##
	}                                                                                               ##
}                                                                                                   ##
## Output Function End ###############################################################################

## Loop through data array and perform action
foreach ($i in $data)
	{
$error.clear()## Setup some variables
$ext = $i.EXTENSION.Trim()
$lastName = $i.LASTNAME.Trim()
$firstName = $i.FIRSTNAME.Trim()
$displayName = $lastName + ", " + $firstName
$uAlias = $i.Alias.Trim()
$filter = "(&(sn=$lastName)(givenName=$firstName)(ipPhone=$ext))"
$Props = 'Name', 'samAccountName','telephoneNumber','ipPhone'

		Write-Host "Searching ADDS for phone numbers assigned to "$displayName"! Expecting ext:"$ext"." -foregroundcolor white
		#Get-QADUser -includeAllProperties -LdapFilter $filter | Select $Props | out-csv $csvFile
		$count = Get-QADUser -includeAllProperties -LdapFilter $filter | Measure-Object
		$x = Get-QADUser -includeAllProperties -LdapFilter $filter
		write-Host "Found"$count.count"users for this query!"
			if ($count.count -eq 0)
				{
					write-host "I could not find ""$displayname"" with an IP Phone extension of ""$ext""in Active Directory!" -foregroundcolor magenta
					write-log "I could not find ""$($displayname)"" with an IP Phone extension of ""$($ext)""in Active Directory! A Manual phone update may be required for this user."
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
							#write-log "$($x.sn),$($x.givenName),$($x.samAccountname),$($x.telephoneNumber),$($x.ipPhone),$($ext),$($uAlias)"
						}
					
				}
		If ($error.count -gt 0)
			{
				write-log $("*" * 50)
				write-log "There was an error! The error was..."
				write-log $error
				write-log $("*" * 50)
			}
		Else
			{
				#write-log "CustomAttribute15 for Dynamic DistributionGroup, has been set for user: $($User)."
			}
		

}