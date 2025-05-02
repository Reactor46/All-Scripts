
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
$script:logfile = ".\smtp-query-v2-$(get-date -format MMddyyHHmmss).log"            
$script:Seperator = @"

"@            
$script:loginitialized = $false            
$script:FileHeader = @"
OLD_ALIAS, OLD_EXTENSION, NEW_ALIAS, NEW_EXTENSION, OLD_SMTP_ADDRESS, NEW_SMTP_ADDRESS
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
$ext = $i.OLD_EXTENSION.Trim()
$newExt = $i.NEW_ALIAS.Trim()
$oldAlias = $i.OLD_ALIAS.Trim()
$oldSMTP = $i.OLD_SMTP_ADDRESS.Trim()
$lastName = $oldAlias.Substring(2)
$firstInit = $oldAlias.TocharArray()
$firstInit = $firstInit[0]
$filter = "(&(ipPhone=$ext)(samAccountName=$firstInit*$lastName))"
$blank = ",,,,,"
#$Props = 'Name', 'samAccountName','telephoneNumber','ipPhone'

		Write-Host "Searching ADDS for user with extension "$ext"." -foregroundcolor white
		#Get-QADUser -includeAllProperties -LdapFilter $filter | Select $Props | out-csv $csvFile
		$count = Get-QADUser -includeAllProperties -LdapFilter $filter | Measure-Object
		$x = Get-QADUser -includeAllProperties -LdapFilter $filter
		write-Host "Found"$count.count"users for this query!"
			if ($count.count -eq 0)
				{
					write-host "I could not find user with an IP Phone extension of ""$ext""in Active Directory!" -foregroundcolor magenta
					#write-log "I could not find user with an IP Phone extension of ""$($ext)""in Active Directory! A Manual phone update may be required for this user."
					write-log "$($oldAlias),$($ext),,$($ext),$($oldSMTP),"
				}
			Else
				{
					if ($count.count -gt 1)
						{
							write-Host "I found more than one record for extension ""$ext!""" -foregroundcolor red
							#write-log "I found more than one record for extension ""$($ext)"", skipping record update. Please update phone record manually for this user."
							write-log "$($oldAlias),$($ext),ERROR,MORE THAN 1 RESULT,$($oldSMTP),ERROR"
						}
					Else
						{
							write-log "$($oldAlias),$($ext),$($x.samAccountName),$($ext),$($oldSMTP),$($x.PrimarySMTPAddress)"
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