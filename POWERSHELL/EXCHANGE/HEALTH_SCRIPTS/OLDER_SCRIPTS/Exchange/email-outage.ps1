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
$script:logfile = ".\email-outage-$(get-date -format MMddyyHHmmss).log"            
$script:Seperator = @"

"@            
$script:loginitialized = $false            
$script:FileHeader = @"
$("-" * 50)
***Application Information***
Filename:email-outage.ps1
Created by:  Shon Miles
Last Modified:  $(Get-Date -Date (get-item .\email-outage.ps1).LastWriteTime -f MM/dd/yyyy)
$("-" * 50)
"@

## Import data from CSV file and store it in variable 'data' array
$data = import-csv $args[0]

## Loop through data array and perform action
foreach ($i in $data)
	{
$error.clear()
## Setup some variables
$User = $i.Alias.Trim()
$flag = ""

		Write-Host "Setting CustomAttribute15 for Dynamic DistributionGroup, Email Outage for"$User"!" -foregroundcolor white
		Set-Mailbox $User -CustomAttribute15 $flag
		
		If ($error.count -gt 0)
			{
				write-log $("*" * 50)
				write-log "There was an error trying to set CustomAttribute15 for$($User)! The error was..."
				write-log $error
				write-log $("*" * 50)
			}
		Else
			{
				write-log "CustomAttribute15 for Dynamic DistributionGroup, has been set for user: $($User)."
			}
		

}