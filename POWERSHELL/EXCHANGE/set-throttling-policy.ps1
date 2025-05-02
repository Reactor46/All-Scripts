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
$script:logfile = ".\set-throttling-policy-$(get-date -format MMddyyHHmmss).log"            
$script:Seperator = @"

"@            
$script:loginitialized = $false            
$script:FileHeader = @"
$("-" * 50)
***Application Information***
Filename:set-throttling-policy.ps1
Created by:  Shon Miles
Last Modified:  $(Get-Date -Date (get-item .\set-throttling-policy.ps1).LastWriteTime -f MM/dd/yyyy)
$("-" * 50)
"@

## Import data from CSV file and store it in variable 'data' array
$data = import-csv $args[0]
$policy = "FacultyStaffRateLimit"
## Loop through data array and perform action
foreach ($i in $data)
	{
$error.clear()
## Setup some variables
$user = $i.Alias.Trim()
		## Show the user what is going on.
		Write-Host "Applying" $policy" Throttling policy to" $user"!" -foregroundcolor white
		
		## Make the update to the user.
		Set-Mailbox $user -ThrottlingPolicy $policy
		
		## Grab any errors and write them to the log.
		If ($error.count -gt 0)
			{
				write-log $("*" * 50)
				write-log "There was an error applying the throttling policy to $($user)! The error was..."
				write-log $error
				write-log $("*" * 50)
			}
		Else
			{
				write-log "Throttling policy has been applied to  $($user)."
			}
		

}