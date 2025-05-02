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
$script:logfile = ".\remove-smtp-address-$(get-date -format MMddyyHHmmss).log"            
$script:Seperator = @"

"@            
$script:loginitialized = $false            
$script:FileHeader = @"
$("-" * 50)
***Application Information***
Filename:remove-smtp-address.ps1
Created by:  Shon Miles
Last Modified:  $(Get-Date -Date (get-item .\remove-smtp-address.ps1).LastWriteTime -f MM/dd/yyyy)
$("-" * 50)
"@
## Import data from CSV file and store it in variable 'data' array
$data = import-csv $args[0]
foreach ($i in $data)
	{
$error.clear()
## Setup some variables
$Alias = $i.Alias.Trim()
$emailRemove = $Alias + "@dmaccwest.org"

		Write-Host "Removing SMTP email address" $emailRemove ", from" $Alias"." -foregroundcolor white
		Set-Mailbox $Alias -EmailAddresses @{remove=$emailRemove} -EmailAddressPolicyEnabled $false
			If ($error.count -gt 0)
				{
					write-log $("*" * 50)
					write-log "There was an error removing e-mail address from $($UPN)! The error was..."
					write-log $error
					write-log $("*" * 50)
				}
			Else
				{
					write-log "SMTP email, $($emailRemove), was removed from user, $($Alias)"
				}
}