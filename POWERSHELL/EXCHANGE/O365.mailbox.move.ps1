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
$script:logfile = ".\iO365.mailbox.move-$(get-date -format MMddyyHHmmss).log"            
$script:Seperator = @"

"@            
$script:loginitialized = $false            
$script:FileHeader = @"
$("-" * 50)
***Application Information***
Filename:O365.mailbox.move.ps1
Created by:  Shon Miles
Last Modified:  $(Get-Date -Date (get-item .\O365.mailbox.move.ps1).LastWriteTime -f MM/dd/yyyy)
$("-" * 50)
"@

## Import data from CSV file and store it in variable 'data' array
$data = import-csv $args[0]

## Loop through data array and perform action
foreach ($i in $data)
	{
$error.clear()
## Setup some variables
$Alias = $i.Alias.Trim()
$UPN = $i.Alias.Trim() +"@dmacc.edu"
$exGuid = get-mailbox $Alias | select ExchangeGuid | ft -HideTableHeaders
		
		write-host $exGuid
		
		If ($error.count -gt 0)
			{
				write-log $("*" * 50)
				write-log "There was an error retrieving the Exchange GUID for $($Alias)! The error was..."
				write-log $error
				write-log $("*" * 50)
			}
		Else
			{
				write-log "I found the Exchange Guid for: $($Alias) it is $($exGuid)."
			}
		

}