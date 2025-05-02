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
$script:logfile = ".\add.member.dg-$(get-date -format MMddyyHHmmss).log"            
$script:Seperator = @"

"@            
$script:loginitialized = $false            
$script:FileHeader = @"
$("-" * 50)
***Application Information***
Filename:add.member.dg.ps1
Created by:  Shon Miles
Last Modified:  $(Get-Date -Date (get-item .\add.member.dg.ps1).LastWriteTime -f MM/dd/yyyy)
$("-" * 50)
"@

## Import data from CSV file and store it in variable 'data' array
$data = import-csv $args[0]

## Loop through data array and perform action
foreach ($i in $data)
	{
$error.clear()
## Setup some variables
$DG = "O365-20130613"
$UPN = $i.Alias.Trim()

		Write-Host "Adding" $UPN "to"$DG -foregroundcolor white
		Add-DistributionGroupMember -Identity $DG -Member $UPN -domaincontroller andc08
		
		If ($error.count -gt 0)
			{
				write-log $("*" * 50)
				write-log "There was an error trying to add $($UPN)! The error was..."
				write-log $error
				write-log $("*" * 50)
			}
		Else
			{
				write-log "$($UPN) was added to $($DG)."
			}
		

}