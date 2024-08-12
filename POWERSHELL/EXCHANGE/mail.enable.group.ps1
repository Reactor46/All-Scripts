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
$script:logfile = ".\mail-enable-groups$(get-date -format MMddyyHHmmss).log"            
$script:Seperator = @"

"@            
$script:loginitialized = $false            
$script:FileHeader = @"
$("-" * 50)
***Application Information***
Filename:mail.enable.groups.ps1
Created by:  Shon Miles
Last Modified:  $(Get-Date -Date (get-item .\mail.enable.group.ps1).LastWriteTime -f MM/dd/yyyy)
$("-" * 50)
"@

## Import data from CSV file and store it in variable 'data' array
$data = import-csv $args[0]
$DC = "andc08.campus.dmacc.edu"

## Loop through data array and perform action
foreach ($i in $data)
	{
$error.clear()
## Setup some variables
$Group = $i.dg.Trim()

		Write-Host "Email enabling group for " $Group"!" -foregroundcolor white
		Enable-DistributionGroup $Group -domaincontroller $DC
		Write-Host "Allowing Mail flow from un-authenticated users for " $Group"!" -foregroundcolor yellow
		Set-DistributionGroup $Group -RequireSenderAuthenticationEnabled $false -domaincontroller $DC
		
		If ($error.count -gt 0)
			{
				write-log $("*" * 50)
				write-log "There was an error for $($Group)! The error was..."
				write-log $error
				write-log $("*" * 50)
			}
		Else
			{
				write-log "Email has been enabled for Group: $($Group)."
			}
		

}