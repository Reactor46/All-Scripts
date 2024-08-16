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
$script:logfile = ".\get.user.mbx.db$(get-date -format MMddyyHHmmss).log"            
$script:Seperator = @"

"@            
$script:loginitialized = $false            
$script:FileHeader = @"
$("-" * 50)
***Application Information***
Filename:get.user.mbx.db.ps1
Created by:  Shon Miles
Last Modified:  $(Get-Date -Date (get-item .\get.user.mbx.db.ps1).LastWriteTime -f MM/dd/yyyy)
$("-" * 50)
"@
## Import data from CSV file and store it in variable 'data' array
$data = import-csv $args[0]

## Loop through data array and perform action
foreach ($i in $data)
	{
$error.clear()
## Setup some variables
$MailboxDB = "Students_18"
$User = $i.Logon
$userDB = get-mailbox $User | foreach {$_.Database}
		
		If ($userDB -eq $MailboxDB)
			{
				Write-Host "User is in Students_18 mailbox datbase"	
			}
		Else
			{
				Write-host "User is not in the Students_18 mailbox database"
			}
}