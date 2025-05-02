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
$script:logfile = ".\ex.index.repair-$(get-date -format MMddyyHHmmss).log"            
$script:Seperator = @"

"@            
$script:loginitialized = $false            
$script:FileHeader = @"
$("-" * 50)
***Application Information***
Filename:ex.index.repair.ps1
Created by:  Shon Miles
Last Modified:  $(Get-Date -Date (get-item .\ex.index.repair.ps1).LastWriteTime -f MM/dd/yyyy)
$("-" * 50)
"@

## Import data from CSV file and store it in variable 'data' array
$data = import-csv $args[0]

## Loop through data array and perform action
foreach ($i in $data)
	{
$error.clear()
## Setup some variables
$indexName = $i.Name.Trim()
$CIS = $i.ContentIndexState.Trim()

		Write-Host "Checking for failed indexes on" $indexName "!" -foregroundcolor white
		switch ($CIS)
			{
				"Failed" 	{
								Write-Host $indexName "has a failed index, I will repair it for you!" -foregroundcolor yellow
								Update-MailboxDatabaseCopy $indexName -CatalogOnly
								If ($error.count -gt 0)
									{
										write-log $("*" * 50)
										write-log "The error was..."
										write-log $error
										write-log $("*" * 50)
									}
								Else
									{
										write-log "Success: $($indexName)."
									}
							}
				"Healthy"	{Write-Host $indexName "has a healthy index!" -foregroundcolor white}
				default 	{Write-Host "I was not able to detemine the content index state for " $indexName}
			}
		
		#Set-QADGroup -Identity $Group -GroupScope 'Universal' -ea SilentlyContinue
}