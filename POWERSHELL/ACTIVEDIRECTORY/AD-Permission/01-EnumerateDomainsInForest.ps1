<#
.SYNOPSIS
01-EnumerateDomainsInForest enumerates all domains in a forest.

.DESCRIPTION
This script collects data about each domain in a forest and stores this data in
01_allDomainsInForest.csv file. 

Input Files:  none

Output Files:

- 01_allDomainsInForest.csv - list of all domains and their properties:
--- timeStamp
--- forestName	
--- forestMode	
--- domainName

This file is used as input file for scripts listed below:
- 04-Report_BA_Account_Settings.ps1
- 05-CreateRSOPFiles.ps1
- 06-ReportUserRightsAssignments.ps1

.PARAMETER 
The scipt doesn't have parameters
#>


#function createForestDomainsTable($csvReportFile, $csvErrorsLog) {
function createForestDomainsTable{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
    [string]$csvReportFile,
    [string]$csvErrorsLog 
    )

    $objOutputObject = [PSCustomObject] @{
        timeStamp=Get-Date -Format  hh:mm:ss.fff
        forestName = $null
        forestMode = $null
        domainName = $null
    }

    $forest=Get-ADForest

    $objOutputObject.forestName=$forest.Name
    $objOutputObject.forestMode=$forest.ForestMode
    $domainNum=0

    $allDomains=(Get-ADForest).domains
    foreach($domain in $allDomains) {
        $objOutputObject.domainName=$domain
        Write-Output "$domainNum) `t $domain" | Out-Host
        
        $objOutputObject | Export-Csv $csvReportFile -Force -NoTypeInformation -Append
        $domainNum++
    }
    $currTime=Get-Date -Format  hh:mm:ss.fff
    Write-Output  "$currTime : Task is completed"  | Out-Host
}
function defDataFolder($dataSF){
    $myScriptPath=(Get-Variable -Value -Scope 1 MyInvocation).InvocationName
    $scriptFolder=Split-Path $myScriptPath -Parent
    $dataFolder=$scriptFolder | Join-Path -Resolve -ChildPath $dataSF
    $dataFolder=$dataFolder+"\"
    return $dataFolder
}

function delFile{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param([string]$path2File)

    if (Test-Path $path2File) {
        Remove-Item $path2File -Confirm:$true 
    }
}


Clear-Host 

$path2Files=defDataFolder "Data"

$csvReportFile=$path2Files+"01_allDomainsInForest.csv"
$csvErrorsLog=$path2Files+"01_allDomainsInForestErrorsLog.csv"

delFile $csvReportFile
delFile $csvErrorsLog

createForestDomainsTable $csvReportFile $csvErrorsLog
Write-Output "`nPath to report file: $csvReportFile" | Out-Host
Write-Output "`nPath to log file: $csvErrorsLog" | Out-Host
"DONE!"