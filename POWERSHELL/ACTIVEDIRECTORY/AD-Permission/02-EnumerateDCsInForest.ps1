<#
.SYNOPSIS
02-EnumerateDCsInForest.ps1 enumerates all domain controllers in a forest.

.DESCRIPTION
This script collects data about each domain controller in a forest and stores this information in 
02_allDCsInForest.csv file.

Input Files:  none

Output Files: 

- 02_allDCsInForest.csv - list of all detected domain controllers and their properties:
--- ComputerObjectDN
--- defaultPartition
--- Domain
--- Enabled
--- Forest
--- HostName
--- InvocationId
--- IPv4Address
--- IPv6Address
--- IsGlobalCatalog
--- IsReadOnly
--- LdapPort
--- Name
--- NTDSSettingsObjectDN
--- OperatingSystem
--- OperatingSystemHotfix
--- OperatingSystemServicePack
--- OperatingSystemVersion
--- OperationMasterRoles
--- Oartitions
--- OerverObjectDN
--- OerverObjectGuid
--- Oite
--- OslPort


- 02_allDCsInForestErrorsLog.csv - errors Log:
--- timeStamp
--- forestName	
--- forestMode	
--- domainName
--- errorMessage

The 02_allDCsInForest.csv file is an input file for scripts listed below:
- 03-EnumerateComputersInForest.ps1

.PARAMETER 
The scipt doesn't have parameters
#>

#function createForestDomainDCsTable($csvReportFile, $csvErrorsLog) {
function createForestDomainDCsTable{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
    [string]$csvReportFile,
    [string]$csvErrorsLog 
    )

    $objOutputObject = [PSCustomObject] @{
        timeStamp=$null
        forestName = $null
        forestMode = $null
        domainName = $null
        errorMessage="could not be contacted"
    }

    $forest=Get-ADForest

    $objOutputObject.forestName=$forest.Name
    $objOutputObject.forestMode=$forest.ForestMode
    $domainNum=0
    
    $allDomains=(Get-ADForest).domains
    foreach($domain in $allDomains) {
        $objOutputObject.domainName=$domain
        $Error.Clear()

        $currTime=Get-Date -Format  hh:mm:ss.fff
        Write-Output "`n$currTime : Attempt to access a $domain" | Out-Host
        Write-Output "`n## `t DOMAIN `t `t `t DC" | Out-Host
        #$dc=Get-ADDomainController -Discover -DomainName $domain -ea SilentlyContinue
        Get-ADDomainController -Discover -DomainName $domain -ea SilentlyContinue
        if ($Error.count -eq 0) {
            $dcs=Get-ADDomainController -Filter * -Server $domain
            [string]$hn=$dcs.HostName
            Write-Output "$domainNum) `t $domain `t $hn" | Out-Host
            $dcs | Export-Csv $csvReportFile -Force -NoTypeInformation -Append
        }
        else {
            Write-Output "$domainNum) `t $domain `t could not be contacted" | Out-Host
            $objOutputObject.timeStamp=Get-Date -Format  hh:mm:ss.fff
            $objOutputObject | Export-Csv $csvErrorsLog -Force -NoTypeInformation -Append
        }         
        $domainNum++
        Write-Output "----------------------------------------------------------------" | Out-Host
    }
    $currTime=Get-Date -Format  hh:mm:ss.fff
    Write-Output "$currTime : Task is completed" | Out-Host
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

$csvReportFile=$path2Files+"02_allDCsInForest.csv"
$csvErrorsLog=$path2Files+"02_allDCsInForestErrorsLog.csv"
delFile $csvReportFile
delFile $csvErrorsLog

createForestDomainDCsTable $csvReportFile $csvErrorsLog
Write-Output "`nPath to report file: $csvReportFile" | Out-Host
Write-Output "`nPath to log file: $csvErrorsLog" | Out-Host
"DONE!"