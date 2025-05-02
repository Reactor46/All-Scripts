<#
.SYNOPSIS
03-EnumerateComputersInForest.ps1 enumerates all computers in a forest.

.DESCRIPTION
This script collects data about each Windows computer in a forest and stores this information in 
03_allComputersInForest file.

Input Files:

- 02_allDCsInForest.csv - created by 02-EnumerateDCsInForest.ps1 script

Output Files: 

- 03_allComputersInForest.csv - list of all computers in a forest and their properties:
--- TimeStamp
--- Domain
--- Name
--- Description
--- Enabled
--- DistinguishedName
--- Created
--- LastLogonDate
--- logonCount
--- LastBadPasswordAttempt
--- OperatingSystem
--- OperatingSystemServicePack
--- OperatingSystemVersion
--- sn
--- IPv4Address

- 03_allComputersInForestLog.csv - errors Log:
--- timeStamp
--- domainName
--- dcName
--- enumYN

The 03_allComputersInForest.csv file is an input file for scripts listed below:
- 05-CreateRSOPFiles.ps1

.PARAMETER 
The scipt doesn't have parameters
#>

#function createForestDomainComputersTable($DCs, $csvReportFile, $csvLogFile) {
function createForestDomainComputersTable {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
    [Object]$DCs,
    [string]$csvReportFile,
    [string]$csvLogFile 
    )

    $objOutputObject = [PSCustomObject] @{
        timeStamp=$null
        domainName = $null
        dcName=$null
        enumYN="N"
    }
    
    $n=0
    $prevDom=$null

    foreach ($DC in $DCs)
    {
        $currDomain=$DC.domain
        Write-Output "$n `t $prevDom `t $currDomain  ?"  | Out-Host
        if ($prevDom -ne $currDomain) {
            $compsEnumYN="N"
            $dcHostName=$DC.name
            $srcbase=$DC.DefaultPartition
            
            Write-Output "$n `t $prevDom <> $currDomain"  | Out-Host

            If (Test-Connection -ComputerName $dcHostName -Count 2 -Quiet) {
                $timeStamp=Get-Date
                try {
                    $adComputers=Get-ADComputer -Filter "*" -SearchBase $srcBase -Server $dcHostName  -properties * -ea STOP | 
                    Select-Object @{LABEL=’TimeStamp’;EXPRESSION={$timeStamp}},
                    @{LABEL=’Domain’;EXPRESSION={$srcBase}},
                    Name,Description,Enabled,DistinguishedName,
                    Created,LastLogonDate,logonCount,LastBadPasswordAttempt,
                    OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion,
                    sn,IPv4Address
                    $compsEnumYN="Y"
                }
                catch  {
                    Write-Output "GET-ADCOMPUTER $dcHostName ERROR!!!"  | Out-Host
                    $compsEnumYN="N"
                }

                if ($compsEnumYN -eq "Y") {
                    $adComputers | Sort-Object Domain, Name |
                    Export-Csv $csvReportFile -NoTypeInformation -Append

                    Write-Output "Computers are enumerated in: $srcbase. Domain Controller=$dcHostName" | Out-Host

                    $prevDom=$currDomain
                    $objOutputObject.timeStamp=Get-Date -Format  hh:mm:ss.fff
                    $objOutputObject.domainName=$currDomain
                    $objOutputObject.dcName=$dcHostName
                    $objOutputObject.enumYN=$compsEnumYN

                    $objOutputObject | Export-Csv $csvLogFile -NoTypeInformation -Append  
                }
            }
            else {
                Write-Output "$numDCsInDomain `t $dcHostName is Down" | Out-Host
            }
        }
        $n++
    }
    
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

function doesInputFileExist($inputFile){
    if (-Not (Test-Path $inputFile)) {
        Write-Output "$inputFile doesn't exist. The script is STOPPED."  | Out-Host
        EXIT
    }
}


Clear-Host
$path2Files=defDataFolder "Data"

$csvInputFile=$path2Files+"02_allDCsInForest.csv"
doesInputFileExist $csvInputFile

$csvReportFile=$path2Files+"03_allComputersInForest.csv"
$csvLogFile=$path2Files+"03_allComputersInForestLOG.csv"
delFile $csvReportFile
delFile $csvLogFile

$DCs=Import-Csv $csvInputFile

createForestDomainComputersTable $DCs $csvReportFile  $csvLogFile

#createForestDomainComputersTable -DCs $DCs -csvReportFile $csvReportFile -csvLogFile $csvLogFile

Write-Output "`nPath to report file: $csvReportFile" | Out-Host
Write-Output "Path to Log File: $csvLogFile" | Out-Host
"DONE!"