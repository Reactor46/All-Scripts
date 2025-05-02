<#
.SYNOPSIS
05-CreateRSOPFiles.ps1 script creates RSOPs for target computers in a forest.

.DESCRIPTION
This script create RSOPs for:
-- any accessible computer in 
-- each ServersOU and WorkstationsOU in 
-- each clientOU in
-- each accessible Domain in
-- Forest

Warning: GPMC should be installed on an auditor computer

Input Files:

- 01_allDomainsInForest.csv    - created by 01-EnumerateDomainsInForest.ps1 script
- 03_allComputersInForest.csv  - created by 03-EnumerateComputersInForest.ps1 script

Output:

- RSOP files - RSOP files in HTML format (one for each target computer)

- 05_ListOfRSOPs.csv - list of target computers and their RSOP files:
--- timeStamp
---	compDN
---	rsopFileName
 

Output files are input files for 06-ReportUserRightsAssignments.ps1

.PARAMETER 
The scipt doesn't have parameters
#>
function logRSOP{
[CmdletBinding(SupportsShouldProcess=$true)]
    Param(
    [string]$csvRepFile,
    [string]$cmpDN 
    )

    $objOutputObject = [PSCustomObject] @{
        timeStamp=$null
        compDN=$null
        rsopFileName=$null
    }
    $objOutputObject.timeStamp=Get-Date
    $objOutputObject.compDN  =  $cmpDN
    $objOutputObject.rsopFileName  =  $script:rsopFile
    $objOutputObject | Export-Csv $csvRepFile -NoTypeInformation -Append

}

function FindClientOU($compName,$compDN){
    $deleteThisFromDN="CN="+$compName+","
    $ou=$compDN.Replace($deleteThisFromDN, "")
    return $ou
}

function CheckClientServersOU ($compType, $compName, $compDN){
    #Write-Output "CheckClientServersOU: $compType, $compName, $compDN"  | Out-Host
    #$clientOU=FindClientOU $compName $compDN
    #Write-Output "CheckClientServersOU: $clientOU"  | Out-Host
    if ($compType -eq "Server") {
        $retValue=$script:clientServersOUs -contains $clientOU
    }
    if ($compType -eq "Workstation") {
        $retValue=$script:clientWorkstationsOUs -contains $clientOU
    }

    return $retValue
}

function RegisterClientOU ($compType, $compName, $compDN){
    $clientOU=FindClientOU $compName $compDN
    if ($compType -eq "Server") {
        $script:clientServersOUs+=$clientOU
    }
    if ($compType -eq "Workstation") {
        $script:clientWorkstationsOUs+=$clientOU
    }
}

function testIsComputerAccessible ($compName) {
    #Write-Output "testIsComputerAccessible: $compName"  | Out-Host
    
    $hostStatus=$false
    $s=$compName.trim()
    If (Test-Connection -ComputerName $s -Count 2 -Quiet) {
        $hostStatus=$true
    }
    #Write-Output "testIsComputerAccessible: $compName `t $hostStatus"  | Out-Host
    return $hostStatus
    
    #return $true
}

function CheckComputerType ($compOS) {
    #Write-Output "CheckComputerType: $compOS"  | Out-Host
    if ($compOS.Contains("Windows")) {
        if ($compOS.Contains("Server")) {
            return "Server"
        }
        else {
            return "Workstation"
        }
    }
    else {
        return "UnknownOS"
    }
}


function ProcessRSOP($path2RSOPfile,$compFQDN){
  
    $script:rsopFile=$path2RSOPfile+"RSOP_"+$compFQDN.Replace(".","_") + ".HTML"

    Write-Output "`nRSOP:$compFQDN" | Out-Host
    Write-Output "File:$script:rsopFile" | Out-Host
    $RepType="HTML"
    $status=$true


    $Error.Clear()
    Get-GPResultantSetOfPolicy -Computer  $compFQDN -ReportType $RepType -Path $script:rsopFile -ea SilentlyContinue
    if ($Error.count -ne 0) {
            Write-Output "Logon to $compFQDN" | Out-Host
            $Error.Clear()
            Enter-PSSession -ComputerName $compFQDN #-Credential $UserCredential
            if ($Error.count -ne 0) {
                Write-Output "Can't open a session with $compFQDN" | Out-Host
                $status=$false
            }
            else {       
                gpresult  /S $compFQDN /SCOPE Computer /USER administrator /H $script:rsopFile /F
                if ($Error.count -ne 0) {
                    Write-Output "Can't run 'gpresult' on $compFQDN" | Out-Host
                    $status=$false
                }
                Exit-PSSession 
            }
    }
 

    return $status
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
        Write-Output "$inputFile doesn't exist. The script is STOPPED." | Out-Host
        EXIT
    }
}

Clear-Host

$path2Files=defDataFolder "Data"

$csvInputFileDomains=$path2Files + "01_allDomainsInForest.csv"
$csvInputFileComputers=$path2Files + "03_allComputersInForest.csv"
$csvReportFile=$path2Files + "05_ListOfRSOPs.csv"

$sOU="ServersOU"
$wOU="WorkstationsOU"
$script:clientServersOUs=@()
$script:clientWorkstationsOUs=@()
$script:rsopFile=$null         #Is used in functions: logRSOP and ProcessRSOP

doesInputFileExist $csvInputFileDomains
doesInputFileExist $csvInputFileComputers

$Domains=Import-Csv $csvInputFileDomains
$compsInForest=Import-Csv $csvInputFileComputers

delFile $csvReportFile
#$UserCredential = Get-Credential

foreach ($domRec in $Domains) {
    $domNme="DC="+$domRec.domainName
    $domNme=$domNme.Replace(".",",DC=")
    Write-Output "`ndomainName: $domNme / $domNme0" | Out-Host
 
    $compsInDomain=$compsInForest  | Where-Object {$_.Domain -eq $domNme}
    #"***************"
    #$compsInDomain.DistinguishedName
    #"-----------------------------------"
    $compsInServersAndWorkstationsOUs=$compsInDomain | 
    Where-Object {($_.DistinguishedName -match $sOU) -or ($_.DistinguishedName -match $wOU)}
    #$compsInServersAndWorkstationsOUs.DistinguishedName
    #"==================================="
    
    foreach ($comp in $compsInServersAndWorkstationsOUs) {
        $isComputerAccessible=testIsComputerAccessible $comp.Name
        if ($isComputerAccessible) {
            $compType=CheckComputerType $comp.OperatingSystem
            if (($compType -eq "Server") -or ($compType -eq "Workstation")) {
                $isOURegistered=CheckClientServersOU $compType $comp.Name $comp.DistinguishedName
                if ($isOURegistered -eq $false) {
                    $compFQDN=$comp.Name+"."+$domRec.domainName
                    $isReportCreated=ProcessRSOP $path2Files $compFQDN #$UserCredential
                    if ($isReportCreated) {
                        logRSOP $csvReportFile $comp.DistinguishedName
                        RegisterClientOU $compType $comp.Name $comp.DistinguishedName
                    }
                }
            }
        }
    }
    
}
"*********************************************"

Write-Output "`nPath to Report File: $csvReportFile" | Out-Host

"DONE!"