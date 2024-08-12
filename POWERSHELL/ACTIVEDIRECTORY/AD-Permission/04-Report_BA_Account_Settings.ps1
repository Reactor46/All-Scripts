<#
.SYNOPSIS
04-Report_BA_Account_Settings.ps1 checks Buit-in Domain Administrator (BA) account properties 
in each Domain in a forest.

.DESCRIPTION
This script collects values of BA accounts' properties listed below  
- Enabled;
- SmartcardLogonRequired;
- AccountNotDelegatedeach 
on every domain in a forest and stores this information in 04_BA_DomAdminAccountsProps.csv file.

Input Files:

- 01_allDomainsInForest.csv - created by 01-EnumerateDomainsInForest.ps1 script

Output Files: 

- 04_BA_DomAdminAccountsProps.csv - properties of BA domain Administrator accounts in a forest:
--- Domain
--- SamAccountName
--- Enabled
--- SmartcardLogonRequired
--- AccountNotDelegated

- 04_BA_DomAdminAccountsPropsLOG - errors Log:
--- timeStamp
--- domainName
--- userName
--- logText

.PARAMETER 
The scipt doesn't have parameters
#>

#function reportADUserAccountProperties($userAcc,$domNme,$csvReportFile, $csvLog)
function reportADUserAccountProperties{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
    [string]$userAcc,
    [string]$domNme,
    [string]$csvReportFile,
    [string]$csvLog 
    )


    $objOutputObject = [PSCustomObject] @{
        timeStamp=$null
        domainName = $null
        userName=$null
        logText=$null
    }

    try {  
        Get-ADUser -Filter {SamAccountName -eq $userAcc} -Server $domNme -Properties * |
        Select-Object  @{LABEL=’Domain’;EXPRESSION={$domNme}},SamAccountName,Enabled,SmartcardLogonRequired,AccountNotDelegated |
        Export-Csv $csvReportFile -NoTypeInformation -Append
    }
    catch  {
        $objOutputObject.timeStamp=Get-Date
        $objOutputObject.domainName=$domNme
        $objOutputObject.userName=$userAcc
        $objOutputObject.logText="Domain/Account can't be found"
        $objOutputObject | Export-Csv $csvLogFile -NoTypeInformation -Append
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

$csvInputFile=$path2Files+"01_allDomainsInForest.csv"
doesInputFileExist $csvInputFile

$csvReportFile=$path2Files+"04_BA_DomAdminAccountsProps.csv"
$csvLogFile=$path2Files+"04_BA_DomAdminAccountsPropsLOG.csv"
delFile $csvReportFile
delFile $csvLogFile

$Domains=Import-Csv $csvInputFile
$userAcc="Administrator"
foreach ($domRec in $Domains)
{
    $domNme=$domRec.domainName
    Write-Output "Domain: `t $domNme"  | Out-Host
    reportADUserAccountProperties $userAcc $domNme $csvReportFile $csvErrorsLog    
}

Write-Output "`nPath to Report File: $csvReportFile" | Out-Host
Write-Output "Path to Log File: $csvLogFile" | Out-Host
"DONE!"