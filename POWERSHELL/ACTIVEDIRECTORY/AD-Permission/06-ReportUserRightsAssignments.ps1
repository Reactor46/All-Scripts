<#
.SYNOPSIS
06-ReportUserRightsAssignments.ps1 script reports User Rights assignment settings for 
member servers and workstations located in ServersOU and WorkstationsOU in each domain in a forest.

.DESCRIPTION
This script parses RSOP file created by 05-CreateRSOPFiles.ps1 script and stores User Rights Assignments 
in a .csv file for further analysis

Input Files:

- UserRightsAssignment.csv  - a list of User Rights Assignments that should be reported
- 05_ListOfRSOPs.csv        - list of RSOP files created by 05-CreateRSOPFiles.ps1 script
- RSOP files                - created by 05-CreateRSOPFiles.ps1 script

Output:

- 06_ReportUserRightsAssignments.csv - list of Admin accounts and UserRightsAssignments linked to these accounts:
--- compName
---	policy
--- targetAccount
--- compDistinguishedName
--- timeStamp
---	rsopFileName

.PARAMETER 
The scipt doesn't have parameters
#>


function exportUARA {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
    [string]$fileName,
    [string]$compName, 
    [Object]$userRightsAssignments, 
    [string]$csvReportFile
    )
        $objOutputObject = [PSCustomObject] @{
        compName=$null
        policy=$null
        targetAccount=$null
        compDistinguishedName=$compName
        timeStamp=Get-Date
        rsopFileName=$fileName
    }
    
    if (-Not (Test-Path $fileName)) {
        Write-Output "$fileName doesn't exist." | Out-Host
        $objOutputObject.compName="RSOP file doesn't exist."
        $objOutputObject | Export-Csv $csvReportFile -NoTypeInformation -Append
        return
    }

    $strIsNotApplied=" NOT applied."
    [string]$htmlTitles = Get-Content $fileName | Select-String -Pattern "<title>"
    [string]$compNme = $htmlTitles.Replace("<title>","")
    $compNme=$compNme.Replace("</title>",":")
    $compNme=($compNme.Split(":"))[0]

    $objOutputObject.compName=$compNme

    Write-Output "<$compNme> RSOP. User Rights Assignments." | Out-Host
    Write-Output ("====================") | Out-Host

    foreach ($uRAOBJ in $userRightsAssignments)
    {
        $uRA=$uRAOBJ.Policy
        $rec4Log="Policy: <"+$uRA+">"
        $objOutputObject.policy=$uRA

        $string2Search="<tr><td>"+$uRA+"</td><td>"
        [string]$uRAProps = Get-Content $fileName | Select-String -Pattern $string2Search
        if (($uRAProps -eq "") -or ($null -eq $uRAProps)) {

            $objOutputObject.targetAccount=$strIsNotApplied
            $objOutputObject | Export-Csv $csvReportFile -NoTypeInformation -Append

            $rec4Log=$rec4Log+$strIsNotApplied
            Write-Output ("`n$rec4Log") | Out-Host
            
        }
        else {
            Write-Output ("`n$rec4Log") | Out-Host
            $uRAProps=$uRAProps.Replace($string2Search,"")
            $uRAProps=$uRAProps.Replace("</td><td>",":")
            $usersList=($uRAProps.Split(":"))[0]

            Write-Output "Accounts:" | Out-Host
    
            $listOfAccounts=$usersList.Split(",")
   
            $i=1
            foreach ($account in $listOfAccounts){
                $strAccnt=$account.Trim()
                Write-Output ("$i $strAccnt") | Out-Host
                $objOutputObject.targetAccount=$strAccnt
                $objOutputObject | Export-Csv $csvReportFile -NoTypeInformation -Append
                $i++
            }
            Write-Output ("====================") | Out-Host
        }
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
        Write-Output "$inputFile doesn't exist. The script is STOPPED." | Out-Host
        EXIT
    }
}

Clear-Host
$path2Files=defDataFolder "Data"

$inputPolicies_File=$path2Files + "UserRightsAssignment.csv"
$csvInputFileRSOPsList=$path2Files + "05_ListOfRSOPs.csv"
$csvReportFile=$path2Files + "06_ReportUserRightsAssignments.csv"

doesInputFileExist $inputPolicies_File
doesInputFileExist $csvInputFileRSOPsList

delFile $csvReportFile

$userRightsAssignments=Import-Csv $inputPolicies_File
$rsopFiles=Import-Csv $csvInputFileRSOPsList

foreach ($rsopFile in $rsopFiles) {
    #$xxx=$rsopFile.rsopFileName
    #Write-Output "rsopFile: $xxx" | Out-Host
    #$fileName=$path2Files+$rsopFile.rsopFileName
    [string]$fileName=$rsopFile.rsopFileName
    Write-Output "fileName:$fileName" | Out-Host
    $compName=$rsopFile.compDN
    #Write-Output "compName: $compName" | Out-Host	

    exportUARA $fileName $compName $userRightsAssignments $csvReportFile
    #exportUARA -fileName $fileName -compName $compName -userRightsAssignments $userRightsAssignments -csvReportFile $csvReportFile
}
 
Write-Output "`nPath to Report File: $csvReportFile" | Out-Host

"DONE!"