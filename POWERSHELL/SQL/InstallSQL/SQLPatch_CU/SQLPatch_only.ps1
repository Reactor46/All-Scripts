#import-module D:\Powershell\Test\common.psm1
#. ./function.ps1
function get-SQlinstances 
{
$Cn=$env:COMPUTERNAME
$sqlserviceins=Get-WmiObject -ComputerName $Cn win32_service -EA stop | where {$_.name -Like 'MSSQLSERVER' -or $_.name -like 'MSSQL$*' }  

[System.Collections.ArrayList]$instancenames =@{}
foreach ($instance in $sqlserviceins)
{
    if($instance.name -eq "MSSQLServer")
        {
            
            $instancenames +=$instance.name;
        }

        else
        {
            $str=$($instance.name);
            $str=$str.split('$')[1];
            $instancenames +=$str;

        }
    }
    return $instancenames
}
#######################################
# Starting SQL SP Installation #
#######################################

$servicePackExec1=read-host "Please enter SQL server patch or SQL CU media location with .EXE extension"
#$servicePackExec="'"+$servicePackExec1+"'"
$servicePackExec=$servicePackExec1 -replace ' ', '` '

$SQLserverinstanceName= get-SQlinstances 

write-host "InstanceName is:" $SQLserverinstanceName

"Starting Service Pack or CU Installation..."

$patchCmd = $servicePackExec+" /Action=Patch /Quiet /IAcceptSQLServerLicenseTerms /Instancename=""$SQLserverinstanceName"""
$patchCmd
Invoke-Expression $patchCmd
#$script:exitcode=(start-process -Filepath $patchCmd -passthru -wait).exitcode;
<#$running=$true;
while( $running -eq $true)
{
    $running= ( ( get-process | where processName -eq "setup").Length -gt 0);
    start-sleep -s 10
}
#>
#$SQlpatch=Invoke-Expression $patchCmd

## have to take the name of the process and wait for the completion of the pid because service packs
## return prompt immediately and then run in background
$process=[System.IO.Path]::GetFileNameWithoutExtension($servicePackExec)
$nid = (Get-Process $process).id
Wait-Process -id $nid

#>
