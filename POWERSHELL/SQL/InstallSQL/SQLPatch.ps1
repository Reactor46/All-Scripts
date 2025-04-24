#import-module D:\Powershell\Test\common.psm1
. ./function.ps1

#######################################
# Starting SQL SP Installation #
#######################################
$hostName = get-content env:computername
$servicePackExec1=get-SQlserverpatchmedia
#$servicePackExec="'"+$servicePackExec1+"'"
$servicePackExec=$servicePackExec1 -replace ' ', '` '
$instance=Get-instacneName
$SQLserverinstanceName=$hostName+"\"+$instance.ToUpper()
write-host "InstanceName is:" $SQLserverinstanceName


"Starting Service Pack Installation..."
$patchCmd = $servicePackExec+" /Action=Patch /Quiet /IAcceptSQLServerLicenseTerms /Instancename=""$instance"""

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
