#import-module D:\Powershell\Test\common.psm1
. ./function.ps1

#######################################
# Starting SQL SP Installation #
#######################################
$hostName = get-content env:computername
$SQLCULoc=SQLCUmedia
write-host "CU location is:$SQLCULoc"
if (!(Test-Path ($SQLCULoc)))
{
    write-host " SQL CU not required"
}
else 
{
    $servicePackExec=$SQLCULoc -replace ' ', '` '
    $instance=Get-instacneName
    $SQLserverinstanceName=$hostName+"\"+$instance.ToUpper()
    write-host "InstanceName is:" $SQLserverinstanceName

    "Starting SQL server CU Installation..."
    $patchCmd = $servicePackExec+" /Action=Patch /Quiet /IAcceptSQLServerLicenseTerms /Instancename=""$instance"""
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
}
