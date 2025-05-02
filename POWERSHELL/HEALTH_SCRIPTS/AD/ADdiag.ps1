#requires -Version 2

<#
    Usage: ADDiag.ps1 [-Report ".\file.htm"] [-TimeOut n (seconds)] [-DCDIAGTests "desired tests"] [-MinFreeSpace n (in GB)] [-NoOpenReport]
           See defaults in "param" section
    
    Author:         Damian Bermatov
    Name:           ADDiag.ps1
    Description:    This script runs most common Active Directory tests (specially DCDIAG) and creates a HTML report
    Created:        170430
    Last review:    170518
    Notes:          Use the optional parameters to configure the report file, timeout, temporary file, and more. You can find the default configuration in th "param" section.
    References:     In order to get a more details, plese run "dcdiag.exe /?" or visit the dcdiag web site"
    Thanks to Martin Schvartzman
    170430 - D@mian

=[ DISCLAIMER ]===============================================================================================================
 This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
 THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  
 We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object
 code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software 
 product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the 
 Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or 
 lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.
=[/ DISCLAIMER ]===============================================================================================================
#>

param(
    $Report = ".\ADDiag.htm"  # Report file name
    ,$Timeout = "120"         # DCDIAG timeout in seconds for each test. Minimum recommended: 120
    ,$DCDIAGTests = @("Advertising","CheckSDRefDom","CheckSecurityError","CrossRefValidation","CutoffServers","DNS","FRSEvent","DFSREvent","SysVolCheck","LocatorCheck","InterSite","KCCEvent","KnowsOfRoleHolders","MachineAccount","NCSecDesc","NetLogons","ObjectsReplicated","OutboundSecureChannels","RegisterInDNS","Replications","RIDManager","Services","SystemLog","Topology","VerifyEnterpriseReferences","VerifyReferences","VerifyReplicas")
    ,$MinFreeSpace = "15"     # Minimum free space considered as good
    ,[switch]$NoOpenReport  # Will not open the HTML report on finish
)

#region =[ Config ]=========================================================================================================================
$HtmlLineFormat = "<td colspan='8' bgcolor= '{0}' align=center><font face='Arial' color='Black' size='2'>{1}</font></td>{2}"
$SvcCount = "0"
Clear-Host
Write-Host " "
Write-Host "The report will be saved as" $Report
Write-Host "The configured timeout for each test is $Timeout seconds"
if ($NoOpenReport){
    Write-Host "As configured, the report will be saved but not will be opened after finishing all the tests"
} else {
    Write-Host "The report will be opened after finishing all the tests because NoOpenReport parameter was not found"
}
Write-Host " "
Write-Host "- - - - - - - - - - - - - - - -"

#endregion =[ Config ]=======================================================================================================================

#region =[ Helper Functions ]================================================================================================================
function AddToLog {
    param(
        $Message,
        [ValidateSet('Success','Failed','TimeOut')]$Status = 'Success',
        $Prefix = '',
        [switch]$NoCloseTR
    )
    $append = if($NoCloseTR) { '' } else { '</tr>' }
    $ConsoleColors = @{Success = 'Green'; Failed = 'Red'; TimeOut = 'Yellow'}
    $HtmlColors = @{Success = '#00cc33'; Failed = '#ff3333'; TimeOut = '#F7FE2E'}
    Write-Host -Object $Message -ForegroundColor $ConsoleColors[$Status]
    $Prefix + ($HtmlLineFormat -f  $HtmlColors[$Status], $Message, $append) | Out-File -FilePath $Report -Append
}
#endregion =[ Helper Functions ]================================================================================================================

#region =[ HTML 1/2 ]======================================================================================================================
$Now = Get-Date
@"
<html>
    <head>
        <meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
        <title>ADDS report</title>
        <STYLE TYPE='text/css'>
            <!--
            td {font-family: Arial; font-size: 12px; border: 0px; padding-top: 5px; padding-right: 5px; padding-bottom: 5px; padding-left: 5px;} 
            body { margin-left: 5px; margin-top: 5px; margin-right: 5px; margin-bottom: 5px; table {border: thin solid #000000;}
            --> 
        </style> 
    </head>
    <body> 
        <table width='100%'>
            <tr bgcolor='#000099'>
                <td colspan='10' height='25' align='center'>
                    <font face='Arial' color='White' size='4'><strong>Active Directory Directory Services diagnosis report, $Now</strong></font>
                </td>
            </tr>
            </table>
    <table width='100%'>
        <tr bgcolor='Blue'>
            <td width='10%' align='center'><font face='Arial' color='White' size='2'>Name</td>
            <td width='20%' align='center'><font face='Arial' color='White' size='2'>Test name</td>
            <td colspan='8' width='70%' align='center'><font face='Arial' color='White' size='2'>Result</td>
        </tr>
"@ | Out-File -FilePath $Report -Force
#endregion =[ HTML 1/2 ]=====================================================================================================================

#region =[ Domains checks ]================================================================================================================
If (-not (Get-module ActiveDirectory)) {
    Import-Module ActiveDirectory -ErrorAction Stop
}
$Domains = Get-ADForest | Select-Object -ExpandProperty Domains
Write-Host " "
Write-Host "Starting general checks"
Write-Host "- - - - - - - - - - - - - - - -"

#region =[ Repadmin showbackup check ]=====================================================================================================
Write-Host "Checking last succesfully backup registered"
@"
    <tr colspan='8'>
        <td ></td>
        <td bgcolor='#5882FA' align=center><font face='Arial' color='White' size='2'>Last registered backup</td>
        <td colspan='8' bgcolor='#E0E0F8' align=center><font face='Arial' color='Black' size='2'> repadmin /showbackup</td>
    </tr>
"@ | Out-File -FilePath $Report -Append
$Count = 0
$ReadResults = & repadmin.exe /showbackup
foreach ($ReadResult in $ReadResults){
    $Count++
    if ($Count -gt 10){
        if ($ReadResult.Length -gt 1){
            $ReadResult = $ReadResult -replace "dSASignature"," "
            $ReadResult = $ReadResult -replace "\s+","</td><td>"
            $ReadResult = $ReadResult -replace "<td>1</td>"," "
            $ReadResult = $ReadResult -replace "<td></td>"," "
            @"
                <tr>
                    <td></td>
                    <td></td>
                    <td width='10%'>$ReadResult</td>
                </tr>
"@ | Out-File -FilePath $Report -Append
        }
    }
}
#endregion =[ Repadmin showbackup check ]====================================================================================================

#region =[ FSMO roles holders ]==============================================================================================================
Write-Host "Checking FSMO holders"
@"
    <tr colspan='8'>
        <td ></td>
        <td bgcolor='#5882FA' align=center><font face='Arial' color='White' size='2'>FSMO roles holders</td>
        <td colspan='8' bgcolor='#E0E0F8' align=center><font face='Arial' color='Black' size='2'> netdom query FSMO</td>
    </tr>
"@ | Out-File -FilePath $Report -Append
$Count = 0
$ReadResults = & netdom.exe query FSMO
foreach ($ReadResult in $ReadResults){
    $Count++
    if ($Count -lt 7){
        if ($Count -eq 6){
            @"
                <tr>
                    <td></td>
                    <td></td>
"@ | Out-File -FilePath $Report -Append
            if ($ReadResult -like "The command completed successfully."){
                AddToLog -Message "Netdom query FSMO completed successfully" -Status Success
            } else {
                AddToLog -Message "Netdom query FSMO failed" -Status Failed
            }
            
        } else {
            if ($ReadResult.Length -gt 1){
                $ReadResult = $ReadResult -replace "\s+"," "
                $ReadResult = $ReadResult -replace "<td></td>"," "
                @"
                    <tr>
                        <td></td>
                        <td></td>
                        <td>$ReadResult</td>
                    </tr>
"@ | Out-File -FilePath $Report -Append
            }
        }
    }
}
#endregion =[ FSMO roles holders ]=========================================================================================================

#region =[ Repadmin replsum check ]========================================================================================================
Write-Host "Checking replication information"
@"
    <tr colspan='8'>
        <td ></td>
        <td bgcolor='#5882FA' align=center><font face='Arial' color='White' size='2'>Replication information</td>
        <td colspan='8' bgcolor='#E0E0F8' align=center><font face='Arial' color='Black' size='2'> repadmin /replsum</td>
    </tr>
"@ | Out-File -FilePath $Report -Append
$Count = 0
$ReadResults = & repadmin.exe /replsum
foreach ($ReadResult in $ReadResults){
    $Count++
    if ($Count -gt 9){
        if ($ReadResult.Length -gt 1){
            $ReadResult = $ReadResult -replace "/"," "
            $ReadResult = $ReadResult -replace "\s+","</td><td>"
            $ReadResult = $ReadResult -replace "<td></td>"," "
            $ReadResult = $ReadResult -replace "largest</td><td>delta","Largest delta"
            $ReadResult = $ReadResult -replace "Source</td><td>DSA","<td>Source DSA"
            $ReadResult = $ReadResult -replace "Destination</td><td>DSA","<td>Destination DSA"
            @"
                <tr>
                    <td></td>
                    <td></td>
                    $ReadResult
                </tr>
"@ | Out-File -FilePath $Report -Append
        }
    }
}
#endregion =[ Repadmin replsum check ]=======================================================================================================

foreach ($Domain in $Domains){
    Write-Host "Starting checks on $Domain"
    $DomMode = (Get-ADDomain).DomainMode
    @"
        <tr>
            <td bgcolor='#808080' align=center><font face='Arial' color='White' size='2'>$Domain</td>
            <td bgcolor= '#E0E0F8' align=center><font face='Arial' color='Black' size='2'>Domain functional level</td>
            <td colspan='8' bgcolor= '#E0E0F8' align=center><font face='Arial' color='Black' size='2'>$DomMode</td>
        </tr>
        <tr>
            <td></td>
            <td bgcolor='#5882FA' align=center><font face='Arial' color='White' size='2'>Replication failures</td>
"@ | Out-File -FilePath $Report -Append
    if ((Get-command "Get-ADReplicationFailure" -ErrorAction SilentlyContinue).Length -ge 1){
        $ReplFails = (Get-ADReplicationFailure -Target $Domain -ErrorAction SilentlyContinue).failureCount
        Write-Host "Checking replication errors"
        if ($ReplFails -eq 0){
            AddToLog -Message "No replication errors found" -Status Success
        } else {
            AddToLog -Message "The replication errors counter is not 0, consider running Powershell Get-ADReplicationFailure" -Status TimeOut
        }
    } else {
        AddToLog -Message "Unable to Get-ADReplicationFailure because it needs at least Windows Server 2012R2" -Status TimeOut
    }
}
#endregion =[ Domains checks ]===================================================================================================================

#region =[ DCs Checks ]========================================================================================================================
$DCs = (Get-ADForest).Domains | %{ Get-ADDomainController -Filter * -Server $_ }
foreach ($DC in $DCs){
    $DCName = $DC.HostName.ToString()
    Write-Host " "
    Write-Host "Starting checks on $DC.hostname"
    Write-Host "- - - - - - - - - - - - - - - -"
    @"
    <tr>
        <td bgcolor='#808080' align=center><font face='Arial' color='White' size='2'>$DCName</td>
"@ | Out-File -FilePath $Report -Append
    
    #region =[ Connectivity check ]=========================================================================================================
    if (Test-Connection -ComputerName $DC.hostname -Count 1 -ErrorAction SilentlyContinue) {
        Write-Host `t Connectivity check passed -ForegroundColor Green
        $OS = Get-WmiObject -Class win32_OperatingSystem -ComputerName $DC.hostname -ErrorAction SilentlyContinue
        $Lastboot = $OS.ConvertToDateTime($os.LastBootUpTime)
        $OS = $OS.Caption
        AddToLog -Message 'Ping is ok' -Status Success -Prefix @"
                <td bgcolor='#E0E0F8' align=center><font face='Arial' color='Black' size='2'>Operating system</td>
                <td colspan='8' bgcolor= '#E0E0F8' align=center><font face='Arial' color='Black' size='2'>$OS</td>
            </tr>
            <tr>
                <td></td>
                <td bgcolor='#E0E0F8' align=center><font face='Arial' color='Black' size='2'>Last boot time</td>
                <td colspan='8' bgcolor= '#E0E0F8' align=center><font face='Arial' color='Black' size='2'>$Lastboot</td>
            </tr>
            <tr>
                <td></td>
                <td bgcolor='#5882FA' align=center><font face='Arial' color='White' size='2'>Ping check</td>
"@ 
    #endregion =[ Connectivity check ]========================================================================================================

        #region =[ Logical disks check ]====================================================================================================
        $Disks = Get-WmiObject -Class win32_logicaldisk -ComputerName $DC.hostname -Filter 'DriveType=3' -ErrorAction SilentlyContinue
        foreach ($Disk in $Disks){
            $DiskDrive = $Disk.DeviceID.ToString()
            @"
                <tr>
                <td></td>
                <td bgcolor='#5882FA' align=center><font face='Arial' color='White' size='2'>Drive $DiskDrive</td>
"@ | Out-File -FilePath $Report -Append
            $DiskFreeSpace = $Disk.FreeSpace/1GB
            if ($DiskFreeSpace -gt $MinFreeSpace){
                AddToLog -Message $DiskFreeSpace -Status Success
            } else {
                AddToLog -Message $DiskFreeSpace -Status Failed
            }
        }
        #endregion =[ Logical disks check ]====================================================================================================

        #region =[ Services check ]==========================================================================================================
        $Services = Get-WMIObject -Class Win32_Service -Filter "State='Stopped'" -ComputerName $DC.hostname -ErrorAction SilentlyContinue
        foreach($service in $Services) {
            $Log = "
                <tr>
                    <td></td>
                    <td bgcolor='#5882FA' align=center><font face='Arial' color='White' size='2'>Services check</td>"
            if(-not (($Service.exitcode -eq 0) -or ($Service.exitcode -eq 1077))){
                $SvcCount = "1"
                $Svc = ($service.Displayname).tostring()
                AddToLog -Message $Svc -Status Failed -Prefix $Log
            }
        }
        if ($SvcCount -lt 1){

            AddToLog -Message 'No services with error exit codes where found' -Status Success -Prefix $Log
        }
        $SvcCount = "0"
        #endregion =[ Services check ]========================================================================================================

        #region =[ DCDIAG check ]===========================================================================================================
        foreach ($DCDIAGTest in $DCDIAGTests){
            write-host "Running $DCDIAGTest test on"$DC.HostName
            if ($DCDIAGTest -like "DNS"){
                Write-Host "Usually DNS test takes more time than other tests. Please wait and/ or consider reviewing 'Timeout' configuration in 'param section'"
            }
            $Job = start-job -Name ADdiag -scriptblock {dcdiag.exe /s:$($args[0]) /test:$($args[1])} -ArgumentList $DC.hostname,$DCDIAGTest
            Wait-Job -Name ADdiag -Timeout $Timeout | Out-Null
            $Log = "
                <tr>
                <td></td>
                <td bgcolor='#5882FA' align=center><font face='Arial' color='White' size='2'>DCDIAG $DCDIAGTest</td>"
            if($Job.State -eq 'Running'){
                AddToLog -Message "$DCDIAGTest timeout, consider running: dcdiag.exe /test:$DCDIAGtest" -Status TimeOut -Prefix $Log
                Stop-Job -Name ADdiag -ErrorAction SilentlyContinue
            } else {
                $ReadResults = (Receive-Job -Name ADdiag -Keep) -match "test $DCDIAGTest"
                foreach ($ReadResult in $ReadResults){
                    if ($ReadResult -match '\.+\s(?<DomainName>.*)\spassed\stest\s(?<TestName>.*)'){
                        AddToLog -Message $DCDIAGTest -Status Success -Prefix $Log
                    } else {
                        AddToLog -Message "$DCDIAGTest failed, consider running: dcdiag.exe /test:$DCDIAGtest" -Status Failed -Prefix $Log
                    }
                }
            }
            Remove-Job -Name ADdiag -Force -ErrorAction SilentlyContinue
        }
        #endregion =[ DCDIAG check ]==============================================================================================================

    } else {
        AddToLog -Message 'No ping' -Status Failed -Prefix @"
                <td></td>
                <td></td>
            </tr>
            <tr>
                <td></td>
                <td bgcolor='#5882FA' align=center><font face='Arial' color='White' size='2'>Ping check</td>
"@
    }
}
#region =[ Checks ]==============================================================================================================================

#region =[ HTML 2/2 ]============================================================================================================================
@"
            <tr>
                <td colspan='10' height='25' align='center'>
                    <font face='Arial' color='Black' size='2'>For more details about 'dcdiag checks' consider visiting its web site'</font>
                </td>
            </tr>
            <tr>
            </tr>
        </table>
    </body>
</html>
"@ | Out-File -FilePath $Report -Append
#endregion =[ HTML 2/2 ]============================================================================================================================

#region =[ End section ]==========================================================================================================================
Get-Job -Name ADdiag -ErrorAction SilentlyContinue | Remove-Job -Force -ErrorAction SilentlyContinue
Write-Host "Finished" -ForegroundColor Blue
if (-not($NoOpenReport)){
    Invoke-Item $Report
    }
$Report = ""
$Log = ""
#endregion =[ End section ]=========================================================================================================================
