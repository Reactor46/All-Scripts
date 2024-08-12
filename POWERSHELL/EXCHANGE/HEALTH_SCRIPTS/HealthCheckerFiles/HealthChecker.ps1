<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 21.09.28.1529

<#
.NOTES
	Name: HealthChecker.ps1
	Requires: Exchange Management Shell and administrator rights on the target Exchange
	server as well as the local machine.
    Major Release History:
        4/20/2021  - Initial Public Release on CSS-Exchange
        11/10/2020 - Initial Public Release of version 3.
        1/18/2017 - Initial Public Release of version 2.
        3/30/2015 - Initial Public Release.

.SYNOPSIS
	Checks the target Exchange server for various configuration recommendations from the Exchange product group.
.DESCRIPTION
	This script checks the Exchange server for various configuration recommendations outlined in the
	"Exchange 2013 Performance Recommendations" section on Microsoft Docs, found here:

	https://docs.microsoft.com/en-us/exchange/exchange-2013-sizing-and-configuration-recommendations-exchange-2013-help

	Informational items are reported in Grey.  Settings found to match the recommendations are
	reported in Green.  Warnings are reported in yellow.  Settings that can cause performance
	problems are reported in red.  Please note that most of these recommendations only apply to Exchange
	2013/2016.  The script will run against Exchange 2010/2007 but the output is more limited.
.PARAMETER Server
	This optional parameter allows the target Exchange server to be specified.  If it is not the
	local server is assumed.
.PARAMETER OutputFilePath
	This optional parameter allows an output directory to be specified.  If it is not the local
	directory is assumed.  This parameter must not end in a \.  To specify the folder "logs" on
	the root of the E: drive you would use "-OutputFilePath E:\logs", not "-OutputFilePath E:\logs\".
.PARAMETER MailboxReport
	This optional parameter gives a report of the number of active and passive databases and
	mailboxes on the server.
.PARAMETER LoadBalancingReport
    This optional parameter will check the connection count of the Default Web Site for every server
    running Exchange 2013/2016 with the Client Access role in the org.  It then breaks down servers by percentage to
    give you an idea of how well the load is being balanced.
.PARAMETER CasServerList
    Used with -LoadBalancingReport.  A comma separated list of CAS servers to operate against.  Without
    this switch the report will use all 2013/2016 Client Access servers in the organization.
.PARAMETER SiteName
	Used with -LoadBalancingReport.  Specifies a site to pull CAS servers from instead of querying every server
    in the organization.
.PARAMETER XMLDirectoryPath
    Used in combination with BuildHtmlServersReport switch for the location of the HealthChecker XML files for servers
    which you want to be included in the report. Default location is the current directory.
.PARAMETER BuildHtmlServersReport
    Switch to enable the script to build the HTML report for all the servers XML results in the XMLDirectoryPath location.
.PARAMETER HtmlReportFile
    Name of the HTML output file from the BuildHtmlServersReport. Default is ExchangeAllServersReport.html
.PARAMETER DCCoreRatio
    Gathers the Exchange to DC/GC Core ratio and displays the results in the current site that the script is running in.
.PARAMETER AnalyzeDataOnly
    Switch to analyze the existing HealthChecker XML files. The results are displayed on the screen and an HTML report is generated.
.PARAMETER SkipVersionCheck
    No version check is performed when this switch is used.
.PARAMETER SaveDebugLog
    The debug log is kept even if the script is executed successfully.
.PARAMETER ScriptUpdateOnly
    Switch to check for the latest version of the script and perform an auto update. No elevated permissions or EMS are required.
.PARAMETER Verbose
	This optional parameter enables verbose logging.
.EXAMPLE
	.\HealthChecker.ps1 -Server SERVERNAME
	Run against a single remote Exchange server
.EXAMPLE
	.\HealthChecker.ps1 -Server SERVERNAME -MailboxReport -Verbose
	Run against a single remote Exchange server with verbose logging and mailbox report enabled.
.EXAMPLE
    Get-ExchangeServer | ?{$_.AdminDisplayVersion -Match "^Version 15"} | %{.\HealthChecker.ps1 -Server $_.Name}
    Run against all Exchange 2013/2016 servers in the Organization.
.EXAMPLE
    .\HealthChecker.ps1 -LoadBalancingReport
    Run a load balancing report comparing all Exchange 2013/2016 CAS servers in the Organization.
.EXAMPLE
    .\HealthChecker.ps1 -LoadBalancingReport -CasServerList CAS01,CAS02,CAS03
    Run a load balancing report comparing servers named CAS01, CAS02, and CAS03.
.LINK
    https://docs.microsoft.com/en-us/exchange/exchange-2013-sizing-and-configuration-recommendations-exchange-2013-help
    https://docs.microsoft.com/en-us/exchange/exchange-2013-virtualization-exchange-2013-help#requirements-for-hardware-virtualization
    https://docs.microsoft.com/en-us/exchange/plan-and-deploy/virtualization?view=exchserver-2019#requirements-for-hardware-virtualization
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Variables are being used')]
[CmdletBinding(DefaultParameterSetName = "HealthChecker")]
param(
    [Parameter(Mandatory = $false, ParameterSetName = "HealthChecker")]
    [Parameter(Mandatory = $false, ParameterSetName = "MailboxReport")]
    [string]$Server = ($env:COMPUTERNAME),
    [Parameter(Mandatory = $false)]
    [ValidateScript( { -not $_.ToString().EndsWith('\') })][string]$OutputFilePath = ".",
    [Parameter(Mandatory = $false, ParameterSetName = "MailboxReport")]
    [switch]$MailboxReport,
    [Parameter(Mandatory = $false, ParameterSetName = "LoadBalancingReport")]
    [switch]$LoadBalancingReport,
    [Parameter(Mandatory = $false, ParameterSetName = "LoadBalancingReport")]
    [array]$CasServerList = $null,
    [Parameter(Mandatory = $false, ParameterSetName = "LoadBalancingReport")]
    [string]$SiteName = ([string]::Empty),
    [Parameter(Mandatory = $false, ParameterSetName = "HTMLReport")]
    [Parameter(Mandatory = $false, ParameterSetName = "AnalyzeDataOnly")]
    [ValidateScript( { -not $_.ToString().EndsWith('\') })][string]$XMLDirectoryPath = ".",
    [Parameter(Mandatory = $false, ParameterSetName = "HTMLReport")]
    [switch]$BuildHtmlServersReport,
    [Parameter(Mandatory = $false, ParameterSetName = "HTMLReport")]
    [string]$HtmlReportFile = "ExchangeAllServersReport.html",
    [Parameter(Mandatory = $false, ParameterSetName = "DCCoreReport")]
    [switch]$DCCoreRatio,
    [Parameter(Mandatory = $false, ParameterSetName = "AnalyzeDataOnly")]
    [switch]$AnalyzeDataOnly,
    [Parameter(Mandatory = $false)][switch]$SkipVersionCheck,
    [Parameter(Mandatory = $false)][switch]$SaveDebugLog,
    [Parameter(Mandatory = $false, ParameterSetName = "ScriptUpdateOnly")]
    [switch]$ScriptUpdateOnly
)

$BuildVersion = "21.09.28.1529"

$VirtualizationWarning = @"
Virtual Machine detected.  Certain settings about the host hardware cannot be detected from the virtual machine.  Verify on the VM Host that:

    - There is no more than a 1:1 Physical Core to Virtual CPU ratio (no oversubscribing)
    - If Hyper-Threading is enabled do NOT count Hyper-Threaded cores as physical cores
    - Do not oversubscribe memory or use dynamic memory allocation

Although Exchange technically supports up to a 2:1 physical core to vCPU ratio, a 1:1 ratio is strongly recommended for performance reasons.  Certain third party Hyper-Visors such as VMWare have their own guidance.

VMWare recommends a 1:1 ratio.  Their guidance can be found at https://aka.ms/HC-VMwareBP2019.
Related specifically to VMWare, if you notice you are experiencing packet loss on your VMXNET3 adapter, you may want to review the following article from VMWare:  https://aka.ms/HC-VMwareLostPackets.

For further details, please review the virtualization recommendations on Microsoft Docs here: https://aka.ms/HC-Virtualization.

"@

$Script:VerboseEnabled = $false
#this is to set the verbose information to a different color
if ($PSBoundParameters["Verbose"]) {
    #Write verbose output in cyan since we already use yellow for warnings
    $Script:VerboseEnabled = $true
    $VerboseForeground = $Host.PrivateData.VerboseForegroundColor
    $Host.PrivateData.VerboseForegroundColor = "Cyan"
}



Function Add-AnalyzedResultInformation {
    param(
        [object]$Details,
        [string]$Name,
        [object]$OutColumns,
        [string]$HtmlName,
        [object]$DisplayGroupingKey,
        [int]$DisplayCustomTabNumber = -1,
        [object]$DisplayTestingValue,
        [string]$DisplayWriteType = "Grey",
        [bool]$AddDisplayResultsLineInfo = $true,
        [bool]$AddHtmlDetailRow = $true,
        [string]$HtmlDetailsCustomValue = "",
        [bool]$AddHtmlOverviewValues = $false,
        [bool]$AddHtmlActionRow = $false,
        #[string]$ActionSettingClass = "",
        #[string]$ActionSettingValue,
        #[string]$ActionRecommendedDetailsClass = "",
        #[string]$ActionRecommendedDetailsValue,
        #[string]$ActionMoreInformationClass = "",
        #[string]$ActionMoreInformationValue,
        [HealthChecker.AnalyzedInformation]$AnalyzedInformation
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand): $name"

    if ($AddDisplayResultsLineInfo) {
        if (!($AnalyzedInformation.DisplayResults.ContainsKey($DisplayGroupingKey))) {
            Write-Verbose "Adding Display Grouping Key: $($DisplayGroupingKey.Name)"
            [System.Collections.Generic.List[HealthChecker.DisplayResultsLineInfo]]$list = New-Object System.Collections.Generic.List[HealthChecker.DisplayResultsLineInfo]
            $AnalyzedInformation.DisplayResults.Add($DisplayGroupingKey, $list)
        }

        $lineInfo = New-Object HealthChecker.DisplayResultsLineInfo

        if ($null -ne $OutColumns) {
            $lineInfo.OutColumns = $OutColumns
            $lineInfo.WriteType = "OutColumns"
            $lineInfo.TestingValue = $OutColumns
        } else {

            $lineInfo.DisplayValue = $Details
            $lineInfo.Name = $Name

            if ($DisplayCustomTabNumber -ne -1) {
                $lineInfo.TabNumber = $DisplayCustomTabNumber
            } else {
                $lineInfo.TabNumber = $DisplayGroupingKey.DefaultTabNumber
            }

            if ($null -ne $DisplayTestingValue) {
                $lineInfo.TestingValue = $DisplayTestingValue
            } else {
                $lineInfo.TestingValue = $Details
            }

            $lineInfo.WriteType = $DisplayWriteType
        }

        $AnalyzedInformation.DisplayResults[$DisplayGroupingKey].Add($lineInfo)
    }

    if ($AddHtmlDetailRow) {
        if (!($analyzedResults.HtmlServerValues.ContainsKey("ServerDetails"))) {
            [System.Collections.Generic.List[HealthChecker.HtmlServerInformationRow]]$list = New-Object System.Collections.Generic.List[HealthChecker.HtmlServerInformationRow]
            $AnalyzedInformation.HtmlServerValues.Add("ServerDetails", $list)
        }

        $detailRow = New-Object HealthChecker.HtmlServerInformationRow

        if ($displayWriteType -ne "Grey") {
            $detailRow.Class = $displayWriteType
        }

        if ([string]::IsNullOrEmpty($HtmlName)) {
            $detailRow.Name = $Name
        } else {
            $detailRow.Name = $HtmlName
        }

        if ([string]::IsNullOrEmpty($HtmlDetailsCustomValue)) {
            $detailRow.DetailValue = $Details
        } else {
            $detailRow.DetailValue = $HtmlDetailsCustomValue
        }

        $AnalyzedInformation.HtmlServerValues["ServerDetails"].Add($detailRow)
    }

    if ($AddHtmlOverviewValues) {
        if (!($analyzedResults.HtmlServerValues.ContainsKey("OverviewValues"))) {
            [System.Collections.Generic.List[HealthChecker.HtmlServerInformationRow]]$list = New-Object System.Collections.Generic.List[HealthChecker.HtmlServerInformationRow]
            $AnalyzedInformation.HtmlServerValues.Add("OverviewValues", $list)
        }

        $overviewValue = New-Object HealthChecker.HtmlServerInformationRow

        if ($displayWriteType -ne "Grey") {
            $overviewValue.Class = $displayWriteType
        }

        if ([string]::IsNullOrEmpty($HtmlName)) {
            $overviewValue.Name = $Name
        } else {
            $overviewValue.Name = $HtmlName
        }

        if ([string]::IsNullOrEmpty($HtmlDetailsCustomValue)) {
            $overviewValue.DetailValue = $Details
        } else {
            $overviewValue.DetailValue = $HtmlDetailsCustomValue
        }

        $AnalyzedInformation.HtmlServerValues["OverviewValues"].Add($overviewValue)
    }

    if ($AddHtmlActionRow) {
        #TODO
    }

    return $AnalyzedInformation
}

Function Get-DisplayResultsGroupingKey {
    param(
        [string]$Name,
        [bool]$DisplayGroupName = $true,
        [int]$DisplayOrder,
        [int]$DefaultTabNumber = 1
    )
    $obj = New-Object HealthChecker.DisplayResultsGroupingKey
    $obj.Name = $Name
    $obj.DisplayGroupName = $DisplayGroupName
    $obj.DisplayOrder = $DisplayOrder
    $obj.DefaultTabNumber = $DefaultTabNumber
    return $obj
}
Function Invoke-AnalyzerEngine {
    param(
        [HealthChecker.HealthCheckerExchangeServer]$HealthServerObject
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $analyzedResults = New-Object HealthChecker.AnalyzedInformation
    $analyzedResults.HealthCheckerExchangeServer = $HealthServerObject

    #Display Grouping Keys
    $order = 0
    $keyBeginningInfo = Get-DisplayResultsGroupingKey -Name "BeginningInfo" -DisplayGroupName $false -DisplayOrder ($order++) -DefaultTabNumber 0
    $keyExchangeInformation = Get-DisplayResultsGroupingKey -Name "Exchange Information"  -DisplayOrder ($order++)
    $keyHybridInformation = Get-DisplayResultsGroupingKey -Name "Hybrid Information" -DisplayOrder ($order++)
    $keyOSInformation = Get-DisplayResultsGroupingKey -Name "Operating System Information" -DisplayOrder ($order++)
    $keyHardwareInformation = Get-DisplayResultsGroupingKey -Name "Processor/Hardware Information" -DisplayOrder ($order++)
    $keyNICSettings = Get-DisplayResultsGroupingKey -Name "NIC Settings Per Active Adapter" -DisplayOrder ($order++) -DefaultTabNumber 2
    $keyFrequentConfigIssues = Get-DisplayResultsGroupingKey -Name "Frequent Configuration Issues" -DisplayOrder ($order++)
    $keySecuritySettings = Get-DisplayResultsGroupingKey -Name "Security Settings" -DisplayOrder ($order++)
    $keyWebApps = Get-DisplayResultsGroupingKey -Name "Exchange Web App Pools" -DisplayOrder ($order++)

    #Set short cut variables
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $osInformation = $HealthServerObject.OSInformation
    $hardwareInformation = $HealthServerObject.HardwareInformation

    if (!$Script:DisplayedScriptVersionAlready) {
        $analyzedResults = Add-AnalyzedResultInformation -Name "Exchange Health Checker Version" -Details $BuildVersion `
            -DisplayGroupingKey $keyBeginningInfo `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $analyzedResults
    }

    if ($HealthServerObject.HardwareInformation.ServerType -eq [HealthChecker.ServerType]::VMWare -or
        $HealthServerObject.HardwareInformation.ServerType -eq [HealthChecker.ServerType]::HyperV) {
        $analyzedResults = Add-AnalyzedResultInformation -Details $VirtualizationWarning -DisplayWriteType "Yellow" `
            -DisplayGroupingKey $keyBeginningInfo `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $analyzedResults
    }

    #########################
    # Exchange Information
    #########################
    Write-Verbose "Working on Exchange Information"

    $analyzedResults = Add-AnalyzedResultInformation -Name "Name" -Details ($HealthServerObject.ServerName) `
        -DisplayGroupingKey $keyExchangeInformation `
        -AddHtmlOverviewValues $true `
        -HtmlName "Server Name" `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "Generation Time" -Details $HealthServerObject.GenerationTime `
        -DisplayGroupingKey $keyExchangeInformation `
        -AddHtmlOverviewValues $true `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "Version" -Details ($exchangeInformation.BuildInformation.FriendlyName) `
        -DisplayGroupingKey $keyExchangeInformation `
        -AddHtmlOverviewValues $true `
        -HtmlName "Exchange Version" `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "Build Number" -Details ($exchangeInformation.BuildInformation.BuildNumber) `
        -DisplayGroupingKey $keyExchangeInformation `
        -AnalyzedInformation $analyzedResults

    if ($exchangeInformation.BuildInformation.SupportedBuild -eq $false) {
        $daysOld = ($date - ([System.Convert]::ToDateTime([DateTime]$exchangeInformation.BuildInformation.ReleaseDate, [System.Globalization.DateTimeFormatInfo]::InvariantInfo))).Days

        $analyzedResults = Add-AnalyzedResultInformation -Name "Error" -Details ("Out of date Cumulative Update. Please upgrade to one of the two most recently released Cumulative Updates. Currently running on a build that is {0} days old." -f $daysOld) `
            -DisplayGroupingKey $keyExchangeInformation `
            -DisplayWriteType "Red" `
            -DisplayCustomTabNumber 2 `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $analyzedResults
    }

    if ($null -ne $exchangeInformation.BuildInformation.LocalBuildNumber) {
        $local = $exchangeInformation.BuildInformation.LocalBuildNumber
        $remote = $exchangeInformation.BuildInformation.BuildNumber

        if ($local.Substring(0, $local.LastIndexOf(".")) -ne $remote.Substring(0, $remote.LastIndexOf("."))) {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Warning" -Details ("Running commands from a different version box can cause issues. Local Tools Server Version: {0}" -f $local) `
                -DisplayGroupingKey $keyExchangeInformation `
                -DisplayWriteType "Yellow" `
                -DisplayCustomTabNumber 2 `
                -AddHtmlDetailRow $false `
                -AnalyzedInformation $analyzedResults
        }
    }

    if ($null -ne $exchangeInformation.BuildInformation.KBsInstalled) {
        $analyzedResults = Add-AnalyzedResultInformation -Details ("Exchange IU or Security Hotfix Detected.") `
            -DisplayGroupingKey $keyExchangeInformation `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $analyzedResults

        foreach ($kb in $exchangeInformation.BuildInformation.KBsInstalled) {
            $analyzedResults = Add-AnalyzedResultInformation -Details $kb `
                -DisplayGroupingKey $keyExchangeInformation `
                -DisplayCustomTabNumber 2 `
                -AddHtmlDetailRow $false `
                -AnalyzedInformation $analyzedResults
        }
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "Server Role" -Details ($exchangeInformation.BuildInformation.ServerRole) `
        -DisplayGroupingKey $keyExchangeInformation `
        -AddHtmlOverviewValues $true `
        -AnalyzedInformation $analyzedResults

    if ($exchangeInformation.BuildInformation.ServerRole -le [HealthChecker.ExchangeServerRole]::Mailbox) {
        $dagName = [System.Convert]::ToString($exchangeInformation.GetMailboxServer.DatabaseAvailabilityGroup)
        if ([System.String]::IsNullOrWhiteSpace($dagName)) {
            $dagName = "Standalone Server"
        }
        $analyzedResults = Add-AnalyzedResultInformation -Name "DAG Name" -Details $dagName `
            -DisplayGroupingKey $keyExchangeInformation `
            -AnalyzedInformation $analyzedResults
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "AD Site" -Details ([System.Convert]::ToString(($exchangeInformation.GetExchangeServer.Site)).Split("/")[-1]) `
        -DisplayGroupingKey $keyExchangeInformation `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "MAPI/HTTP Enabled" -Details ($exchangeInformation.MapiHttpEnabled) `
        -DisplayGroupingKey $keyExchangeInformation `
        -AnalyzedInformation $analyzedResults

    if ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013 -and
        $exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge -and
        $exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Mailbox) {

        if ($null -ne $exchangeInformation.ApplicationPools -and
            $exchangeInformation.ApplicationPools.Count -gt 0) {
            $mapiFEAppPool = $exchangeInformation.ApplicationPools["MSExchangeMapiFrontEndAppPool"]
            [bool]$enabled = $mapiFEAppPool.GCServerEnabled
            [bool]$unknown = $mapiFEAppPool.GCUnknown
            $warning = [string]::Empty
            $displayWriteType = "Green"
            $displayValue = "Server"

            if ($hardwareInformation.TotalMemory -ge 21474836480 -and
                $enabled -eq $false) {
                $displayWriteType = "Red"
                $displayValue = "Workstation --- Error"
                $warning = "To Fix this issue go into the file MSExchangeMapiFrontEndAppPool_CLRConfig.config in the Exchange Bin directory and change the GCServer to true and recycle the MAPI Front End App Pool"
            } elseif ($unknown) {
                $displayValue = "Unknown --- Warning"
                $displayWriteType = "Yellow"
            } elseif (!($enabled)) {
                $displayWriteType = "Yellow"
                $displayValue = "Workstation --- Warning"
                $warning = "You could be seeing some GC issues within the Mapi Front End App Pool. However, you don't have enough memory installed on the system to recommend switching the GC mode by default without consulting a support professional."
            }

            $analyzedResults = Add-AnalyzedResultInformation -Name "MAPI Front End App Pool GC Mode" -Details $displayValue `
                -DisplayGroupingKey $keyExchangeInformation `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType $displayWriteType `
                -AnalyzedInformation $analyzedResults
        } else {
            $warning = "Unable to determine MAPI Front End App Pool GC Mode status. This may be a temporary issue. You should try to re-run the script"
        }

        if ($warning -ne [string]::Empty) {
            $analyzedResults = Add-AnalyzedResultInformation -Details $warning `
                -DisplayGroupingKey $keyExchangeInformation `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType "Yellow" `
                -AddHtmlDetailRow $false `
                -AnalyzedInformation $analyzedResults
        }
    }

    if (-not ([string]::IsNullOrWhiteSpace($exchangeInformation.GetWebServicesVirtualDirectory.InternalNLBBypassUrl))) {
        $analyzedResults = Add-AnalyzedResultInformation -Name "EWS Internal Bypass URL Set" -Details ("$($exchangeInformation.GetWebServicesVirtualDirectory.InternalNLBBypassUrl) - Can cause issues after KB 5001779") `
            -DisplayGroupingKey $keyExchangeInformation `
            -DisplayWriteType "Red" `
            -AnalyzedInformation $analyzedResults
    }

    #########################
    # Hybrid Information
    #########################
    Write-Verbose "Working on Hybrid Configuration Information"
    if ($exchangeInformation.BuildInformation.MajorVersion -ge [HealthChecker.ExchangeMajorVersion]::Exchange2013 -and
        $null -ne $exchangeInformation.GetHybridConfiguration) {

        $analyzedResults = Add-AnalyzedResultInformation -Name "Organization Hybrid enabled" -Details "True" `
            -DisplayGroupingKey $keyHybridInformation `
            -AnalyzedInformation $analyzedResults

        if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.OnPremisesSmartHost))) {
            $onPremSmartHostDomain = ($exchangeInformation.GetHybridConfiguration.OnPremisesSmartHost).ToString()
            $onPremSmartHostWriteType = "Grey"
        } else {
            $onPremSmartHostDomain = "No on-premises smart host domain configured for hybrid use"
            $onPremSmartHostWriteType = "Yellow"
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "On-Premises Smart Host Domain" -Details $onPremSmartHostDomain `
            -DisplayGroupingKey $keyHybridInformation `
            -DisplayWriteType $onPremSmartHostWriteType `
            -AnalyzedInformation $analyzedResults

        if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.Domains))) {
            $domainsConfiguredForHybrid = $exchangeInformation.GetHybridConfiguration.Domains
            $domainsConfiguredForHybridWriteType = "Grey"
        } else {
            $domainsConfiguredForHybridWriteType = "Yellow"
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "Domain(s) configured for Hybrid use" `
            -DisplayGroupingKey $keyHybridInformation `
            -DisplayWriteType $domainsConfiguredForHybridWriteType `
            -AnalyzedInformation $analyzedResults

        if ($domainsConfiguredForHybrid.Count -ge 1) {
            foreach ($domain in $domainsConfiguredForHybrid) {
                $analyzedResults = Add-AnalyzedResultInformation -Details $domain `
                    -DisplayGroupingKey $keyHybridInformation `
                    -DisplayWriteType $domainsConfiguredForHybridWriteType `
                    -DisplayCustomTabNumber 2 `
                    -AnalyzedInformation $analyzedResults
            }
        } else {
            $analyzedResults = Add-AnalyzedResultInformation -Details "No domain configured for Hybrid use" `
                -DisplayGroupingKey $keyHybridInformation `
                -DisplayWriteType $domainsConfiguredForHybridWriteType `
                -DisplayCustomTabNumber 2 `
                -AnalyzedInformation $analyzedResults
        }

        if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.EdgeTransportServers))) {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Edge Transport Server(s)" `
                -DisplayGroupingKey $keyHybridInformation `
                -AnalyzedInformation $analyzedResults

            foreach ($edgeServer in $exchangeInformation.GetHybridConfiguration.EdgeTransportServers) {
                $analyzedResults = Add-AnalyzedResultInformation -Details $edgeServer `
                    -DisplayGroupingKey $keyHybridInformation `
                    -DisplayCustomTabNumber 2  `
                    -AnalyzedInformation $analyzedResults
            }

            if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.ReceivingTransportServers)) -or
                (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.SendingTransportServers)))) {
                $analyzedResults = Add-AnalyzedResultInformation -Details "When configuring the EdgeTransportServers parameter, you must configure the ReceivingTransportServers and SendingTransportServers parameter values to null" `
                    -DisplayGroupingKey $keyHybridInformation `
                    -DisplayWriteType "Yellow" `
                    -DisplayCustomTabNumber 2 `
                    -AnalyzedInformation $analyzedResults
            }
        } else {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Receiving Transport Server(s)" `
                -DisplayGroupingKey $keyHybridInformation `
                -AnalyzedInformation $analyzedResults

            if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.ReceivingTransportServers))) {
                foreach ($receivingTransportSrv in $exchangeInformation.GetHybridConfiguration.ReceivingTransportServers) {
                    $analyzedResults = Add-AnalyzedResultInformation -Details $receivingTransportSrv `
                        -DisplayGroupingKey $keyHybridInformation `
                        -DisplayCustomTabNumber 2 `
                        -AnalyzedInformation $analyzedResults
                }
            } else {
                $analyzedResults = Add-AnalyzedResultInformation -Details "No Receiving Transport Server configured for Hybrid use" `
                    -DisplayGroupingKey $keyHybridInformation `
                    -DisplayCustomTabNumber 2 `
                    -DisplayWriteType "Yellow" `
                    -AnalyzedInformation $analyzedResults
            }

            $analyzedResults = Add-AnalyzedResultInformation -Name "Sending Transport Server(s)" `
                -DisplayGroupingKey $keyHybridInformation `
                -AnalyzedInformation $analyzedResults

            if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.SendingTransportServers))) {
                foreach ($sendingTransportSrv in $exchangeInformation.GetHybridConfiguration.SendingTransportServers) {
                    $analyzedResults = Add-AnalyzedResultInformation -Details $sendingTransportSrv `
                        -DisplayGroupingKey $keyHybridInformation `
                        -DisplayCustomTabNumber 2 `
                        -AnalyzedInformation $analyzedResults
                }
            } else {
                $analyzedResults = Add-AnalyzedResultInformation -Details "No Sending Transport Server configured for Hybrid use" `
                    -DisplayGroupingKey $keyHybridInformation `
                    -DisplayCustomTabNumber 2 `
                    -DisplayWriteType "Yellow" `
                    -AnalyzedInformation $analyzedResults
            }
        }

        if ($exchangeInformation.GetHybridConfiguration.ServiceInstance -eq 1) {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Service Instance" -Details "Office 365 operated by 21Vianet" `
                -DisplayGroupingKey $keyHybridInformation `
                -AnalyzedInformation $analyzedResults
        } elseif ($exchangeInformation.GetHybridConfiguration.ServiceInstance -ne 0) {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Service Instance" -Details ($exchangeInformation.GetHybridConfiguration.ServiceInstance) `
                -DisplayGroupingKey $keyHybridInformation `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults

            $analyzedResults = Add-AnalyzedResultInformation -Details "You are using an invalid value. Please set this value to 0 (null) or re-run HCW" `
                -DisplayGroupingKey $keyHybridInformation `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults
        }

        if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.TlsCertificateName))) {
            $analyzedResults = Add-AnalyzedResultInformation -Name "TLS Certificate Name" -Details ($exchangeInformation.GetHybridConfiguration.TlsCertificateName).ToString() `
                -DisplayGroupingKey $keyHybridInformation `
                -AnalyzedInformation $analyzedResults
        } else {
            $analyzedResults = Add-AnalyzedResultInformation -Name "TLS Certificate Name" -Details "No valid certificate found" `
                -DisplayGroupingKey $keyHybridInformation `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "Feature(s) enabled for Hybrid use" `
            -DisplayGroupingKey $keyHybridInformation `
            -AnalyzedInformation $analyzedResults

        if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.Features))) {
            foreach ($feature in $exchangeInformation.GetHybridConfiguration.Features) {
                $analyzedResults = Add-AnalyzedResultInformation -Details $feature `
                    -DisplayGroupingKey $keyHybridInformation `
                    -DisplayCustomTabNumber 2  `
                    -AnalyzedInformation $analyzedResults
            }
        } else {
            $analyzedResults = Add-AnalyzedResultInformation -Details "No feature(s) enabled for Hybrid use" `
                -DisplayGroupingKey $keyHybridInformation `
                -DisplayCustomTabNumber 2  `
                -AnalyzedInformation $analyzedResults
        }
    }

    ##############################
    # Exchange Test Services
    ##############################
    Write-Verbose "Working on results from Test-ServiceHealth"
    $servicesNotRunning = $exchangeInformation.ExchangeServicesNotRunning
    if ($null -ne $servicesNotRunning) {
        $analyzedResults = Add-AnalyzedResultInformation -Name "Services Not Running" `
            -DisplayGroupingKey $keyExchangeInformation `
            -AnalyzedInformation $analyzedResults

        foreach ($stoppedService in $servicesNotRunning) {
            $analyzedResults = Add-AnalyzedResultInformation -Details $stoppedService `
                -DisplayGroupingKey $keyExchangeInformation `
                -DisplayCustomTabNumber 2  `
                -DisplayWriteType "Yellow" `
                -AnalyzedInformation $analyzedResults
        }
    }

    ##############################
    # Exchange Server Maintenance
    ##############################
    Write-Verbose "Working on Exchange Server Maintenance"
    $serverMaintenance = $exchangeInformation.ServerMaintenance

    if (($serverMaintenance.InactiveComponents).Count -eq 0 -and
        ($null -eq $serverMaintenance.GetClusterNode -or
        $serverMaintenance.GetClusterNode.State -eq "Up") -and
        ($null -eq $serverMaintenance.GetMailboxServer -or
            ($serverMaintenance.GetMailboxServer.DatabaseCopyActivationDisabledAndMoveNow -eq $false -and
        $serverMaintenance.GetMailboxServer.DatabaseCopyAutoActivationPolicy -eq "Unrestricted"))) {
        $analyzedResults = Add-AnalyzedResultInformation -Name "Exchange Server Maintenance" -Details "Server is not in Maintenance Mode" `
            -DisplayGroupingKey $keyExchangeInformation `
            -DisplayWriteType "Green" `
            -AnalyzedInformation $analyzedResults
    } else {
        $analyzedResults = Add-AnalyzedResultInformation -Details "Exchange Server Maintenance" `
            -DisplayGroupingKey $keyExchangeInformation `
            -AnalyzedInformation $analyzedResults

        if (($serverMaintenance.InactiveComponents).Count -ne 0) {
            foreach ($inactiveComponent in $serverMaintenance.InactiveComponents) {
                $analyzedResults = Add-AnalyzedResultInformation -Name "Component" -Details $inactiveComponent `
                    -DisplayGroupingKey $keyExchangeInformation `
                    -DisplayCustomTabNumber 2  `
                    -DisplayWriteType "Red" `
                    -AnalyzedInformation $analyzedResults
            }

            $analyzedResults = Add-AnalyzedResultInformation -Details "For more information: https://aka.ms/HC-ServerComponentState" `
                -DisplayGroupingKey $keyExchangeInformation `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType "Yellow" `
                -AnalyzedInformation $analyzedResults
        }

        if ($serverMaintenance.GetMailboxServer.DatabaseCopyActivationDisabledAndMoveNow -or
            $serverMaintenance.GetMailboxServer.DatabaseCopyAutoActivationPolicy -eq "Blocked") {
            $displayValue = "`r`n`t`tDatabaseCopyActivationDisabledAndMoveNow: {0} --- should be 'false'`r`n`t`tDatabaseCopyAutoActivationPolicy: {1} --- should be 'unrestricted'" -f `
                $serverMaintenance.GetMailboxServer.DatabaseCopyActivationDisabledAndMoveNow,
            $serverMaintenance.GetMailboxServer.DatabaseCopyAutoActivationPolicy

            $analyzedResults = Add-AnalyzedResultInformation -Name "Database Copy Maintenance" -Details $displayValue `
                -DisplayGroupingKey $keyExchangeInformation `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults
        }

        if ($null -ne $serverMaintenance.GetClusterNode -and
            $serverMaintenance.GetClusterNode.State -ne "Up") {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Cluster Node" -Details ("'{0}' --- should be 'Up'" -f $serverMaintenance.GetClusterNode.State) `
                -DisplayGroupingKey $keyExchangeInformation `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults
        }
    }

    #########################
    # Operating System
    #########################
    Write-Verbose "Working on Operating System"

    $analyzedResults = Add-AnalyzedResultInformation -Name "Version" -Details ($osInformation.BuildInformation.FriendlyName) `
        -DisplayGroupingKey $keyOSInformation `
        -AddHtmlOverviewValues $true `
        -HtmlName "OS Version" `
        -AnalyzedInformation $analyzedResults

    $upTime = "{0} day(s) {1} hour(s) {2} minute(s) {3} second(s)" -f $osInformation.ServerBootUp.Days,
    $osInformation.ServerBootUp.Hours,
    $osInformation.ServerBootUp.Minutes,
    $osInformation.ServerBootUp.Seconds

    $analyzedResults = Add-AnalyzedResultInformation -Name "System Up Time" -Details $upTime `
        -DisplayGroupingKey $keyOSInformation `
        -DisplayTestingValue ($osInformation.ServerBootUp) `
        -AddHtmlDetailRow $false `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "Time Zone" -Details ($osInformation.TimeZone.CurrentTimeZone) `
        -DisplayGroupingKey $keyOSInformation `
        -AddHtmlOverviewValues $true `
        -AnalyzedInformation $analyzedResults

    $writeValue = $false
    $warning = @("Windows can not properly detect any DST rule changes in your time zone. Set 'Adjust for daylight saving time automatically to on'")

    if ($osInformation.TimeZone.DstIssueDetected) {
        $writeType = "Red"
    } elseif ($osInformation.TimeZone.DynamicDaylightTimeDisabled -ne 0) {
        $writeType = "Yellow"
    } else {
        $warning = [string]::Empty
        $writeValue = $true
        $writeType = "Grey"
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "Dynamic Daylight Time Enabled" -Details $writeValue `
        -DisplayGroupingKey $keyOSInformation `
        -DisplayWriteType $writeType `
        -AnalyzedInformation $analyzedResults

    if ($warning -ne [string]::Empty) {
        $analyzedResults = Add-AnalyzedResultInformation -Details $warning `
            -DisplayGroupingKey $keyOSInformation `
            -DisplayWriteType "Yellow" `
            -DisplayCustomTabNumber 2 `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $analyzedResults
    }

    if ([string]::IsNullOrEmpty($osInformation.TimeZone.TimeZoneKeyName)) {
        $analyzedResults = Add-AnalyzedResultInformation -Name "Time Zone Key Name" -Details "Empty --- Warning Need to switch your current time zone to a different value, then switch it back to have this value populated again." `
            -DisplayGroupingKey $keyOSInformation `
            -DisplayWriteType "Yellow" `
            -AnalyzedInformation $analyzedResults
    }

    if ($exchangeInformation.NETFramework.OnRecommendedVersion) {
        $analyzedResults = Add-AnalyzedResultInformation -Name ".NET Framework" -Details ($osInformation.NETFramework.FriendlyName) `
            -DisplayGroupingKey $keyOSInformation `
            -DisplayWriteType "Green" `
            -AddHtmlOverviewValues $true `
            -AnalyzedInformation $analyzedResults
    } else {
        $testObject = New-Object PSCustomObject
        $testObject | Add-Member -MemberType NoteProperty -Name "CurrentValue" -Value ($osInformation.NETFramework.FriendlyName)
        $testObject | Add-Member -MemberType NoteProperty -Name "MaxSupportedVersion" -Value ($exchangeInformation.NETFramework.MaxSupportedVersion)
        $displayFriendly = Get-NETFrameworkVersion -NetVersionKey $exchangeInformation.NETFramework.MaxSupportedVersion
        $displayValue = "{0} - Warning Recommended .NET Version is {1}" -f $osInformation.NETFramework.FriendlyName, $displayFriendly.FriendlyName
        $analyzedResults = Add-AnalyzedResultInformation -Name ".NET Framework" -Details $displayValue `
            -DisplayGroupingKey $keyOSInformation `
            -DisplayWriteType "Yellow" `
            -DisplayTestingValue $testObject `
            -HtmlDetailsCustomValue ($osInformation.NETFramework.FriendlyName) `
            -AddHtmlOverviewValues $true `
            -AnalyzedInformation $analyzedResults
    }

    $displayValue = [string]::Empty
    $displayWriteType = "Yellow"
    $totalPhysicalMemory = $hardwareInformation.TotalMemory
    $maxPageSize = $osInformation.PageFile.MaxPageSize
    Write-Verbose "Total Memory: $totalPhysicalMemory"
    Write-Verbose "Page File: $maxPageSize"
    $testingValue = New-Object PSCustomObject
    $testingValue | Add-Member -MemberType NoteProperty -Name "TotalPhysicalMemory" -Value $totalPhysicalMemory
    $testingValue | Add-Member -MemberType NoteProperty -Name "MaxPageSize" -Value $maxPageSize
    $testingValue | Add-Member -MemberType NoteProperty -Name "MultiPageFile" -Value ($osInformation.PageFile.PageFile.Count -gt 1)
    $testingValue | Add-Member -MemberType NoteProperty -Name "RecommendedPageFile" -Value 0
    if ($maxPageSize -eq 0) {
        $displayValue = "Error: System is set to automatically manage the pagefile size."
        $displayWriteType = "Red"
    } elseif ($osInformation.PageFile.PageFile.Count -gt 1) {
        $displayValue = "Multiple page files detected. `r`n`t`tError: This has been know to cause performance issues please address this."
        $displayWriteType = "Red"
    } elseif ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) {
        $testingValue.RecommendedPageFile = ($recommendedPageFileSize = [Math]::Round(($totalPhysicalMemory / 1MB) / 4))
        Write-Verbose "Recommended Page File Size: $recommendedPageFileSize"
        if ($recommendedPageFileSize -ne $maxPageSize) {
            $displayValue = "{0}MB `r`n`t`tWarning: Page File is not set to 25% of the Total System Memory which is {1}MB. Recommended is {2}MB" -f $maxPageSize, ([Math]::Round($totalPhysicalMemory / 1MB)), $recommendedPageFileSize
        } else {
            $displayValue = "{0}MB" -f $recommendedPageFileSize
            $displayWriteType = "Grey"
        }
    } elseif ($totalPhysicalMemory -ge 34359738368) {
        #32GB = 1024 * 1024 * 1024 * 32 = 34,359,738,368
        if ($maxPageSize -eq 32778) {
            $displayValue = "{0}MB" -f $maxPageSize
            $displayWriteType = "Grey"
        } else {
            $displayValue = "{0}MB `r`n`t`tWarning: Pagefile should be capped at 32778MB for 32GB plus 10MB - Article: https://aka.ms/HC-SystemRequirements2016#hardware-requirements-for-exchange-2016" -f $maxPageSize
        }
    } else {
        $testingValue.RecommendedPageFile = ($recommendedPageFileSize = [Math]::Round(($totalPhysicalMemory / 1MB) + 10))

        if ($recommendedPageFileSize -ne $maxPageSize) {
            $displayValue = "{0}MB `r`n`t`tWarning: Page File is not set to Total System Memory plus 10MB which should be {1}MB" -f $maxPageSize, $recommendedPageFileSize
        } else {
            $displayValue = "{0}MB" -f $maxPageSize
            $displayWriteType = "Grey"
        }
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "Page File Size" -Details $displayValue `
        -DisplayGroupingKey $keyOSInformation `
        -DisplayWriteType $displayWriteType `
        -DisplayTestingValue $testingValue `
        -AnalyzedInformation $analyzedResults

    if ($osInformation.PowerPlan.HighPerformanceSet) {
        $analyzedResults = Add-AnalyzedResultInformation -Name "Power Plan" -Details ($osInformation.PowerPlan.PowerPlanSetting) `
            -DisplayGroupingKey $keyOSInformation `
            -DisplayWriteType "Green" `
            -AnalyzedInformation $analyzedResults
    } else {
        $displayValue = "{0} --- Error" -f $osInformation.PowerPlan.PowerPlanSetting
        $analyzedResults = Add-AnalyzedResultInformation -Name "Power Plan" -Details $displayValue `
            -DisplayGroupingKey $keyOSInformation `
            -DisplayWriteType "Red" `
            -AnalyzedInformation $analyzedResults
    }

    if ($osInformation.NetworkInformation.HttpProxy -eq "<None>") {
        $analyzedResults = Add-AnalyzedResultInformation -Name "Http Proxy Setting" -Details ($osInformation.NetworkInformation.HttpProxy) `
            -DisplayGroupingKey $keyOSInformation `
            -HtmlDetailsCustomValue "None" `
            -AnalyzedInformation $analyzedResults
    } else {
        $displayValue = "{0} --- Warning this can cause client connectivity issues." -f $osInformation.NetworkInformation.HttpProxy
        $analyzedResults = Add-AnalyzedResultInformation -Name "Http Proxy Setting" -Details $displayValue `
            -DisplayGroupingKey $keyOSInformation `
            -DisplayWriteType "Yellow" `
            -DisplayTestingValue ($osInformation.NetworkInformation.HttpProxy) `
            -AnalyzedInformation $analyzedResults
    }

    $displayWriteType2012 = "Yellow"
    $displayWriteType2013 = "Yellow"
    $displayValue2012 = "Unknown"
    $displayValue2013 = "Unknown"

    if ($null -ne $osInformation.VcRedistributable) {

        if (Test-VisualCRedistributableUpToDate -Year 2012 -Installed $osInformation.VcRedistributable) {
            $displayWriteType2012 = "Green"
            $displayValue2012 = "$((Get-VisualCRedistributableInfo 2012).VersionNumber) Version is current"
        } elseif (Test-VisualCRedistributableInstalled -Year 2012 -Installed $osInformation.VcRedistributable) {
            $displayValue2012 = "Redistributable is outdated"
        }

        if (Test-VisualCRedistributableUpToDate -Year 2013 -Installed $osInformation.VcRedistributable) {
            $displayWriteType2013 = "Green"
            $displayValue2013 = "$((Get-VisualCRedistributableInfo 2013).VersionNumber) Version is current"
        } elseif (Test-VisualCRedistributableInstalled -Year 2013 -Installed $osInformation.VcRedistributable) {
            $displayValue2013 = "Redistributable is outdated"
        }
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "Visual C++ 2012" -Details $displayValue2012 `
        -DisplayGroupingKey $keyOSInformation `
        -DisplayWriteType $displayWriteType2012 `
        -AnalyzedInformation $analyzedResults

    if ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {
        $analyzedResults = Add-AnalyzedResultInformation -Name "Visual C++ 2013" -Details $displayValue2013 `
            -DisplayGroupingKey $keyOSInformation `
            -DisplayWriteType $displayWriteType2013 `
            -AnalyzedInformation $analyzedResults
    }

    if (($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge -and
            ($displayWriteType2012 -eq "Yellow" -or
            $displayWriteType2013 -eq "Yellow")) -or
        $displayWriteType2012 -eq "Yellow") {

        $analyzedResults = Add-AnalyzedResultInformation -Details "Note: For more information about the latest C++ Redistributeable please visit: https://aka.ms/HC-LatestVC`r`n`t`tThis is not a requirement to upgrade, only a notification to bring to your attention." `
            -DisplayGroupingKey $keyOSInformation `
            -DisplayCustomTabNumber 2 `
            -DisplayWriteType "Yellow" `
            -AnalyzedInformation $analyzedResults
    }

    $displayValue = "False"
    $writeType = "Grey"

    if ($osInformation.ServerPendingReboot.PendingReboot) {
        $displayValue = "True --- Warning a reboot is pending and can cause issues on the server."
        $writeType = "Yellow"
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "Server Pending Reboot" -Details $displayValue `
        -DisplayGroupingKey $keyOSInformation `
        -DisplayWriteType $writeType `
        -DisplayTestingValue ($osInformation.ServerPendingReboot.PendingReboot) `
        -AnalyzedInformation $analyzedResults

    ################################
    # Processor/Hardware Information
    ################################
    Write-Verbose "Working on Processor/Hardware Information"

    $analyzedResults = Add-AnalyzedResultInformation -Name "Type" -Details ($hardwareInformation.ServerType) `
        -DisplayGroupingKey $keyHardwareInformation `
        -AddHtmlOverviewValues $true `
        -Htmlname "Hardware Type" `
        -AnalyzedInformation $analyzedResults

    if ($hardwareInformation.ServerType -eq [HealthChecker.ServerType]::Physical -or
        $hardwareInformation.ServerType -eq [HealthChecker.ServerType]::AmazonEC2) {
        $analyzedResults = Add-AnalyzedResultInformation -Name "Manufacturer" -Details ($hardwareInformation.Manufacturer) `
            -DisplayGroupingKey $keyHardwareInformation `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Name "Model" -Details ($hardwareInformation.Model) `
            -DisplayGroupingKey $keyHardwareInformation `
            -AnalyzedInformation $analyzedResults
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "Processor" -Details ($hardwareInformation.Processor.Name) `
        -DisplayGroupingKey $keyHardwareInformation `
        -AnalyzedInformation $analyzedResults

    $value = $hardwareInformation.Processor.NumberOfProcessors
    $processorName = "Number of Processors"

    if ($hardwareInformation.ServerType -ne [HealthChecker.ServerType]::Physical) {
        $analyzedResults = Add-AnalyzedResultInformation -Name $processorName -Details $value `
            -DisplayGroupingKey $keyHardwareInformation `
            -AnalyzedInformation $analyzedResults

        <# Comment out for now. Not sure if we have a lot of value here as i believe this changed in newer vmware hosts versions.
        if ($hardwareInformation.ServerType -eq [HealthChecker.ServerType]::VMWare) {
            $analyzedResults = Add-AnalyzedResultInformation -Details "Note: Please make sure you are following VMware's performance recommendation to get the most out of your guest machine. VMware blog 'Does corespersocket Affect Performance?' https://blogs.vmware.com/vsphere/2013/10/does-corespersocket-affect-performance.html" `
                -DisplayGroupingKey $keyHardwareInformation `
                -DisplayCustomTabNumber 2 `
                -AnalyzedInformation $analyzedResults
        }
    #>
    } elseif ($value -gt 2) {
        $analyzedResults = Add-AnalyzedResultInformation -Name $processorName -Details ("{0} - Error: Recommended to only have 2 Processors" -f $value) `
            -DisplayGroupingKey $keyHardwareInformation `
            -DisplayWriteType "Red" `
            -DisplayTestingValue $value `
            -HtmlDetailsCustomValue $value `
            -AnalyzedInformation $analyzedResults
    } else {
        $analyzedResults = Add-AnalyzedResultInformation -Name $processorName -Details $value `
            -DisplayGroupingKey $keyHardwareInformation `
            -DisplayWriteType "Green" `
            -AnalyzedInformation $analyzedResults
    }

    $physicalValue = $hardwareInformation.Processor.NumberOfPhysicalCores
    $logicalValue = $hardwareInformation.Processor.NumberOfLogicalCores

    $displayWriteType = "Green"

    if (($logicalValue -gt 24 -and
            $exchangeInformation.BuildInformation.MajorVersion -lt [HealthChecker.ExchangeMajorVersion]::Exchange2019) -or
        $logicalValue -gt 48) {
        $displayWriteType = "Yellow"
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "Number of Physical Cores" -Details $physicalValue `
        -DisplayGroupingKey $keyHardwareInformation `
        -DisplayWriteType $displayWriteType `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "Number of Logical Cores" -Details $logicalValue `
        -DisplayGroupingKey $keyHardwareInformation `
        -DisplayWriteType $displayWriteType `
        -AddHtmlOverviewValues $true `
        -AnalyzedInformation $analyzedResults

    $displayValue = "Disabled"
    $displayWriteType = "Green"
    $displayTestingValue = $false
    $additionalDisplayValue = [string]::Empty
    $additionalWriteType = "Red"

    if ($logicalValue -gt $physicalValue) {

        if ($hardwareInformation.ServerType -ne [HealthChecker.ServerType]::HyperV) {
            $displayValue = "Enabled --- Error: Having Hyper-Threading enabled goes against best practices and can cause performance issues. Please disable as soon as possible."
            $displayTestingValue = $true
            $displayWriteType = "Red"
        } else {
            $displayValue = "Enabled --- Not Applicable"
            $displayTestingValue = $true
            $displayWriteType = "Grey"
        }

        if ($hardwareInformation.ServerType -eq [HealthChecker.ServerType]::AmazonEC2) {
            $additionalDisplayValue = "Error: For high-performance computing (HPC) application, like Exchange, Amazon recommends that you have Hyper-Threading Technology disabled in their service. More information: https://aka.ms/HC-EC2HyperThreading"
        }

        if ($hardwareInformation.Processor.Name.StartsWith("AMD")) {
            $additionalDisplayValue = "This script may incorrectly report that Hyper-Threading is enabled on certain AMD processors. Check with the manufacturer to see if your model supports SMT."
            $additionalWriteType = "Yellow"
        }
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "Hyper-Threading" -Details $displayValue `
        -DisplayGroupingKey $keyHardwareInformation `
        -DisplayWriteType $displayWriteType `
        -DisplayTestingValue $displayTestingValue `
        -AnalyzedInformation $analyzedResults

    if (!([string]::IsNullOrEmpty($additionalDisplayValue))) {
        $analyzedResults = Add-AnalyzedResultInformation -Details $additionalDisplayValue `
            -DisplayGroupingKey $keyHardwareInformation `
            -DisplayWriteType $additionalWriteType `
            -DisplayCustomTabNumber 2 `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $analyzedResults
    }

    #NUMA BIOS CHECK - AKA check to see if we can properly see all of our cores on the box
    $displayWriteType = "Yellow"
    $testingValue = "Unknown"
    $displayValue = [string]::Empty

    if ($hardwareInformation.Model.Contains("ProLiant")) {
        $name = "NUMA Group Size Optimization"

        if ($hardwareInformation.Processor.EnvironmentProcessorCount -eq -1) {
            $displayValue = "Unknown `r`n`t`tWarning: If this is set to Clustered, this can cause multiple types of issues on the server"
        } elseif ($hardwareInformation.Processor.EnvironmentProcessorCount -ne $logicalValue) {
            $displayValue = "Clustered `r`n`t`tError: This setting should be set to Flat. By having this set to Clustered, we will see multiple different types of issues."
            $testingValue = "Clustered"
            $displayWriteType = "Red"
        } else {
            $displayValue = "Flat"
            $testingValue = "Flat"
            $displayWriteType = "Green"
        }
    } else {
        $name = "All Processor Cores Visible"

        if ($hardwareInformation.Processor.EnvironmentProcessorCount -eq -1) {
            $displayValue = "Unknown `r`n`t`tWarning: If we aren't able to see all processor cores from Exchange, we could see performance related issues."
        } elseif ($hardwareInformation.Processor.EnvironmentProcessorCount -ne $logicalValue) {
            $displayValue = "Failed `r`n`t`tError: Not all Processor Cores are visible to Exchange and this will cause a performance impact"
            $displayWriteType = "Red"
            $testingValue = "Failed"
        } else {
            $displayWriteType = "Green"
            $displayValue = "Passed"
            $testingValue = "Passed"
        }
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name $name -Details $displayValue `
        -DisplayGroupingKey $keyHardwareInformation `
        -DisplayWriteType $displayWriteType `
        -DisplayTestingValue $testingValue `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "Max Processor Speed" -Details ($hardwareInformation.Processor.MaxMegacyclesPerCore) `
        -DisplayGroupingKey $keyHardwareInformation `
        -AnalyzedInformation $analyzedResults

    if ($hardwareInformation.Processor.ProcessorIsThrottled) {
        $currentSpeed = $hardwareInformation.Processor.CurrentMegacyclesPerCore
        $analyzedResults = Add-AnalyzedResultInformation -Name "Current Processor Speed" -Details ("{0} --- Error: Processor appears to be throttled." -f $currentSpeed) `
            -DisplayGroupingKey $keyHardwareInformation `
            -DisplayWriteType "Red" `
            -DisplayTestingValue $currentSpeed `
            -AnalyzedInformation $analyzedResults

        $displayValue = "Error: Power Plan is NOT set to `"High Performance`". This change doesn't require a reboot and takes affect right away. Re-run script after doing so"

        if ($osInformation.PowerPlan.HighPerformanceSet) {
            $displayValue = "Error: Power Plan is set to `"High Performance`", so it is likely that we are throttling in the BIOS of the computer settings."
        }

        $analyzedResults = Add-AnalyzedResultInformation -Details $displayValue `
            -DisplayGroupingKey $keyHardwareInformation `
            -DisplayWriteType "Red" `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $analyzedResults
    }

    $totalPhysicalMemory = [System.Math]::Round($hardwareInformation.TotalMemory / 1024 / 1024 / 1024)
    $displayWriteType = "Yellow"
    $displayDetails = [string]::Empty

    if ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) {

        if ($totalPhysicalMemory -gt 256) {
            $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to be scaled at or below 256 GB of Memory" -f $totalPhysicalMemory
        } elseif ($totalPhysicalMemory -lt 64 -and
            $exchangeInformation.BuildInformation.ServerRole -eq [HealthChecker.ExchangeServerRole]::Edge) {
            $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to have a minimum of 64GB of RAM installed on the machine." -f $totalPhysicalMemory
        } elseif ($totalPhysicalMemory -lt 128) {
            $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to have a minimum of 128GB of RAM installed on the machine." -f $totalPhysicalMemory
        } else {
            $displayDetails = "{0} GB" -f $totalPhysicalMemory
            $displayWriteType = "Grey"
        }
    } elseif ($totalPhysicalMemory -gt 192 -and
        $exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) {
        $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to be scaled at or below 192 GB of Memory." -f $totalPhysicalMemory
    } elseif ($totalPhysicalMemory -gt 96 -and
        $exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013) {
        $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to be scaled at or below 96GB of Memory." -f $totalPhysicalMemory
    } else {
        $displayDetails = "{0} GB" -f $totalPhysicalMemory
        $displayWriteType = "Grey"
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "Physical Memory" -Details $displayDetails `
        -DisplayGroupingKey $keyHardwareInformation `
        -DipslayTestingValue $totalPhysicalMemory `
        -DisplayWriteType $displayWriteType `
        -AddHtmlOverviewValues $true `
        -AnalyzedInformation $analyzedResults

    ################################
    #NIC Settings Per Active Adapter
    ################################
    Write-Verbose "Working on NIC Settings Per Active Adapter Information"

    foreach ($adapter in $osInformation.NetworkInformation.NetworkAdapters) {

        if ($adapter.Description -eq "Remote NDIS Compatible Device") {
            Write-Verbose "Remote NDSI Compatible Device found. Ignoring NIC."
            continue
        }

        $value = "{0} [{1}]" -f $adapter.Description, $adapter.Name
        $analyzedResults = Add-AnalyzedResultInformation -Name "Interface Description" -Details $value `
            -DisplayGroupingKey $keyNICSettings `
            -DisplayCustomTabNumber 1 `
            -AnalyzedInformation $analyzedResults

        if ($osInformation.BuildInformation.MajorVersion -ge [HealthChecker.OSServerVersion]::Windows2012R2) {
            Write-Verbose "On Windows 2012 R2 or new. Can provide more details on the NICs"

            $driverDate = $adapter.DriverDate
            $detailsValue = $driverDate

            if ($hardwareInformation.ServerType -eq [HealthChecker.ServerType]::Physical -or
                $hardwareInformation.ServerType -eq [HealthChecker.ServerType]::AmazonEC2) {

                if ($null -eq $driverDate -or
                    $driverDate -eq [DateTime]::MaxValue) {
                    $detailsValue = "Unknown"
                } elseif ((New-TimeSpan -Start $date -End $driverDate).Days -lt [int]-365) {
                    $analyzedResults = Add-AnalyzedResultInformation -Details "Warning: NIC driver is over 1 year old. Verify you are at the latest version." `
                        -DisplayGroupingKey $keyNICSettings `
                        -DisplayWriteType "Yellow" `
                        -AddHtmlDetailRow $false `
                        -AnalyzedInformation $analyzedResults
                }
            }

            $analyzedResults = Add-AnalyzedResultInformation -Name "Driver Date" -Details $detailsValue `
                -DisplayGroupingKey $keyNICSettings `
                -AnalyzedInformation $analyzedResults

            $analyzedResults = Add-AnalyzedResultInformation -Name "Driver Version" -Details ($adapter.DriverVersion) `
                -DisplayGroupingKey $keyNICSettings `
                -AnalyzedInformation $analyzedResults

            $analyzedResults = Add-AnalyzedResultInformation -Name "MTU Size" -Details ($adapter.MTUSize) `
                -DisplayGroupingKey $keyNICSettings `
                -AnalyzedInformation $analyzedResults

            $analyzedResults = Add-AnalyzedResultInformation -Name "Max Processors" -Details ($adapter.NetAdapterRss.MaxProcessors) `
                -DisplayGroupingKey $keyNICSettings `
                -AnalyzedInformation $analyzedResults

            $analyzedResults = Add-AnalyzedResultInformation -Name "Max Processor Number" -Details ($adapter.NetAdapterRss.MaxProcessorNumber) `
                -DisplayGroupingKey $keyNICSettings `
                -AnalyzedInformation $analyzedResults

            $analyzedResults = Add-AnalyzedResultInformation -Name "Number of Receive Queues" -Details ($adapter.NetAdapterRss.NumberOfReceiveQueues) `
                -DisplayGroupingKey $keyNICSettings `
                -AnalyzedInformation $analyzedResults

            $writeType = "Yellow"
            $testingValue = $null

            if ($adapter.RssEnabledValue -eq 0) {
                $detailsValue = "False --- Warning: Enabling RSS is recommended."
                $testingValue = $false
            } elseif ($adapter.RssEnabledValue -eq 1) {
                $detailsValue = "True"
                $testingValue = $true
                $writeType = "Green"
            } else {
                $detailsValue = "No RSS Feature Detected."
            }

            $analyzedResults = Add-AnalyzedResultInformation -Name "RSS Enabled" -Details $detailsValue `
                -DisplayGroupingKey $keyNICSettings `
                -DisplayWriteType $writeType `
                -DisplayTestingValue $testingValue `
                -AnalyzedInformation $analyzedResults
        } else {
            Write-Verbose "On Windows 2012 or older and can't get advanced NIC settings"
        }

        $linkSpeed = $adapter.LinkSpeed
        $displayValue = "{0} --- This may not be accurate due to virtualized hardware" -f $linkSpeed

        if ($hardwareInformation.ServerType -eq [HealthChecker.ServerType]::Physical -or
            $hardwareInformation.ServerType -eq [HealthChecker.ServerType]::AmazonEC2) {
            $displayValue = $linkSpeed
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "Link Speed" -Details $displayValue `
            -DisplayGroupingKey $keyNICSettings `
            -DisplayTestingValue $linkSpeed `
            -AnalyzedInformation $analyzedResults

        $displayValue = "{0}" -f $adapter.IPv6Enabled
        $displayWriteType = "Grey"
        $testingValue = $adapter.IPv6Enabled

        if ($osInformation.NetworkInformation.IPv6DisabledComponents -ne 255 -and
            $adapter.IPv6Enabled -eq $false) {
            $displayValue = "{0} --- Warning" -f $adapter.IPv6Enabled
            $displayWriteType = "Yellow"
            $testingValue = $false
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "IPv6 Enabled" -Details $displayValue `
            -DisplayGroupingKey $keyNICSettings `
            -DisplayWriteType $displayWriteType `
            -DisplayTestingValue $TestingValue `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Name "IPv4 Address" `
            -DisplayGroupingKey $keyNICSettings `
            -AnalyzedInformation $analyzedResults

        foreach ($address in $adapter.IPv4Addresses) {
            $displayValue = "{0}\{1}" -f $address.Address, $address.Subnet

            if ($address.DefaultGateway -ne [string]::Empty) {
                $displayValue += " Gateway: {0}" -f $address.DefaultGateway
            }

            $analyzedResults = Add-AnalyzedResultInformation -Name "Address" -Details $displayValue `
                -DisplayGroupingKey $keyNICSettings `
                -DisplayCustomTabNumber 3 `
                -AnalyzedInformation $analyzedResults
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "IPv6 Address" `
            -DisplayGroupingKey $keyNICSettings `
            -AnalyzedInformation $analyzedResults

        foreach ($address in $adapter.IPv6Addresses) {
            $displayValue = "{0}\{1}" -f $address.Address, $address.Subnet

            if ($address.DefaultGateway -ne [string]::Empty) {
                $displayValue += " Gateway: {0}" -f $address.DefaultGateway
            }

            $analyzedResults = Add-AnalyzedResultInformation -Name "Address" -Details $displayValue `
                -DisplayGroupingKey $keyNICSettings `
                -DisplayCustomTabNumber 3 `
                -AnalyzedInformation $analyzedResults
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "DNS Server" -Details $adapter.DnsServer `
            -DisplayGroupingKey $keyNICSettings `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Name "Registered In DNS" -Details $adapter.RegisteredInDns `
            -DisplayGroupingKey $keyNICSettings `
            -AnalyzedInformation $analyzedResults

        #Assuming that all versions of Hyper-V doesn't allow sleepy NICs
        if (($hardwareInformation.ServerType -ne [HealthChecker.ServerType]::HyperV) -and ($adapter.PnPCapabilities -ne "MultiplexorNoPnP")) {
            $displayWriteType = "Grey"
            $displayValue = $adapter.SleepyNicDisabled

            if (!$adapter.SleepyNicDisabled) {
                $displayWriteType = "Yellow"
                $displayValue = "False --- Warning: It's recommended to disable NIC power saving options`r`n`t`t`tMore Information: https://aka.ms/HC-NICPowerManagement"
            }

            $analyzedResults = Add-AnalyzedResultInformation -Name "Sleepy NIC Disabled" -Details $displayValue `
                -DisplayGroupingKey $keyNICSettings `
                -DisplayWriteType $displayWriteType `
                -DisplayTestingValue $adapter.SleepyNicDisabled `
                -AnalyzedInformation $analyzedResults
        }

        $adapterDescription = $adapter.Description
        $cookedValue = 0
        $foundCounter = $false

        if ($null -eq $osInformation.NetworkInformation.PacketsReceivedDiscarded) {
            Write-Verbose "PacketsReceivedDiscarded is null"
            continue
        }

        foreach ($prdInstance in $osInformation.NetworkInformation.PacketsReceivedDiscarded) {
            $instancePath = $prdInstance.Path
            $startIndex = $instancePath.IndexOf("(") + 1
            $charLength = $instancePath.Substring($startIndex, ($instancePath.IndexOf(")") - $startIndex)).Length
            $instanceName = $instancePath.Substring($startIndex, $charLength)
            $possibleInstanceName = $adapterDescription.Replace("#", "_")

            if ($instanceName -eq $adapterDescription -or
                $instanceName -eq $possibleInstanceName) {
                $cookedValue = $prdInstance.CookedValue
                $foundCounter = $true
                break
            }
        }

        $displayWriteType = "Yellow"
        $displayValue = $cookedValue
        $baseDisplayValue = "{0} --- {1}: This value should be at 0."
        $knownIssue = $false

        if ($foundCounter) {

            if ($cookedValue -eq 0) {
                $displayWriteType = "Green"
            } elseif ($cookedValue -lt 1000) {
                $displayValue = $baseDisplayValue -f $cookedValue, "Warning"
            } else {
                $displayWriteType = "Red"
                $displayValue = [string]::Concat(($baseDisplayValue -f $cookedValue, "Error"), "We are also seeing this value being rather high so this can cause a performance impacted on a system.")
            }

            if ($adapterDescription -like "*vmxnet3*" -and
                $cookedValue -gt 0) {
                $knownIssue = $true
            }
        } else {
            $displayValue = "Couldn't find value for the counter."
            $cookedValue = $null
            $displayWriteType = "Grey"
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "Packets Received Discarded" -Details $displayValue `
            -DisplayGroupingKey $keyNICSettings `
            -DisplayTestingValue $cookedValue `
            -DisplayWriteType $displayWriteType `
            -AnalyzedInformation $analyzedResults

        if ($knownIssue) {
            $analyzedResults = Add-AnalyzedResultInformation -Details "Known Issue with vmxnet3: 'Large packet loss at the guest operating system level on the VMXNET3 vNIC in ESXi (2039495)' - https://aka.ms/HC-VMwareLostPackets" `
                -DisplayGroupingKey $keyNICSettings `
                -DisplayWriteType "Yellow" `
                -DisplayCustomTabNumber 3 `
                -AddHtmlDetailRow $false `
                -AnalyzedInformation $analyzedResults
        }
    }

    if ($osInformation.NetworkInformation.NetworkAdapters.Count -gt 1) {
        $analyzedResults = Add-AnalyzedResultInformation -Details "Multiple active network adapters detected. Exchange 2013 or greater may not need separate adapters for MAPI and replication traffic.  For details please refer to https://aka.ms/HC-PlanHA#network-requirements" `
            -DisplayGroupingKey $keyNICSettings `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $analyzedResults
    }

    if ($osInformation.NetworkInformation.IPv6DisabledOnNICs) {
        $displayWriteType = "Grey"
        $displayValue = "True"
        $testingValue = $true

        if ($osInformation.NetworkInformation.IPv6DisabledComponents -eq -1) {
            $displayWriteType = "Red"
            $testingValue = $false
            $displayValue = "False `r`n`t`tError: IPv6 is disabled on some NIC level settings but not correctly disabled via DisabledComponents registry value. It is currently set to '-1'. `r`n`t`tThis setting cause a system startup delay of 5 seconds. For details please refer to: `r`n`t`thttps://aka.ms/HC-ConfigureIPv6"
        } elseif ($osInformation.NetworkInformation.IPv6DisabledComponents -ne 255) {
            $displayWriteType = "Red"
            $testingValue = $false
            $displayValue = "False `r`n`t`tError: IPv6 is disabled on some NIC level settings but not fully disabled. DisabledComponents registry value currently set to '{0}'. For details please refer to the following articles: `r`n`t`thttps://aka.ms/HC-DisableIPv6`r`n`t`thttps://aka.ms/HC-ConfigureIPv6" -f $osInformation.NetworkInformation.IPv6DisabledComponents
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "Disable IPv6 Correctly" -Details $displayValue `
            -DisplayGroupingKey $keyNICSettings `
            -DisplayWriteType $displayWriteType `
            -DisplayCustomTabNumber 1 `
            -AnalyzedInformation $analyzedResults
    }

    ################
    #TCP/IP Settings
    ################
    Write-Verbose "Working on TCP/IP Settings"

    $tcpKeepAlive = $osInformation.NetworkInformation.TCPKeepAlive

    if ($tcpKeepAlive -eq 0) {
        $displayValue = "Not Set `r`n`t`tError: Without this value the KeepAliveTime defaults to two hours, which can cause connectivity and performance issues between network devices such as firewalls and load balancers depending on their configuration. `r`n`t`tMore details: https://aka.ms/HC-TSPerformanceChecklist"
        $displayWriteType = "Red"
    } elseif ($tcpKeepAlive -lt 900000 -or
        $tcpKeepAlive -gt 1800000) {
        $displayValue = "{0} `r`n`t`tWarning: Not configured optimally, recommended value between 15 to 30 minutes (900000 and 1800000 decimal). `r`n`t`tMore details: https://aka.ms/HC-TSPerformanceChecklist" -f $tcpKeepAlive
        $displayWriteType = "Yellow"
    } else {
        $displayValue = $tcpKeepAlive
        $displayWriteType = "Green"
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "TCP/IP Settings" -Details $displayValue `
        -DisplayGroupingKey $keyFrequentConfigIssues `
        -DisplayWriteType $displayWriteType `
        -DisplayTestingValue $tcpKeepAlive `
        -HtmlName "TCPKeepAlive" `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "RPC Min Connection Timeout" -Details ("{0} `r`n`t`tMore Information: https://aka.ms/HC-RPCSetting" -f $osInformation.NetworkInformation.RpcMinConnectionTimeout) `
        -DisplayGroupingKey $keyFrequentConfigIssues `
        -HtmlName "RPC Minimum Connection Timeout" `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "FIPS Algorithm Policy Enabled" -Details ($exchangeInformation.RegistryValues.FipsAlgorithmPolicyEnabled) `
        -DisplayGroupingKey $keyFrequentConfigIssues `
        -HtmlName "FipsAlgorithmPolicy-Enabled" `
        -AnalyzedInformation $analyzedResults

    $displayValue = $exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage
    $displayWriteType = "Green"

    if ($exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage -ne 0) {
        $displayWriteType = "Red"
        $displayValue = "{0} `r`n`t`tError: This can cause an impact to the server's search performance. This should only be used a temporary fix if no other options are available vs a long term solution." -f $exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "CTS Processor Affinity Percentage" -Details $displayValue `
        -DisplayGroupingKey $keyFrequentConfigIssues `
        -DisplayWriteType $displayWriteType `
        -DisplayTestingValue ($exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage) `
        -HtmlName "CtsProcessorAffinityPercentage" `
        -AnalyzedInformation $analyzedResults

    $displayValue = $osInformation.CredentialGuardEnabled
    $displayWriteType = "Grey"

    if ($osInformation.CredentialGuardEnabled) {
        $displayValue = "{0} `r`n`t`tError: Credential Guard is not supported on an Exchange Server. This can cause a performance hit on the server." -f $osInformation.CredentialGuardEnabled
        $displayWriteType = "Red"
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "Credential Guard Enabled" -Details $displayValue `
        -DisplayGroupingKey $keyFrequentConfigIssues `
        -DisplayWriteType $displayWriteType `
        -AnalyzedInformation $analyzedResults

    if ($null -ne $exchangeInformation.ApplicationConfigFileStatus -and
        $exchangeInformation.ApplicationConfigFileStatus.Count -ge 1) {

        foreach ($configKey in $exchangeInformation.ApplicationConfigFileStatus.Keys) {
            $configStatus = $exchangeInformation.ApplicationConfigFileStatus[$configKey]

            $writeType = "Green"
            $writeName = "{0} Present" -f $configKey
            $writeValue = $configStatus.Present

            if (!$configStatus.Present) {
                $writeType = "Red"
                $writeValue = "{0} --- Error" -f $writeValue
            }

            $analyzedResults = Add-AnalyzedResultInformation -Name $writeName -Details $writeValue `
                -DisplayGroupingKey $keyFrequentConfigIssues `
                -DisplayWriteType $writeType `
                -AnalyzedInformation $analyzedResults
        }
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "LmCompatibilityLevel Settings" -Details ($osInformation.LmCompatibility.RegistryValue) `
        -DisplayGroupingKey $keySecuritySettings `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "Description" -Details ($osInformation.LmCompatibility.Description) `
        -DisplayGroupingKey $keySecuritySettings `
        -DisplayCustomTabNumber 2 `
        -AddHtmlDetailRow $false `
        -AnalyzedInformation $analyzedResults

    ##############
    # TLS Settings
    ##############
    Write-Verbose "Working on TLS Settings"

    $tlsVersions = @("1.0", "1.1", "1.2")
    $currentNetVersion = $osInformation.TLSSettings["NETv4"]

    foreach ($tlsKey in $tlsVersions) {
        $currentTlsVersion = $osInformation.TLSSettings[$tlsKey]

        $analyzedResults = Add-AnalyzedResultInformation -Details ("TLS {0}" -f $tlsKey) `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 1 `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Name ("Server Enabled") -Details ($currentTlsVersion.ServerEnabled) `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Name ("Server Disabled By Default") -Details ($currentTlsVersion.ServerDisabledByDefault) `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Name ("Client Enabled") -Details ($currentTlsVersion.ClientEnabled) `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Name ("Client Disabled By Default") -Details ($currentTlsVersion.ClientDisabledByDefault) `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -AnalyzedInformation $analyzedResults

        if ($currentTlsVersion.ServerEnabled -ne $currentTlsVersion.ClientEnabled) {
            $detectedTlsMismatch = $true
            $analyzedResults = Add-AnalyzedResultInformation -Details ("Error: Mismatch in TLS version for client and server. Exchange can be both client and a server. This can cause issues within Exchange for communication.") `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 3 `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults
        }

        if (($tlsKey -eq "1.0" -or
                $tlsKey -eq "1.1") -and (
                $currentTlsVersion.ServerEnabled -eq $false -or
                $currentTlsVersion.ClientEnabled -eq $false -or
                $currentTlsVersion.ServerDisabledByDefault -or
                $currentTlsVersion.ClientDisabledByDefault) -and
            ($currentNetVersion.SystemDefaultTlsVersions -eq $false -or
            $currentNetVersion.WowSystemDefaultTlsVersions -eq $false)) {
            $analyzedResults = Add-AnalyzedResultInformation -Details ("Error: SystemDefaultTlsVersions is not set to the recommended value. Please visit on how to properly enable TLS 1.2 https://aka.ms/HC-TLSPart2") `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 3 `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults
        }
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "SystemDefaultTlsVersions" -Details ($currentNetVersion.SystemDefaultTlsVersions) `
        -DisplayGroupingKey $keySecuritySettings `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "SystemDefaultTlsVersions - Wow6432Node" -Details ($currentNetVersion.WowSystemDefaultTlsVersions) `
        -DisplayGroupingKey $keySecuritySettings `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "SchUseStrongCrypto" -Details ($currentNetVersion.SchUseStrongCrypto) `
        -DisplayGroupingKey $keySecuritySettings `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "SchUseStrongCrypto - Wow6432Node" -Details ($currentNetVersion.WowSchUseStrongCrypto) `
        -DisplayGroupingKey $keySecuritySettings `
        -AnalyzedInformation $analyzedResults

    $analyzedResults = Add-AnalyzedResultInformation -Name "SecurityProtocol" -Details ($currentNetVersion.SecurityProtocol) `
        -DisplayGroupingKey $keySecuritySettings `
        -AnalyzedInformation $analyzedResults

    <#
    [array]$securityProtocols = $currentNetVersion.SecurityProtocol.Split(",").Trim().ToUpper()
    $lowerTLSVersions = @("1.0", "1.1")

    foreach ($tlsKey in $lowerTLSVersions) {
        $currentTlsVersion = $osInformation.TLSSettings[$tlsKey]
        $securityProtocolCheck = "TLS"
        if ($tlsKey -eq "1.1") {
            $securityProtocolCheck = "TLS11"
        }

        if (($currentTlsVersion.ServerEnabled -eq $false -or
                $currentTlsVersion.ClientEnabled -eq $false) -and
            $securityProtocols.Contains($securityProtocolCheck)) {

            $analyzedResults = Add-AnalyzedResultInformation -Details ("Security Protocol is able to use TLS when we have TLS {0} disabled in the registry. This can cause issues with connectivity. It is recommended to follow the proper TLS settings. In some cases, it may require to also set SchUseStrongCrypto in the registry." -f $tlsKey) `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType "Yellow" `
                -AnalyzedInformation $analyzedResults
        }
    }
#>

    if ($detectedTlsMismatch) {
        $displayValues = @("Exchange Server TLS guidance Part 1: Getting Ready for TLS 1.2: https://aka.ms/HC-TLSPart1",
            "Exchange Server TLS guidance Part 2: Enabling TLS 1.2 and Identifying Clients Not Using It: https://aka.ms/HC-TLSPart2",
            "Exchange Server TLS guidance Part 3: Turning Off TLS 1.0/1.1: https://aka.ms/HC-TLSPart3")

        $analyzedResults = Add-AnalyzedResultInformation -Details "For More Information on how to properly set TLS follow these blog posts:" `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -DisplayWriteType "Yellow" `
            -AnalyzedInformation $analyzedResults

        foreach ($displayValue in $displayValues) {
            $analyzedResults = Add-AnalyzedResultInformation -Details $displayValue `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayWriteType "Yellow" `
                -DisplayCustomTabNumber 3 `
                -AnalyzedInformation $analyzedResults
        }
    }

    foreach ($certificate in $exchangeInformation.ExchangeCertificates) {

        if ($certificate.LifetimeInDays -ge 60) {
            $displayColor = "Green"
        } elseif ($certificate.LifetimeInDays -ge 30) {
            $displayColor = "Yellow"
        } else {
            $displayColor = "Red"
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "Certificate" `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 1 `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Name "FriendlyName" -Details $certificate.FriendlyName `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Name "Thumbprint" -Details $certificate.Thumbprint `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Name "Lifetime in days" -Details $certificate.LifetimeInDays `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -DisplayWriteType $displayColor `
            -AnalyzedInformation $analyzedResults

        if ($certificate.LifetimeInDays -lt 0) {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Certificate has expired" -Details $true `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults
        } else {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Certificate has expired" -Details $false `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -AnalyzedInformation $analyzedResults
        }

        $certStatusWriteType = [string]::Empty

        if ($null -ne $certificate.Status) {
            Switch ($certificate.Status) {
                ("Unknown") { $certStatusWriteType = "Yellow" }
                ("Valid") { $certStatusWriteType = "Grey" }
                ("Revoked") { $certStatusWriteType = "Red" }
                ("DateInvalid") { $certStatusWriteType = "Red" }
                ("Untrusted") { $certStatusWriteType = "Yellow" }
                ("Invalid") { $certStatusWriteType = "Red" }
                ("RevocationCheckFailure") { $certStatusWriteType = "Yellow" }
                ("PendingRequest") { $certStatusWriteType = "Yellow" }
                default { $certStatusWriteType = "Yellow" }
            }

            $analyzedResults = Add-AnalyzedResultInformation -Name "Certificate status" -Details $certificate.Status `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType $certStatusWriteType `
                -AnalyzedInformation $analyzedResults
        } else {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Certificate status" -Details "Unknown" `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType "Yellow" `
                -AnalyzedInformation $analyzedResults
        }

        if ($certificate.PublicKeySize -lt 2048) {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Key size" -Details $certificate.PublicKeySize `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults

            $analyzedResults = Add-AnalyzedResultInformation -Details "It's recommended to use a key size of at least 2048 bit" `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults
        } else {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Key size" -Details $certificate.PublicKeySize `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -AnalyzedInformation $analyzedResults
        }

        if ($certificate.SignatureHashAlgorithmSecure -eq 1) {
            $shaDisplayWriteType = "Yellow"
        } else {
            $shaDisplayWriteType = "Grey"
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "Signature Algorithm" -Details $certificate.SignatureAlgorithm `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -DisplayWriteType $shaDisplayWriteType `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Name "Signature Hash Algorithm" -Details $certificate.SignatureHashAlgorithm `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -DisplayWriteType $shaDisplayWriteType `
            -AnalyzedInformation $analyzedResults

        if ($shaDisplayWriteType -eq "Yellow") {
            $analyzedResults = Add-AnalyzedResultInformation -Details "It's recommended to use a hash algorithm from the SHA-2 family `r`n`t`tMore information: https://aka.ms/HC-SSLBP" `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType $shaDisplayWriteType `
                -AnalyzedInformation $analyzedResults
        }

        if ($null -ne $certificate.Services) {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Bound to services" -Details $certificate.Services `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -AnalyzedInformation $analyzedResults
        }

        if ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Current Auth Certificate" -Details $certificate.IsCurrentAuthConfigCertificate `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -AnalyzedInformation $analyzedResults
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "SAN Certificate" -Details $certificate.IsSanCertificate `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Name "Namespaces" `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -AnalyzedInformation $analyzedResults

        foreach ($namespace in $certificate.Namespaces) {
            $analyzedResults = Add-AnalyzedResultInformation -Details $namespace `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 3 `
                -AnalyzedInformation $analyzedResults
        }

        if ($certificate.IsCurrentAuthConfigCertificate -eq $true) {
            $currentAuthCertificate = $certificate
        }
    }

    if ($null -ne $currentAuthCertificate) {
        if ($currentAuthCertificate.LifetimeInDays -gt 0) {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Valid Auth Certificate Found On Server" -Details $true `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 1 `
                -DisplayWriteType "Green" `
                -AnalyzedInformation $analyzedResults
        } else {
            $analyzedResults = Add-AnalyzedResultInformation -Name "Valid Auth Certificate Found On Server" -Details $false `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 1 `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults

            $renewExpiredAuthCert = "Auth Certificate has expired `r`n`t`tMore Information: https://aka.ms/HC-OAuthExpired"
            $analyzedResults = Add-AnalyzedResultInformation -Details $renewExpiredAuthCert `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults
        }
    } elseif ($exchangeInformation.BuildInformation.ServerRole -eq [HealthChecker.ExchangeServerRole]::Edge) {
        $analyzedResults = Add-AnalyzedResultInformation -Name "Valid Auth Certificate Found On Server" -Details $false `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 1 `
            -AnalyzedInformation $analyzedResults

        $analyzedResults = Add-AnalyzedResultInformation -Details "We can't check for Auth Certificates on Edge Transport Servers" `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -AnalyzedInformation $analyzedResults
    } else {
        $analyzedResults = Add-AnalyzedResultInformation -Name "Valid Auth Certificate Found On Server" -Details $false `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 1 `
            -DisplayWriteType "Red" `
            -AnalyzedInformation $analyzedResults

        $createNewAuthCert = "No valid Auth Certificate found. This may cause several problems. `r`n`t`tMore Information: https://aka.ms/HC-FindOAuthHybrid"
        $analyzedResults = Add-AnalyzedResultInformation -Details $createNewAuthCert `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayCustomTabNumber 2 `
            -DisplayWriteType "Red" `
            -AnalyzedInformation $analyzedResults
    }

    $additionalDisplayValue = [string]::Empty
    $smb1Settings = $osInformation.Smb1ServerSettings

    if ($osInformation.BuildInformation.MajorVersion -gt [HealthChecker.OSServerVersion]::Windows2012) {
        $displayValue = "False"
        $writeType = "Green"

        if (-not ($smb1Settings.SuccessfulGetInstall)) {
            $displayValue = "Failed to get install status"
            $writeType = "Yellow"
        } elseif ($smb1Settings.Installed) {
            $displayValue = "True"
            $writeType = "Red"
            $additionalDisplayValue = "SMB1 should be uninstalled"
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "SMB1 Installed" -Details $displayValue `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayWriteType $writeType `
            -AnalyzedInformation $analyzedResults
    }

    $writeType = "Green"
    $displayValue = "True"

    if (-not ($smb1Settings.SuccessfulGetBlocked)) {
        $displayValue = "Failed to get block status"
        $writeType = "Yellow"
    } elseif (-not($smb1Settings.IsBlocked)) {
        $displayValue = "False"
        $writeType = "Red"
        $additionalDisplayValue += " SMB1 should be blocked"
    }

    $analyzedResults = Add-AnalyzedResultInformation -Name "SMB1 Blocked" -Details $displayValue `
        -DisplayGroupingKey $keySecuritySettings `
        -DisplayWriteType $writeType `
        -AnalyzedInformation $analyzedResults

    if ($additionalDisplayValue -ne [string]::Empty) {
        $additionalDisplayValue += "`r`n`t`tMore Information: https://aka.ms/HC-SMB1"

        $analyzedResults = Add-AnalyzedResultInformation -Details $additionalDisplayValue.Trim() `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayWriteType "Yellow" `
            -DisplayCustomTabNumber 2 `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $analyzedResults
    }

    ##########################
    #Exchange Web App GC Mode#
    ##########################
    if ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {
        Write-Verbose "Working on Exchange Web App GC Mode"

        $outputObjectDisplayValue = New-Object System.Collections.Generic.List[object]
        foreach ($webAppKey in $exchangeInformation.ApplicationPools.Keys) {

            $appPool = $exchangeInformation.ApplicationPools[$webAppKey]
            $appRestarts = $appPool.AppSettings.add.recycling.periodicRestart
            $appRestartSet = ($appRestarts.PrivateMemory -ne "0" -or
                $appRestarts.Memory -ne "0" -or
                $appRestarts.Requests -ne "0" -or
                $null -ne $appRestarts.Schedule -or
                ($appRestarts.Time -ne "00:00:00" -and
                    ($webAppKey -ne "MSExchangeOWAAppPool" -and
                $webAppKey -ne "MSExchangeECPAppPool")))

            $outputObjectDisplayValue.Add(([PSCustomObject]@{
                        AppPoolName         = $webAppKey
                        State               = $appPool.AppSettings.state
                        GCServerEnabled     = $appPool.GCServerEnabled
                        RestartConditionSet = $appRestartSet
                    })
            )
        }

        $sbStarted = { param($o, $p) if ($p -eq "State") { if ($o."$p" -eq "Started") { "Green" } else { "Red" } } }
        $sbRestart = { param($o, $p) if ($p -eq "RestartConditionSet") { if ($o."$p") { "Red" } else { "Green" } } }
        $analyzedResults = Add-AnalyzedResultInformation -OutColumns ([PSCustomObject]@{
                DisplayObject      = $outputObjectDisplayValue
                ColorizerFunctions = @($sbStarted, $sbRestart)
                IndentSpaces       = 8
            }) `
            -DisplayGroupingKey $keyWebApps `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $analyzedResults

        $periodicStartAppPools = $outputObjectDisplayValue | Where-Object { $_.RestartConditionSet -eq $true }

        if ($null -ne $periodicStartAppPools) {

            $outputObjectDisplayValue = New-Object System.Collections.Generic.List[object]

            foreach ($appPool in $periodicStartAppPools) {
                $periodicRestart = $exchangeInformation.ApplicationPools[$appPool.AppPoolName].AppSettings.add.recycling.periodicRestart
                $schedule = $periodicRestart.Schedule

                if ([string]::IsNullOrEmpty($schedule)) {
                    $schedule = "null"
                }

                $outputObjectDisplayValue.Add(([PSCustomObject]@{
                            AppPoolName   = $appPool.AppPoolName
                            PrivateMemory = $periodicRestart.PrivateMemory
                            Memory        = $periodicRestart.Memory
                            Requests      = $periodicRestart.Requests
                            Schedule      = $schedule
                            Time          = $periodicRestart.Time
                        }))
            }

            $sbColorizer = {
                param($o, $p)
                switch ($p) {
                    { $_ -in "PrivateMemory", "Memory", "Requests" } {
                        if ($o."$p" -eq "0") { "Green" } else { "Red" }
                    }
                    "Time" {
                        if ($o."$p" -eq "00:00:00") { "Green" } else { "Red" }
                    }
                    "Schedule" {
                        if ($o."$p" -eq "null") { "Green" } else { "Red" }
                    }
                }
            }

            $analyzedResults = Add-AnalyzedResultInformation -OutColumns ([PSCustomObject]@{
                    DisplayObject      = $outputObjectDisplayValue
                    ColorizerFunctions = @($sbColorizer)
                    IndentSpaces       = 8
                }) `
                -DisplayGroupingKey $keyWebApps `
                -AddHtmlDetailRow $false `
                -AnalyzedInformation $analyzedResults

            $analyzedResults = Add-AnalyzedResultInformation -Details "Error: The above app pools currently have the periodic restarts set. This restart will cause disruption to end users." `
                -DisplayGroupingKey $keyWebApps `
                -DisplayWriteType "Red" `
                -AddHtmlDetailRow $false `
                -AnalyzedInformation $analyzedResults
        }
    }

    ######################
    # Vulnerability Checks
    ######################

    Function Test-VulnerabilitiesByBuildNumbersForDisplay {
        param(
            [Parameter(Mandatory = $true)][string]$ExchangeBuildRevision,
            [Parameter(Mandatory = $true)][array]$SecurityFixedBuilds,
            [Parameter(Mandatory = $true)][array]$CVENames
        )
        [int]$fileBuildPart = ($split = $ExchangeBuildRevision.Split("."))[0]
        [int]$filePrivatePart = $split[1]
        $Script:breakpointHit = $false

        foreach ($securityFixedBuild in $SecurityFixedBuilds) {
            [int]$securityFixedBuildPart = ($split = $securityFixedBuild.Split("."))[0]
            [int]$securityFixedPrivatePart = $split[1]

            if ($fileBuildPart -eq $securityFixedBuildPart) {
                $Script:breakpointHit = $true
            }

            if (($fileBuildPart -lt $securityFixedBuildPart) -or
                ($fileBuildPart -eq $securityFixedBuildPart -and
                $filePrivatePart -lt $securityFixedPrivatePart)) {
                foreach ($cveName in $CVENames) {
                    $details = "{0}`r`n`t`tSee: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/{0} for more information." -f $cveName
                    $Script:Vulnerabilities += $details
                    $Script:AnalyzedInformation = Add-AnalyzedResultInformation -Name "Security Vulnerability" -Details $details `
                        -DisplayGroupingKey $keySecuritySettings `
                        -DisplayTestingValue $cveName `
                        -DisplayWriteType "Red" `
                        -AddHtmlDetailRow $false `
                        -AnalyzedInformation $Script:AnalyzedInformation
                }

                $Script:AllVulnerabilitiesPassed = $false
                break
            }

            if ($Script:breakpointHit) {
                break
            }
        }
    }

    Function Show-March2021SUOutdatedCUWarning {
        param(
            [Parameter(Mandatory = $true)][hashtable]$KBCVEHT
        )
        Write-Verbose "Calling: Show-March2021SUOutdatedCUWarning"

        foreach ($kbName in $KBCVEHT.Keys) {
            foreach ($cveName in $KBCVEHT[$kbName]) {
                $details = "`r`n`t`tPlease make sure {0} is installed to be fully protected against: {1}" -f $kbName, $cveName
                $Script:AnalyzedInformation = Add-AnalyzedResultInformation -Name "March 2021 Exchange Security Update for unsupported CU detected" -Details $details `
                    -DisplayGroupingKey $keySecuritySettings `
                    -DisplayTestingValue $cveName `
                    -DisplayWriteType "Yellow" `
                    -AddHtmlDetailRow $false `
                    -AnalyzedInformation $Script:AnalyzedInformation
            }
        }
    }

    Function Test-DownloadDomainsConfiguration {
        param(
            [Parameter(Mandatory = $true)][object]$OwaVDirObject,
            [Parameter(Mandatory = $true)][bool]$DownloadDomainsEnabled
        )
        Write-Verbose "Calling: Test-DownloadDomainConfiguration"

        <#
        Unknown 0
        Download Domains disabled 1
        Download Domains enabled and configured as expected 2
        Download Domains enabled and external download host name = internal/external owa url 4
        Download Domains enabled but external download host name not set 8
        Download Domains enabled and internal download host name = internal/external owa url 16
        Download Domains enabled but internal download host name not set 32
        #>

        $downloadDomainsStatus = 0

        if ($DownloadDomainsEnabled) {
            $downloadDomainsStatus += 2

            if (![String]::IsNullOrEmpty($OwaVDirObject.ExternalDownloadHostName)) {

                if (($OwaVDirObject.ExternalDownloadHostName -eq $OwaVDirObject.ExternalUrl.Host) -or
                    ($OwaVDirObject.ExternalDownloadHostName -eq $OwaVDirObject.InternalUrl.Host)) {
                    $downloadDomainsStatus += 4
                }
            } else {
                Write-Verbose "'ExternalDownloadHostName' is not configured"
                $downloadDomainsStatus += 8
            }

            if (![String]::IsNullOrEmpty($OwaVDirObject.InternalDownloadHostName)) {

                if (($OwaVDirObject.InternalDownloadHostName -eq $OwaVDirObject.ExternalUrl.Host) -or
                    ($OwaVDirObject.InternalDownloadHostName -eq $OwaVDirObject.InternalUrl.Host)) {
                    $downloadDomainsStatus += 16
                }
            } else {
                Write-Verbose "'InternalDownloadHostName' is not configured"
                $downloadDomainsStatus += 32
            }

            return $downloadDomainsStatus
        } else {
            $downloadDomainsStatus += 1
            return $downloadDomainsStatus
        }
    }

    $Script:AllVulnerabilitiesPassed = $true
    $Script:Vulnerabilities = @()
    $Script:AnalyzedInformation = $analyzedResults
    [string]$buildRevision = ("{0}.{1}" -f $exchangeInformation.BuildInformation.ExchangeSetup.FileBuildPart, $exchangeInformation.BuildInformation.ExchangeSetup.FilePrivatePart)
    $exchangeCU = $exchangeInformation.BuildInformation.CU
    Write-Verbose "Exchange Build Revision: $buildRevision"
    Write-Verbose "Exchange CU: $exchangeCU"

    if ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013) {

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU19) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1347.5", "1365.3" -CVENames "CVE-2018-0924", "CVE-2018-0940"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU20) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1365.7", "1367.6" -CVENames "CVE-2018-8151", "CVE-2018-8154", "CVE-2018-8159"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU21) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1367.9", "1395.7" -CVENames "CVE-2018-8302"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1395.8" -CVENames "CVE-2018-8265", "CVE-2018-8448"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1395.10" -CVENames "CVE-2019-0586", "CVE-2019-0588"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU22) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1473.3" -CVENames "CVE-2019-0686", "CVE-2019-0724"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1473.4" -CVENames "CVE-2019-0817", "CVE-2019-0858"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1473.5" -CVENames "ADV190018"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU23) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1497.3" -CVENames "CVE-2019-1084", "CVE-2019-1136", "CVE-2019-1137"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1497.4" -CVENames "CVE-2019-1373"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1497.6" -CVENames "CVE-2020-0688", "CVE-2020-0692"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1497.7" -CVENames "CVE-2020-16969"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1497.8" -CVENames "CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1497.10" -CVENames "CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17142", "CVE-2020-17143"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1395.12", "1473.6", "1497.12" -CVENames "CVE-2021-26855", "CVE-2021-26857", "CVE-2021-26858", "CVE-2021-27065"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1497.12" -CVENames "CVE-2021-26412", "CVE-2021-27078", "CVE-2021-26854"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1497.15" -CVENames "CVE-2021-28480", "CVE-2021-28481", "CVE-2021-28482", "CVE-2021-28483"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1497.18" -CVENames "CVE-2021-31195", "CVE-2021-31198", "CVE-2021-31207", "CVE-2021-31209"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1497.23" -CVENames "CVE-2021-31206", "CVE-2021-31196", "CVE-2021-33768"
        }
    } elseif ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) {

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU8) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1261.39", "1415.4" -CVENames "CVE-2018-0924", "CVE-2018-0940", "CVE-2018-0941"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU9) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1415.7", "1466.8" -CVENames "CVE-2018-8151", "CVE-2018-8152", "CVE-2018-8153", "CVE-2018-8154", "CVE-2018-8159"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU10) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1466.9", "1531.6" -CVENames "CVE-2018-8374", "CVE-2018-8302"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1531.8" -CVENames "CVE-2018-8265", "CVE-2018-8448"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU11) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1531.8", "1591.11" -CVENames "CVE-2018-8604"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1531.10", "1591.13" -CVENames "CVE-2019-0586", "CVE-2019-0588"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU12) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1591.16", "1713.6" -CVENames "CVE-2019-0817", "CVE-2018-0858"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1591.17", "1713.7" -CVENames "ADV190018"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1713.5" -CVENames "CVE-2019-0686", "CVE-2019-0724"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU13) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1713.8", "1779.4" -CVENames "CVE-2019-1084", "CVE-2019-1136", "CVE-2019-1137"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1713.9", "1779.5" -CVENames "CVE-2019-1233", "CVE-2019-1266"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU14) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1779.7", "1847.5" -CVENames "CVE-2019-1373"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU15) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1847.7", "1913.7" -CVENames "CVE-2020-0688", "CVE-2020-0692"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1847.10", "1913.10" -CVENames "CVE-2020-0903"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU17) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1979.6", "2044.6" -CVENames "CVE-2020-16875"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU18) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "2106.2" -CVENames "CVE-2021-1730"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "2044.7", "2106.3" -CVENames "CVE-2020-16969"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "2044.8", "2106.4" -CVENames "CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "2044.12", "2106.6" -CVENames "CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17141", "CVE-2020-17142", "CVE-2020-17143"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU19) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "2106.8", "2176.4" -CVENames "CVE-2021-24085"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1415.8", "1466.13", "1531.12", "1591.18", "1713.10", "1779.8", "1847.12", "1913.12", "1979.8", "2044.13", "2106.13", "2176.9" -CVENames "CVE-2021-26855", "CVE-2021-26857", "CVE-2021-26858", "CVE-2021-27065"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "2106.13", "2176.9" -CVENames "CVE-2021-26412", "CVE-2021-27078", "CVE-2021-26854"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU20) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "2176.12", "2242.8" -CVENames "CVE-2021-28480", "CVE-2021-28481", "CVE-2021-28482", "CVE-2021-28483"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "2176.14", "2242.10" -CVENames "CVE-2021-31195", "CVE-2021-31198", "CVE-2021-31207", "CVE-2021-31209"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU21) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "2242.12", "2308.14" -CVENames "CVE-2021-31206", "CVE-2021-31196", "CVE-2021-33768"
        }

        if ($exchangeCU -ge [HealthChecker.ExchangeCULevel]::CU22) {
            Write-Verbose "There are no known vulnerabilities in this build"
        }
    } elseif ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) {

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU1) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "221.14" -CVENames "CVE-2019-0586", "CVE-2019-0588"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "221.16", "330.7" -CVENames "CVE-2019-0817", "CVE-2019-0858"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "221.17", "330.8" -CVENames "ADV190018"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "330.6" -CVENames "CVE-2019-0686", "CVE-2019-0724"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU2) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "330.9", "397.5" -CVENames "CVE-2019-1084", "CVE-2019-1137"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "397.6", "330.10" -CVENames "CVE-2019-1233", "CVE-2019-1266"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU3) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "397.9", "464.7" -CVENames "CVE-2019-1373"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU4) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "464.11", "529.8" -CVENames "CVE-2020-0688", "CVE-2020-0692"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "464.14", "529.11" -CVENames "CVE-2020-0903"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU6) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "595.6", "659.6" -CVENames "CVE-2020-16875"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU7) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "721.2" -CVENames "CVE-2021-1730"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "659.7", "721.3" -CVENames "CVE-2020-16969"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "659.8", "721.4" -CVENames "CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "659.11", "721.6" -CVENames "CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17141", "CVE-2020-17142", "CVE-2020-17143"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU8) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "721.8", "792.5" -CVENames "CVE-2021-24085"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "221.18", "330.11", "397.11", "464.15", "529.13", "595.8", "659.12", "721.13", "792.10" -CVENames "CVE-2021-26855", "CVE-2021-26857", "CVE-2021-26858", "CVE-2021-27065"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "721.13", "792.10" -CVENames "CVE-2021-26412", "CVE-2021-27078", "CVE-2021-26854"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU9) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "792.13", "858.10" -CVENames "CVE-2021-28480", "CVE-2021-28481", "CVE-2021-28482", "CVE-2021-28483"
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "792.15", "858.12" -CVENames "CVE-2021-31195", "CVE-2021-31198", "CVE-2021-31207", "CVE-2021-31209"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU10) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "858.15", "922.13" -CVENames "CVE-2021-31206", "CVE-2021-31196", "CVE-2021-33768"
        }

        if ($exchangeCU -ge [HealthChecker.ExchangeCULevel]::CU11) {
            Write-Verbose "There are no known vulnerabilities in this build"
        }
    } else {
        Write-Verbose "Unknown Version of Exchange"
        $Script:AllVulnerabilitiesPassed = $false
    }

    #Description: March 2021 Exchange vulnerabilities Security Update (SU) check for outdated version (CUs)
    #Affected Exchange versions: Exchange 2013, Exchange 2016, Exchange 2016 (we only provide this special SU for these versions)
    #Fix: Update to a supported CU and apply KB5000871

    if (($exchangeInformation.BuildInformation.March2021SUInstalled) -and ($exchangeInformation.BuildInformation.SupportedBuild -eq $false)) {
        if (($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013) -and
            ($exchangeCU -lt [HealthChecker.ExchangeCULevel]::CU23)) {
            Switch ($exchangeCU) {
                ([HealthChecker.ExchangeCULevel]::CU21) { $KBCveComb = @{KB4340731 = "CVE-2018-8302"; KB4459266 = "CVE-2018-8265", "CVE-2018-8448"; KB4471389 = "CVE-2019-0586", "CVE-2019-0588" } }
                ([HealthChecker.ExchangeCULevel]::CU22) { $KBCveComb = @{KB4487563 = "CVE-2019-0817", "CVE-2019-0858"; KB4503027 = "ADV190018" } }
            }
            $Script:AllVulnerabilitiesPassed = $false
        } elseif (($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) -and
            ($exchangeCU -lt [HealthChecker.ExchangeCULevel]::CU18)) {
            Switch ($exchangeCU) {
                ([HealthChecker.ExchangeCULevel]::CU8) { $KBCveComb = @{KB4073392 = "CVE-2018-0924", "CVE-2018-0940", "CVE-2018-0941"; KB4092041 = "CVE-2018-8151", "CVE-2018-8152", "CVE-2018-8153", "CVE-2018-8154", "CVE-2018-8159" } }
                ([HealthChecker.ExchangeCULevel]::CU9) { $KBCveComb = @{KB4092041 = "CVE-2018-8151", "CVE-2018-8152", "CVE-2018-8153", "CVE-2018-8154", "CVE-2018-8159"; KB4340731 = "CVE-2018-8374", "CVE-2018-8302" } }
                ([HealthChecker.ExchangeCULevel]::CU10) { $KBCveComb = @{KB4340731 = "CVE-2018-8374", "CVE-2018-8302"; KB4459266 = "CVE-2018-8265", "CVE-2018-8448"; KB4468741 = "CVE-2018-8604"; KB4471389 = "CVE-2019-0586", "CVE-2019-0588" } }
                ([HealthChecker.ExchangeCULevel]::CU11) { $KBCveComb = @{KB4468741 = "CVE-2018-8604"; KB4471389 = "CVE-2019-0586", "CVE-2019-0588"; KB4487563 = "CVE-2019-0817", "CVE-2018-0858"; KB4503027 = "ADV190018" } }
                ([HealthChecker.ExchangeCULevel]::CU12) { $KBCveComb = @{KB4487563 = "CVE-2019-0817", "CVE-2018-0858"; KB4503027 = "ADV190018"; KB4515832 = "CVE-2019-1233", "CVE-2019-1266" } }
                ([HealthChecker.ExchangeCULevel]::CU13) { $KBCveComb = @{KB4509409 = "CVE-2019-1084", "CVE-2019-1136", "CVE-2019-1137"; KB4515832 = "CVE-2019-1233", "CVE-2019-1266"; KB4523171 = "CVE-2019-1373" } }
                ([HealthChecker.ExchangeCULevel]::CU14) { $KBCveComb = @{KB4523171 = "CVE-2019-1373"; KB4536987 = "CVE-2020-0688", "CVE-2020-0692"; KB4540123 = "CVE-2020-0903" } }
                ([HealthChecker.ExchangeCULevel]::CU15) { $KBCveComb = @{KB4536987 = "CVE-2020-0688", "CVE-2020-0692"; KB4540123 = "CVE-2020-0903" } }
                ([HealthChecker.ExchangeCULevel]::CU16) { $KBCveComb = @{KB4577352 = "CVE-2020-16875" } }
                ([HealthChecker.ExchangeCULevel]::CU17) { $KBCveComb = @{KB4577352 = "CVE-2020-16875"; KB4581424 = "CVE-2020-16969"; KB4588741 = "CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085"; KB4593465 = "CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17141", "CVE-2020-17142", "CVE-2020-17143" } }
            }
            $Script:AllVulnerabilitiesPassed = $false
        } elseif (($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) -and
            ($exchangeCU -lt [HealthChecker.ExchangeCULevel]::CU7)) {
            Switch ($exchangeCU) {
                ([HealthChecker.ExchangeCULevel]::RTM) { $KBCveComb = @{KB4471389 = "CVE-2019-0586", "CVE-2019-0588"; KB4487563 = "CVE-2019-0817", "CVE-2019-0858"; KB4503027 = "ADV190018" } }
                ([HealthChecker.ExchangeCULevel]::CU1) { $KBCveComb = @{KB4487563 = "CVE-2019-0817", "CVE-2019-0858"; KB4503027 = "ADV190018"; KB4509409 = "CVE-2019-1084", "CVE-2019-1137"; KB4515832 = "CVE-2019-1233", "CVE-2019-1266" } }
                ([HealthChecker.ExchangeCULevel]::CU2) { $KBCveComb = @{KB4509409 = "CVE-2019-1084", "CVE-2019-1137"; KB4515832 = "CVE-2019-1233", "CVE-2019-1266"; KB4523171 = "CVE-2019-1373" } }
                ([HealthChecker.ExchangeCULevel]::CU3) { $KBCveComb = @{KB4523171 = "CVE-2019-1373"; KB4536987 = "CVE-2020-0688", "CVE-2020-0692"; KB4540123 = "CVE-2020-0903" } }
                ([HealthChecker.ExchangeCULevel]::CU4) { $KBCveComb = @{KB4536987 = "CVE-2020-0688", "CVE-2020-0692"; KB4540123 = "CVE-2020-0903" } }
                ([HealthChecker.ExchangeCULevel]::CU5) { $KBCveComb = @{KB4577352 = "CVE-2020-16875" } }
                ([HealthChecker.ExchangeCULevel]::CU6) { $KBCveComb = @{KB4577352 = "CVE-2020-16875"; KB4581424 = "CVE-2020-16969"; KB4588741 = "CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085"; KB4593465 = "CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17141", "CVE-2020-17142", "CVE-2020-17143" } }
            }
            $Script:AllVulnerabilitiesPassed = $false
        } else {
            Write-Verbose "No need to call 'Show-March2021SUOutdatedCUWarning'"
        }
        if ($null -ne $KBCveComb) {
            Show-March2021SUOutdatedCUWarning -KBCVEHT $KBCveComb
        }
    }

    #Description: Check for CVE-2021-34470 rights elevation vulnerability
    #Affected Exchange versions: 2013, 2016, 2019
    #Fix:
    ##Exchange 2013 CU23 + July 2021 SU + /PrepareSchema,
    ##Exchange 2016 CU20 + July 2021 SU + /PrepareSchema or CU21,
    ##Exchange 2019 CU9 + July 2021 SU + /PrepareSchema or CU10
    #Workaround: N/A

    if (($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013) -or
        (($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) -and
            ($exchangeCU -lt [HealthChecker.ExchangeCULevel]::CU21)) -or
        (($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) -and
            ($exchangeCU -lt [HealthChecker.ExchangeCULevel]::CU10))) {
        Write-Verbose "Testing CVE: CVE-2021-34470"

        $displayWriteTypeColor = $null
        if ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013) {
            Test-VulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision -SecurityFixedBuilds "1497.23" -CVENames "CVE-2021-34470"
        }

        if ($null -eq $exchangeInformation.msExchStorageGroup) {
            Write-Verbose "Unable to query classSchema: 'ms-Exch-Storage-Group' information"
            $details = "CVE-2021-34470`r`n`t`tWarning: Unable to query classSchema: 'ms-Exch-Storage-Group' to perform testing."
            $displayWriteTypeColor = "Yellow"
        } elseif ($exchangeInformation.msExchStorageGroup.Properties.posssuperiors -eq "computer") {
            Write-Verbose "Attribute: 'possSuperiors' with value: 'computer' detected in classSchema: 'ms-Exch-Storage-Group'"
            $details = "CVE-2021-34470`r`n`t`tPrepareSchema required: https://aka.ms/HC-July21SU"
            $displayWriteTypeColor = "Red"
        } else {
            Write-Verbose "System NOT vulnerable to CVE-2021-34470"
        }

        if ($null -ne $displayWriteTypeColor) {
            $Script:Vulnerabilities += $details
            $analyzedResults = Add-AnalyzedResultInformation -Name "Security Vulnerability" -Details $details `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayWriteType $displayWriteTypeColor `
                -AddHtmlDetailRow $false `
                -AnalyzedInformation $analyzedResults
            $Script:AllVulnerabilitiesPassed = $false
        }
    } else {
        Write-Verbose "System NOT vulnerable to CVE-2021-34470"
    }

    #Description: Check for CVE-2021-1730 vulnerability
    #Fix available for: Exchange 2016 CU18+, Exchange 2019 CU7+
    #Fix: Configure Download Domains feature
    #Workaround: N/A

    if (((($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) -and
                ($exchangeCU -ge [HealthChecker.ExchangeCULevel]::CU18)) -or
            (($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) -and
                ($exchangeCU -ge [HealthChecker.ExchangeCULevel]::CU7))) -and
        $exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {

        $downloadDomainsConfig = Test-DownloadDomainsConfiguration -OwaVDirObject $exchangeInformation.GetOwaVirtualDirectory -DownloadDomainsEnabled $exchangeInformation.EnableDownloadDomains

        $downloadDomainsOrgDisplayValue = "True"
        $downloadDomainsOrgWriteType = "Green"

        if ($downloadDomainsConfig -band 1) {
            $downloadDomainsOrgDisplayValue = "False"
            $downloadDomainsOrgAdditionalDisplayValue = "Download Domains are not configured. You should configure them to be protected against CVE-2021-1730."
            $downloadDomainsOrgWriteType = "Red"
        }

        $downloadDomainsExtDisplayValue = "True"
        $downloadDomainsExtWriteType = "Green"

        if ($downloadDomainsConfig -band 4) {
            $downloadDomainsExtDisplayValue = "False"
            $downloadDomainsExtAdditionalDisplayValue = "Value is set to the same internal or external url as OWA. Please use a different url to reach a protected state against CVE-2021-1730."
            $downloadDomainsExtWriteType = "Red"
        } elseif ($downloadDomainsConfig -band 8) {
            $downloadDomainsExtDisplayValue = "False"
            $downloadDomainsExtAdditionalDisplayValue = "Value not set. Please configure to reach a protected state against CVE-2021-1730."
            $downloadDomainsExtWriteType = "Red"
        }

        $downloadDomainsIntDisplayValue = "True"
        $downloadDomainsIntWriteType = "Green"

        if ($downloadDomainsConfig -band 16) {
            $downloadDomainsIntDisplayValue = "False"
            $downloadDomainsIntAdditionalDisplayValue = "Value is set to the same internal or external url as OWA. Please use a different url to reach a protected state against CVE-2021-1730."
            $downloadDomainsIntWriteType = "Red"
        } elseif ($downloadDomainsConfig -band 32) {
            $downloadDomainsIntDisplayValue = "False"
            $downloadDomainsIntAdditionalDisplayValue = "Value not set. Please configure to reach a protected state against CVE-2021-1730."
            $downloadDomainsIntWriteType = "Red"
        }

        $analyzedResults = Add-AnalyzedResultInformation -Name "Download Domains Enabled" -Details $downloadDomainsOrgDisplayValue `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayWriteType $downloadDomainsOrgWriteType `
            -AddHtmlDetailRow $true `
            -AnalyzedInformation $analyzedResults

        if (![string]::IsNullOrEmpty($downloadDomainsOrgAdditionalDisplayValue)) {
            $analyzedResults = Add-AnalyzedResultInformation -Details $downloadDomainsOrgAdditionalDisplayValue `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayWriteType "Red" `
                -DisplayCustomTabNumber 2 `
                -AddHtmlDetailRow $true `
                -AnalyzedInformation $analyzedResults
        } else {

            $analyzedResults = Add-AnalyzedResultInformation -Name "ExternalDownloadHostName configured correctly" -Details $downloadDomainsExtDisplayValue `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayWriteType $downloadDomainsExtWriteType `
                -DisplayCustomTabNumber 2 `
                -AddHtmlDetailRow $true `
                -AnalyzedInformation $analyzedResults

            if (![string]::IsNullOrEmpty($downloadDomainsExtAdditionalDisplayValue)) {
                $analyzedResults = Add-AnalyzedResultInformation -Details $downloadDomainsExtAdditionalDisplayValue `
                    -DisplayGroupingKey $keySecuritySettings `
                    -DisplayWriteType "Red" `
                    -DisplayCustomTabNumber 2 `
                    -AddHtmlDetailRow $true `
                    -AnalyzedInformation $analyzedResults
            }

            $analyzedResults = Add-AnalyzedResultInformation -Name "InternalDownloadHostName configured correctly" -Details $downloadDomainsIntDisplayValue `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayWriteType $downloadDomainsIntWriteType `
                -DisplayCustomTabNumber 2 `
                -AddHtmlDetailRow $true `
                -AnalyzedInformation $analyzedResults

            if (![string]::IsNullOrEmpty($downloadDomainsIntAdditionalDisplayValue)) {
                $analyzedResults = Add-AnalyzedResultInformation -Details $downloadDomainsIntAdditionalDisplayValue `
                    -DisplayGroupingKey $keySecuritySettings `
                    -DisplayWriteType "Red" `
                    -DisplayCustomTabNumber 2 `
                    -AddHtmlDetailRow $true `
                    -AnalyzedInformation $analyzedResults
            }
        }

        if ($downloadDomainsOrgWriteType -eq "Red" -or
            $downloadDomainsExtWriteType -eq "Red" -or
            $downloadDomainsIntWriteType -eq "Red") {

            $analyzedResults = Add-AnalyzedResultInformation -Details "Configuration instructions: https://aka.ms/HC-DownloadDomains" `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayWriteType "Red" `
                -DisplayCustomTabNumber 2 `
                -AddHtmlDetailRow $true `
                -AnalyzedInformation $analyzedResults

            $Script:AllVulnerabilitiesPassed = $false
        }
    } else {
        Write-Verbose "Download Domains feature not available because we are on: $($exchangeInformation.BuildInformation.MajorVersion) $exchangeCU or on Edge Transport Server"
    }

    #Description: Check for Exchange Emergency Mitigation Service (EEMS)
    #Introduced in: Exchange 2016 CU22, Exchange 2019 CU11

    if (((($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) -and
                ($exchangeCU -ge [HealthChecker.ExchangeCULevel]::CU22)) -or
            (($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) -and
                ($exchangeCU -ge [HealthChecker.ExchangeCULevel]::CU11))) -and
        $exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {

        if (-not([String]::IsNullOrEmpty($exchangeInformation.ExchangeEmergencyMitigationService.MitigationServiceOrgState))) {
            if (($exchangeInformation.ExchangeEmergencyMitigationService.MitigationServiceOrgState) -and
                ($exchangeInformation.ExchangeEmergencyMitigationService.MitigationServiceSrvState)) {
                $eemsWriteType = "Green"
                $eemsOveralState = "Enabled"
            } elseif (($exchangeInformation.ExchangeEmergencyMitigationService.MitigationServiceOrgState -eq $false) -and
                ($exchangeInformation.ExchangeEmergencyMitigationService.MitigationServiceSrvState)) {
                $eemsWriteType = "Yellow"
                $eemsOveralState = "Disabled on org level"
            } elseif (($exchangeInformation.ExchangeEmergencyMitigationService.MitigationServiceSrvState -eq $false) -and
                ($exchangeInformation.ExchangeEmergencyMitigationService.MitigationServiceOrgState)) {
                $eemsWriteType = "Yellow"
                $eemsOveralState = "Disabled on server level"
            } else {
                $eemsWriteType = "Yellow"
                $eemsOveralState = "Disabled"
            }

            $analyzedResults = Add-AnalyzedResultInformation -Name "Exchange Emergency Mitigation Service" -Details $eemsOveralState `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayWriteType $eemsWriteType `
                -AnalyzedInformation $analyzedResults

            $eemsWinSrvWriteType = "Yellow"
            if (-not([String]::IsNullOrEmpty($exchangeInformation.ExchangeEmergencyMitigationService.MitigationWinServiceState))) {
                if ($exchangeInformation.ExchangeEmergencyMitigationService.MitigationWinServiceState -eq "Running") {
                    $eemsWinSrvWriteType = "Grey"
                }
                $details = $exchangeInformation.ExchangeEmergencyMitigationService.MitigationWinServiceState
            } else {
                $details = "Unknown"
            }

            $analyzedResults = Add-AnalyzedResultInformation -Name "Windows service" -Details $details `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType $eemsWinSrvWriteType `
                -AnalyzedInformation $analyzedResults

            if ($exchangeInformation.ExchangeEmergencyMitigationService.MitigationServiceEndpoint -eq 200) {
                $eemsPatternServiceWriteType = "Grey"
                $eemsPatternServiceStatus = ("{0} - Reachable" -f $exchangeInformation.ExchangeEmergencyMitigationService.MitigationServiceEndpoint)
            } else {
                $eemsPatternServiceWriteType = "Yellow"
                $eemsPatternServiceStatus = "Unreachable`r`n`t`tMore information: https://aka.ms/HelpConnectivityEEMS"
            }
            $analyzedResults = Add-AnalyzedResultInformation -Name "Pattern service" -Details $eemsPatternServiceStatus `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayCustomTabNumber 2 `
                -DisplayWriteType $eemsPatternServiceWriteType `
                -AnalyzedInformation $analyzedResults

            if (-not([String]::IsNullOrEmpty($exchangeInformation.ExchangeEmergencyMitigationService.MitigationsApplied))) {
                foreach ($mitigationApplied in $exchangeInformation.ExchangeEmergencyMitigationService.MitigationsApplied) {
                    $analyzedResults = Add-AnalyzedResultInformation -Name "Mitigation applied" -Details $mitigationApplied `
                        -DisplayGroupingKey $keySecuritySettings `
                        -DisplayCustomTabNumber 2 `
                        -AnalyzedInformation $analyzedResults
                }

                $analyzedResults = Add-AnalyzedResultInformation -Details ("Run: 'Get-Mitigations.ps1' from: '{0}' to learn more." -f $exscripts) `
                    -DisplayGroupingKey $keySecuritySettings `
                    -DisplayCustomTabNumber 2 `
                    -AnalyzedInformation $analyzedResults
            }

            if (-not([String]::IsNullOrEmpty($exchangeInformation.ExchangeEmergencyMitigationService.MitigationsBlocked))) {
                foreach ($mitigationBlocked in $exchangeInformation.ExchangeEmergencyMitigationService.MitigationsBlocked) {
                    $analyzedResults = Add-AnalyzedResultInformation -Name "Mitigation blocked" -Details $mitigationBlocked `
                        -DisplayGroupingKey $keySecuritySettings `
                        -DisplayCustomTabNumber 2 `
                        -DisplayWriteType "Yellow" `
                        -AnalyzedInformation $analyzedResults
                }
            }

            if (-not([String]::IsNullOrEmpty($exchangeInformation.ExchangeEmergencyMitigationService.DataCollectionEnabled))) {
                $analyzedResults = Add-AnalyzedResultInformation -Name "Telemetry enabled" -Details $exchangeInformation.ExchangeEmergencyMitigationService.DataCollectionEnabled `
                    -DisplayGroupingKey $keySecuritySettings `
                    -DisplayCustomTabNumber 2 `
                    -AnalyzedInformation $analyzedResults
            }
        } else {
            Write-Verbose "Unable to validate Exchange Emergency Mitigation Service state"
            $analyzedResults = Add-AnalyzedResultInformation -Name "Exchange Emergency Mitigation Service" -Details "Failed to query config" `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayWriteType "Red" `
                -AnalyzedInformation $analyzedResults
        }
    } else {
        Write-Verbose "Exchange Emergency Mitigation Service feature not available because we are on: $($exchangeInformation.BuildInformation.MajorVersion) $exchangeCU or on Edge Transport Server"
    }

    #Description: Check for CVE-2020-0796 SMBv3 vulnerability
    #Affected OS versions: Windows 10 build 1903 and 1909
    #Fix: KB4551762
    #Workaround: Disable SMBv3 compression

    if ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) {
        Write-Verbose "Testing CVE: CVE-2020-0796"
        $buildNumber = $osInformation.BuildInformation.VersionBuild.Split(".")[2]

        if (($buildNumber -eq 18362 -or
                $buildNumber -eq 18363) -and
            ($osInformation.RegistryValues.CurrentVersionUbr -lt 720)) {
            Write-Verbose "Build vulnerable to CVE-2020-0796. Checking if workaround is in place."
            $writeType = "Red"
            $writeValue = "System Vulnerable"

            if ($osInformation.RegistryValues.LanManServerDisabledCompression -eq 1) {
                Write-Verbose "Workaround to disable affected SMBv3 compression is in place."
                $writeType = "Yellow"
                $writeValue = "Workaround is in place"
            } else {
                Write-Verbose "Workaround to disable affected SMBv3 compression is NOT in place."
                $Script:AllVulnerabilitiesPassed = $false
            }

            $details = "{0}`r`n`t`tSee: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/CVE-2020-0796 for more information." -f $writeValue
            $Script:Vulnerabilities += $details
            $analyzedResults = Add-AnalyzedResultInformation -Name "CVE-2020-0796" -Details $details `
                -DisplayGroupingKey $keySecuritySettings `
                -DisplayWriteType $writeType `
                -AddHtmlDetailRow $false `
                -AnalyzedInformation $analyzedResults
        } else {
            Write-Verbose "System NOT vulnerable to CVE-2020-0796. Information URL: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/CVE-2020-0796"
        }
    } else {
        Write-Verbose "Operating System NOT vulnerable to CVE-2020-0796."
    }

    #Description: Check for CVE-2020-1147
    #Affected OS versions: Every OS supporting .NET Core 2.1 and 3.1 and .NET Framework 2.0 SP2 or above
    #Fix: https://portal.msrc.microsoft.com/en-US/security-guidance/advisory/CVE-2020-1147
    #Workaround: N/A
    $dllFileBuildPartToCheckAgainst = 3630

    if ($osInformation.NETFramework.NetMajorVersion -eq [HealthChecker.NetMajorVersion]::Net4d8) {
        $dllFileBuildPartToCheckAgainst = 4190
    }

    $systemDataDll = $osInformation.NETFramework.FileInformation["System.Data.dll"]
    $systemConfigurationDll = $osInformation.NETFramework.FileInformation["System.Configuration.dll"]
    Write-Verbose "System.Data.dll FileBuildPart: $($systemDataDll.VersionInfo.FileBuildPart) | LastWriteTimeUtc: $($systemDataDll.LastWriteTimeUtc)"
    Write-Verbose "System.Configuration.dll FileBuildPart: $($systemConfigurationDll.VersionInfo.FileBuildPart) | LastWriteTimeUtc: $($systemConfigurationDll.LastWriteTimeUtc)"

    if ($systemDataDll.VersionInfo.FileBuildPart -ge $dllFileBuildPartToCheckAgainst -and
        $systemConfigurationDll.VersionInfo.FileBuildPart -ge $dllFileBuildPartToCheckAgainst -and
        $systemDataDll.LastWriteTimeUtc -ge ([System.Convert]::ToDateTime("06/05/2020", [System.Globalization.DateTimeFormatInfo]::InvariantInfo)) -and
        $systemConfigurationDll.LastWriteTimeUtc -ge ([System.Convert]::ToDateTime("06/05/2020", [System.Globalization.DateTimeFormatInfo]::InvariantInfo))) {
        Write-Verbose ("System NOT vulnerable to {0}. Information URL: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/{0}" -f "CVE-2020-1147")
    } else {
        $Script:AllVulnerabilitiesPassed = $false
        $details = "{0}`r`n`t`tSee: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/{0} for more information." -f "CVE-2020-1147"
        $Script:Vulnerabilities += $details
        $Script:AnalyzedInformation = Add-AnalyzedResultInformation -Name "Security Vulnerability" -Details $details `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayWriteType "Red" `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $Script:AnalyzedInformation
    }

    if ($Script:AllVulnerabilitiesPassed) {
        $Script:AnalyzedInformation = Add-AnalyzedResultInformation -Details "All known security issues in this version of the script passed." `
            -DisplayGroupingKey $keySecuritySettings `
            -DisplayWriteType "Green" `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $Script:AnalyzedInformation

        $Script:AnalyzedInformation = Add-AnalyzedResultInformation -Name "Security Vulnerabilities" -Details "None" `
            -AddDisplayResultsLineInfo $false `
            -AddHtmlOverviewValues $true `
            -AnalyzedInformation $Script:AnalyzedInformation
    } else {

        $details = $Script:Vulnerabilities |
            ForEach-Object {
                return $_ + "<br>"
            }

        $Script:AnalyzedInformation = Add-AnalyzedResultInformation -Name "Security Vulnerabilities" -Details $details `
            -AddDisplayResultsLineInfo $false `
            -DisplayWriteType "Red" `
            -AnalyzedInformation $Script:AnalyzedInformation

        $Script:AnalyzedInformation = Add-AnalyzedResultInformation -Name "Vulnerability Detected" -Details $true `
            -AddDisplayResultsLineInfo $false `
            -DisplayWriteType "Red" `
            -AddHtmlOverviewValues $true `
            -AddHtmlDetailRow $false `
            -AnalyzedInformation $Script:AnalyzedInformation
    }

    Write-Debug("End of Analyzer Engine")
    return $Script:AnalyzedInformation
}






Function Get-RemoteRegistrySubKey {
    [CmdletBinding()]
    param(
        [string]$RegistryHive = "LocalMachine",
        [string]$MachineName,
        [string]$SubKey,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Attempting to open the Base Key $RegistryHive on Machine $MachineName"
        $regKey = $null
    }
    process {

        try {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegistryHive, $MachineName)
            Write-Verbose "Attempting to open the Sub Key '$SubKey'"
            $regKey = $reg.OpenSubKey($SubKey)
            Write-Verbose "Opened Sub Key"
        } catch {
            Write-Verbose "Failed to open the registry"

            if ($null -ne $CatchActionFunction) {
                & $CatchActionFunction
            }
        }
    }
    end {
        return $regKey
    }
}

Function Get-RemoteRegistryValue {
    [CmdletBinding()]
    param(
        [string]$RegistryHive = "LocalMachine",
        [string]$MachineName,
        [string]$SubKey,
        [string]$GetValue,
        [string]$ValueType,
        [scriptblock]$CatchActionFunction
    )

    <#
    Valid ValueType return values (case-sensitive)
    (https://docs.microsoft.com/en-us/dotnet/api/microsoft.win32.registryvaluekind?view=net-5.0)
    Binary = REG_BINARY
    DWord = REG_DWORD
    ExpandString = REG_EXPAND_SZ
    MultiString = REG_MULTI_SZ
    None = No data type
    QWord = REG_QWORD
    String = REG_SZ
    Unknown = An unsupported registry data type
    #>

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $registryGetValue = $null
    }
    process {

        try {

            $regSubKey = Get-RemoteRegistrySubKey -RegistryHive $RegistryHive `
                -MachineName $MachineName `
                -SubKey $SubKey

            if (-not ([System.String]::IsNullOrWhiteSpace($regSubKey))) {
                Write-Verbose "Attempting to get the value $GetValue"
                $registryGetValue = $regSubKey.GetValue($GetValue)
                Write-Verbose "Finished running GetValue()"

                if ($null -ne $registryGetValue -and
                    (-not ([System.String]::IsNullOrWhiteSpace($ValueType)))) {
                    Write-Verbose "Validating ValueType $ValueType"
                    $registryValueType = $regSubKey.GetValueKind($GetValue)
                    Write-Verbose "Finished running GetValueKind()"

                    if ($ValueType -ne $registryValueType) {
                        Write-Verbose "ValueType: $ValueType is different to the returned ValueType: $registryValueType"
                        $registryGetValue = $null
                    } else {
                        Write-Verbose "ValueType matches: $ValueType"
                    }
                }
            }
        } catch {
            Write-Verbose "Failed to get the value on the registry"

            if ($null -ne $CatchActionFunction) {
                & $CatchActionFunction
            }
        }
    }
    end {
        Write-Verbose "Get-RemoteRegistryValue Return Value: '$registryGetValue'"
        return $registryGetValue
    }
}


Function Invoke-CatchActionError {
    [CmdletBinding()]
    param(
        [scriptblock]$CatchActionFunction
    )

    if ($null -ne $CatchActionFunction) {
        & $CatchActionFunction
    }
}

Function Invoke-ScriptBlockHandler {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $ComputerName,

        [Parameter(Mandatory = $true)]
        [scriptblock]
        $ScriptBlock,

        [string]
        $ScriptBlockDescription,

        [object]
        $ArgumentList,

        [bool]
        $IncludeNoProxyServerOption,

        [scriptblock]
        $CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $returnValue = $null
    }
    process {

        if (-not([string]::IsNullOrEmpty($ScriptBlockDescription))) {
            Write-Verbose "Description: $ScriptBlockDescription"
        }

        try {

            if (($ComputerName).Split(".")[0] -ne $env:COMPUTERNAME) {

                $params = @{
                    ComputerName = $ComputerName
                    ScriptBlock  = $ScriptBlock
                    ErrorAction  = "Stop"
                }

                if ($IncludeNoProxyServerOption) {
                    Write-Verbose "Including SessionOption"
                    $params.Add("SessionOption", (New-PSSessionOption -ProxyAccessType NoProxyServer))
                }

                if ($null -ne $ArgumentList) {
                    Write-Verbose "Running Invoke-Command with argument list"
                    $params.Add("ArgumentList", $ArgumentList)
                } else {
                    Write-Verbose "Running Invoke-Command without argument list"
                }

                $returnValue = Invoke-Command @params
            } else {

                if ($null -ne $ArgumentList) {
                    Write-Verbose "Running Script Block Locally with argument list"
                    $returnValue = & $ScriptBlock $ArgumentList
                } else {
                    Write-Verbose "Running Script Block Locally without argument list"
                    $returnValue = & $ScriptBlock
                }
            }
        } catch {
            Write-Verbose "Failed to run $($MyInvocation.MyCommand)"
            Invoke-CatchActionError $CatchActionFunction
        }
    }
    end {
        Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
        return $returnValue
    }
}

Function Invoke-CatchActions {
    param(
        [object]$CopyThisError = $Error[0]
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $Script:ErrorsExcludedCount++
    $Script:ErrorsExcluded += $CopyThisError
    Write-Verbose "Error Excluded Count: $Script:ErrorsExcludedCount"
    Write-Verbose "Error Count: $($Error.Count)"
    Write-Verbose $CopyThisError

    if ($null -ne $CopyThisError.ScriptStackTrace) {
        Write-Verbose $CopyThisError.ScriptStackTrace
    }
}

Function Get-ExchangeAdSchemaClass {
    param(
        [Parameter(Mandatory = $true)][string]$SchemaClassName
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand) to query $SchemaClassName schema class"

    $rootDSE = [ADSI]("LDAP://RootDSE")

    if ([string]::IsNullOrEmpty($rootDSE.schemaNamingContext)) {
        return $null
    }

    $directorySearcher = New-Object System.DirectoryServices.DirectorySearcher
    $directorySearcher.SearchScope = "Subtree"
    $directorySearcher.SearchRoot = [ADSI]("LDAP://" + $rootDSE.schemaNamingContext.ToString())
    $directorySearcher.Filter = "(Name={0})" -f $SchemaClassName

    $findAll = $directorySearcher.FindAll()

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $findAll
}

Function Get-ExchangeApplicationConfigurationFileValidation {
    param(
        [string[]]$ConfigFileLocation
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $results = @{}
    $ConfigFileLocation |
        ForEach-Object {
            $obj = Invoke-ScriptBlockHandler -ComputerName $Script:Server -ScriptBlockDescription "Getting Exchange Application Configuration File Validation" `
                -CatchActionFunction ${Function:Invoke-CatchActions} `
                -ScriptBlock {
                param($Location)
                return [PSCustomObject]@{
                    Present  = ((Test-Path $Location))
                    FileName = ([IO.Path]::GetFileName($Location))
                    FilePath = $Location
                }
            } -ArgumentList $_
            $results.Add($obj.FileName, $obj)
        }
    return $results
}


function Get-AppPool {
    [CmdletBinding()]
    param ()

    begin {
        function Get-IndentLevel ($line) {
            if ($line.StartsWith(" ")) {
                ($line | Select-String "^ +").Matches[0].Length
            } else {
                0
            }
        }

        function Convert-FromAppPoolText {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory = $true)]
                [string[]]
                $Text,

                [Parameter(Mandatory = $false)]
                [int]
                $Line = 0,

                [Parameter(Mandatory = $false)]
                [int]
                $MinimumIndentLevel = 2
            )

            if ($Line -ge $Text.Count) {
                return $null
            }

            $startingIndentLevel = Get-IndentLevel $Text[$Line]
            if ($startingIndentLevel -lt $MinimumIndentLevel) {
                return $null
            }

            $hash = @{}

            while ($Line -lt $Text.Count) {
                $indentLevel = Get-IndentLevel $Text[$Line]
                if ($indentLevel -gt $startingIndentLevel) {
                    # Skip until we get to the next thing at this level
                } elseif ($indentLevel -eq $startingIndentLevel) {
                    # We have a property at this level. Add it to the object.
                    if ($Text[$Line] -match "\[(\S+)\]") {
                        $name = $Matches[1]
                        $value = Convert-FromAppPoolText -Text $Text -Line ($Line + 1) -MinimumIndentLevel $startingIndentLevel
                        $hash[$name] = $value
                    } elseif ($Text[$Line] -match "\s+(\S+):`"(.*)`"") {
                        $name = $Matches[1]
                        $value = $Matches[2].Trim("`"")
                        $hash[$name] = $value
                    }
                } else {
                    # IndentLevel is less than what we started with, so return
                    [PSCustomObject]$hash
                    return
                }

                ++$Line
            }

            [PSCustomObject]$hash
        }

        $appPoolCmd = "$env:windir\System32\inetsrv\appcmd.exe"
    }

    process {
        $appPoolNames = & $appPoolCmd list apppool |
            Select-String "APPPOOL `"(\S+)`" " |
            ForEach-Object { $_.Matches.Groups[1].Value }

        foreach ($appPoolName in $appPoolNames) {
            $appPoolText = & $appPoolCmd list apppool $appPoolName /text:*
            Convert-FromAppPoolText -Text $appPoolText -Line 1
        }
    }
}
Function Get-ExchangeAppPoolsInformation {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $appPool = Invoke-ScriptBlockHandler -ComputerName $Script:Server -ScriptBlock ${Function:Get-AppPool} `
        -ScriptBlockDescription "Getting App Pool information" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    $exchangeAppPoolsInfo = @{}

    $appPool |
        Where-Object { $_.add.name -like "MSExchange*" } |
        ForEach-Object {
            $configContent = Invoke-ScriptBlockHandler -ComputerName $Script:Server -ScriptBlock {
                param(
                    $FilePath
                )
                if (Test-Path $FilePath) {
                    return (Get-Content $FilePath)
                }
                return [string]::Empty
            } `
                -ScriptBlockDescription "Getting Content file for $($_.add.name)" `
                -ArgumentList $_.add.CLRConfigFile `
                -CatchActionFunction ${Function:Invoke-CatchActions}

            $gcUnknown = $true
            $gcServerEnabled = $false

            if (-not ([string]::IsNullOrEmpty($configContent))) {
                $gcSetting = ([xml]$configContent).Configuration.Runtime.gcServer.Enabled
                $gcUnknown = $gcSetting -ne "true" -and $gcSetting -ne "false"
                $gcServerEnabled = $gcSetting -eq "true"
            }
            $exchangeAppPoolsInfo.Add($_.add.Name, [PSCustomObject]@{
                    ConfigContent   = $configContent
                    AppSettings     = $_
                    GCUnknown       = $gcUnknown
                    GCServerEnabled = $gcServerEnabled
                })
        }

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $exchangeAppPoolsInfo
}

Function Get-ExchangeBuildVersionInformation {
    [CmdletBinding()]
    param(
        [object]$AdminDisplayVersion
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed $($AdminDisplayVersion.ToString())"
        $AdminDisplayVersion = $AdminDisplayVersion.ToString()
        $exchangeMajorVersion = [string]::Empty
        [int]$major = 0
        [int]$minor = 0
        [int]$build = 0
        [int]$revision = 0
        $product = $null
        [double]$buildVersion = 0.0
    }
    process {
        $split = $AdminDisplayVersion.Substring(($AdminDisplayVersion.IndexOf(" ")) + 1, 4).Split(".")
        $major = [int]$split[0]
        $minor = [int]$split[1]
        $product = $major + ($minor / 10)

        $buildStart = $AdminDisplayVersion.LastIndexOf(" ") + 1
        $split = $AdminDisplayVersion.Substring($buildStart, ($AdminDisplayVersion.LastIndexOf(")") - $buildStart)).Split(".")
        $build = [int]$split[0]
        $revision = [int]$split[1]
        $revisionDecimal = if ($revision -lt 10) { $revision / 10 } else { $revision / 100 }
        $buildVersion = $build + $revisionDecimal

        Write-Verbose "Determining Major Version based off of $product"

        switch ([string]$product) {
            "14.3" { $exchangeMajorVersion = "Exchange2010" }
            "15" { $exchangeMajorVersion = "Exchange2013" }
            "15.1" { $exchangeMajorVersion = "Exchange2016" }
            "15.2" { $exchangeMajorVersion = "Exchange2019" }
            default { $exchangeMajorVersion = "Unknown" }
        }
    }
    end {
        Write-Verbose "Found Major Version '$exchangeMajorVersion'"
        return [PSCustomObject]@{
            MajorVersion = $exchangeMajorVersion
            Major        = $major
            Minor        = $minor
            Build        = $build
            Revision     = $revision
            Product      = $product
            BuildVersion = $buildVersion
        }
    }
}

Function Get-ExchangeEmergencyMitigationServiceState {
    [CmdletBinding()]
    [OutputType("System.Object")]
    param(
        [Parameter(Mandatory = $true)]
        [object]
        $RequiredInformation,
        [Parameter(Mandatory = $false)]
        [scriptblock]
        $CatchActionFunction
    )
    begin {
        $computerName = $RequiredInformation.ComputerName
        $emergencyMitigationServiceOrgState = $RequiredInformation.MitigationsEnabled
        $exchangeServerConfiguration = $RequiredInformation.GetExchangeServer
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - Computername: $ComputerName"
    }
    process {
        if ($null -ne $emergencyMitigationServiceOrgState) {
            Write-Verbose "Exchange Emergency Mitigation Service detected"
            try {
                $exchangeEmergencyMitigationWinServiceRating = $null
                $emergencyMitigationWinService = Get-Service -ComputerName $ComputerName -Name MSExchangeMitigation
                if (($emergencyMitigationWinService.Status -eq "Running") -and
                    ($emergencyMitigationWinService.StartType -eq "Automatic")) {
                    $exchangeEmergencyMitigationWinServiceRating = "Running"
                } else {
                    $exchangeEmergencyMitigationWinServiceRating = "Investigate"
                }
            } catch {
                Write-Verbose "Failed to query EEMS Windows service data"
                Invoke-CatchActionError $CatchActionFunction
            }

            $eemsEndpoint = Invoke-ScriptBlockHandler -ComputerName $ComputerName -ScriptBlockDescription "Test EEMS pattern service connectivity" `
                -CatchActionFunction $CatchActionFunction `
                -ScriptBlock {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; `
                    Invoke-WebRequest -Method Get -Uri "https://officeclient.microsoft.com/getexchangemitigations" -UseBasicParsing
            }
        }
    }
    end {
        return [PSCustomObject]@{
            MitigationWinServiceState = $exchangeEmergencyMitigationWinServiceRating
            MitigationServiceOrgState = $emergencyMitigationServiceOrgState
            MitigationServiceSrvState = $exchangeServerConfiguration.MitigationsEnabled
            MitigationServiceEndpoint = $eemsEndpoint.StatusCode
            MitigationsApplied        = $exchangeServerConfiguration.MitigationsApplied
            MitigationsBlocked        = $exchangeServerConfiguration.MitigationsBlocked
            DataCollectionEnabled     = $exchangeServerConfiguration.DataCollectionEnabled
        }
    }
}

Function Get-ExchangeServerCertificates {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    try {
        Write-Verbose "Trying to receive certificates from Exchange server: $($Script:Server)"
        $exchangeServerCertificates = Get-ExchangeCertificate -Server $Script:Server -ErrorAction Stop

        if ($null -ne $exchangeServerCertificates) {
            try {
                $authConfig = Get-AuthConfig -ErrorAction Stop
                $authConfigDetected = $true
            } catch {
                $authConfigDetected = $false
                Invoke-CatchActions
            }

            [array]$certObject = @()
            foreach ($cert in $exchangeServerCertificates) {
                try {
                    $certificateLifetime = ([System.Convert]::ToDateTime($cert.NotAfter, [System.Globalization.DateTimeFormatInfo]::InvariantInfo) - (Get-Date)).Days
                    $sanCertificateInfo = $false

                    $currentErrors = $Error.Count
                    if ($null -ne $cert.DnsNameList -and
                        ($cert.DnsNameList).Count -gt 1) {
                        $sanCertificateInfo = $true
                        $certDnsNameList = $cert.DnsNameList
                    } elseif ($null -eq $cert.DnsNameList) {
                        $certDnsNameList = "None"
                    } else {
                        $certDnsNameList = $cert.DnsNameList
                    }
                    if ($currentErrors -lt $Error.Count) {
                        $i = 0
                        while ($i -lt ($Error.Count - $currentErrors)) {
                            Invoke-CatchActions $Error[$i]
                            $i++
                        }
                    }

                    if ($authConfigDetected) {
                        $isAuthConfigInfo = $false

                        if ($cert.Thumbprint -eq $authConfig.CurrentCertificateThumbprint) {
                            $isAuthConfigInfo = $true
                        }
                    } else {
                        $isAuthConfigInfo = "InvalidAuthConfig"
                    }

                    if ([String]::IsNullOrEmpty($cert.FriendlyName)) {
                        $certFriendlyName = ($certDnsNameList[0]).ToString()
                    } else {
                        $certFriendlyName = $cert.FriendlyName
                    }

                    if ([String]::IsNullOrEmpty($cert.Status)) {
                        $certStatus = "Unknown"
                    } else {
                        $certStatus = ($cert.Status).ToString()
                    }

                    if ([String]::IsNullOrEmpty($cert.SignatureAlgorithm.FriendlyName)) {
                        $certSignatureAlgorithm = "Unknown"
                        $certSignatureHashAlgorithm = "Unknown"
                        $certSignatureHashAlgorithmSecure = 0
                    } else {
                        $certSignatureAlgorithm = $cert.SignatureAlgorithm.FriendlyName
                        <#
                            OID Table
                            https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-gpnap/a48b02b2-2a10-4eb0-bed4-1807a6d2f5ad
                            SignatureHashAlgorithmSecure = Unknown 0
                            SignatureHashAlgorithmSecure = Insecure/Weak 1
                            SignatureHashAlgorithmSecure = Secure 2
                        #>
                        switch ($cert.SignatureAlgorithm.Value) {
                            "1.2.840.113549.1.1.5" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.2.840.113549.1.1.4" { $certSignatureHashAlgorithm = "md5"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.2.840.10040.4.3" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.29" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.15" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.3" { $certSignatureHashAlgorithm = "md5"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.2.840.113549.1.1.2" { $certSignatureHashAlgorithm = "md2"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.2.840.113549.1.1.3" { $certSignatureHashAlgorithm = "md4"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.2" { $certSignatureHashAlgorithm = "md4"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.4" { $certSignatureHashAlgorithm = "md4"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.7.2.3.1" { $certSignatureHashAlgorithm = "md2"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.13" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.27" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "2.16.840.1.101.2.1.1.19" { $certSignatureHashAlgorithm = "mosaicSignature"; $certSignatureHashAlgorithmSecure = 0 }
                            "1.3.14.3.2.26" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.2.840.113549.2.5" { $certSignatureHashAlgorithm = "md5"; $certSignatureHashAlgorithmSecure = 1 }
                            "2.16.840.1.101.3.4.2.1" { $certSignatureHashAlgorithm = "sha256"; $certSignatureHashAlgorithmSecure = 2 }
                            "2.16.840.1.101.3.4.2.2" { $certSignatureHashAlgorithm = "sha384"; $certSignatureHashAlgorithmSecure = 2 }
                            "2.16.840.1.101.3.4.2.3" { $certSignatureHashAlgorithm = "sha512"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.113549.1.1.11" { $certSignatureHashAlgorithm = "sha256"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.113549.1.1.12" { $certSignatureHashAlgorithm = "sha384"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.113549.1.1.13" { $certSignatureHashAlgorithm = "sha512"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.113549.1.1.10" { $certSignatureHashAlgorithm = "rsassa-pss"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.10045.4.1" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.2.840.10045.4.3.2" { $certSignatureHashAlgorithm = "sha256"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.10045.4.3.3" { $certSignatureHashAlgorithm = "sha384"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.10045.4.3.4" { $certSignatureHashAlgorithm = "sha512"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.10045.4.3" { $certSignatureHashAlgorithm = "sha256"; $certSignatureHashAlgorithmSecure = 2 }
                            Default { $certSignatureHashAlgorithm = "Unknown"; $certSignatureHashAlgorithmSecure = 0 }
                        }
                    }

                    $certInformationObj = New-Object PSCustomObject
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "FriendlyName" -Value $certFriendlyName
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "Thumbprint" -Value $cert.Thumbprint
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "PublicKeySize" -Value $cert.PublicKey.Key.KeySize
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "SignatureAlgorithm" -Value $certSignatureAlgorithm
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "SignatureHashAlgorithm" -Value $certSignatureHashAlgorithm
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "SignatureHashAlgorithmSecure" -Value $certSignatureHashAlgorithmSecure
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "IsSanCertificate" -Value $sanCertificateInfo
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "Namespaces" -Value $certDnsNameList
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "Services" -Value $cert.Services
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "IsCurrentAuthConfigCertificate" -Value $isAuthConfigInfo
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "LifetimeInDays" -Value $certificateLifetime
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "Status" -Value $certStatus
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "CertificateObject" -Value $cert

                    $certObject += $certInformationObj
                } catch {
                    Write-Verbose "Unable to process certificate: $($cert.Thumbprint)"
                    Invoke-CatchActions
                }
            }
            Write-Verbose "Processed: $($certObject.Count) certificates"
            return $certObject
        } else {
            Write-Verbose "Failed to find any Exchange certificates"
            return $null
        }
    } catch {
        Write-Verbose "Failed to run Get-ExchangeCertificate. Error: $($Error[0].Exception)."
        Invoke-CatchActions
    }
}

Function Get-ExchangeServerMaintenanceState {
    param(
        [Parameter(Mandatory = $false)][array]$ComponentsToSkip
    )
    Write-Verbose "Calling Function: $($MyInvocation.MyCommand)"

    [HealthChecker.ExchangeServerMaintenance]$serverMaintenance = New-Object -TypeName HealthChecker.ExchangeServerMaintenance
    $serverMaintenance.GetServerComponentState = Get-ServerComponentState -Identity $Script:Server -ErrorAction SilentlyContinue

    try {
        $serverMaintenance.GetMailboxServer = Get-MailboxServer -Identity $Script:Server -ErrorAction SilentlyContinue
    } catch {
        Write-Verbose "Failed to run Get-MailboxServer"
        Invoke-CatchActions
    }

    try {
        $serverMaintenance.GetClusterNode = Get-ClusterNode -Name $Script:Server -ErrorAction Stop
    } catch {
        Write-Verbose "Failed to run Get-ClusterNode"
        Invoke-CatchActions
    }

    Write-Verbose "Running ServerComponentStates checks"

    foreach ($component in $serverMaintenance.GetServerComponentState) {
        if (($null -ne $ComponentsToSkip -and
                $ComponentsToSkip.Count -ne 0) -and
            $ComponentsToSkip -notcontains $component.Component) {
            if ($component.State -ne "Active") {
                $latestLocalState = $null
                $latestRemoteState = $null

                if ($null -ne $component.LocalStates -and
                    $component.LocalStates.Count -gt 0) {
                    $latestLocalState = ($component.LocalStates | Sort-Object { $_.TimeStamp } -ErrorAction SilentlyContinue)[-1]
                }

                if ($null -ne $component.RemoteStates -and
                    $component.RemoteStates.Count -gt 0) {
                    $latestRemoteState = ($component.RemoteStates | Sort-Object { $_.TimeStamp } -ErrorAction SilentlyContinue)[-1]
                }

                Write-Verbose "Component: '$($component.Component)' LocalState: '$($latestLocalState.State)' RemoteState: '$($latestRemoteState.State)'"

                if ($latestLocalState.State -eq $latestRemoteState.State) {
                    $serverMaintenance.InactiveComponents += "'{0}' is in Maintenance Mode" -f $component.Component
                } else {
                    if (($null -ne $latestLocalState) -and
                        ($latestLocalState.State -ne "Active")) {
                        $serverMaintenance.InactiveComponents += "'{0}' is in Local Maintenance Mode only" -f $component.Component
                    }

                    if (($null -ne $latestRemoteState) -and
                        ($latestRemoteState.State -ne "Active")) {
                        $serverMaintenance.InactiveComponents += "'{0}' is in Remote Maintenance Mode only" -f $component.Component
                    }
                }
            } else {
                Write-Verbose "Component '$($component.Component)' is Active"
            }
        } else {
            Write-Verbose "Component: $($component.Component) will be skipped"
        }
    }

    return $serverMaintenance
}

Function Get-ExchangeUpdates {
    param(
        [Parameter(Mandatory = $true)][HealthChecker.ExchangeMajorVersion]$ExchangeMajorVersion
    )
    Write-Verbose("Calling: $($MyInvocation.MyCommand) Passed: $ExchangeMajorVersion")
    $RegLocation = [string]::Empty

    if ([HealthChecker.ExchangeMajorVersion]::Exchange2013 -eq $ExchangeMajorVersion) {
        $RegLocation = "SOFTWARE\Microsoft\Updates\Exchange 2013"
    } elseif ([HealthChecker.ExchangeMajorVersion]::Exchange2016 -eq $ExchangeMajorVersion) {
        $RegLocation = "SOFTWARE\Microsoft\Updates\Exchange 2016"
    } else {
        $RegLocation = "SOFTWARE\Microsoft\Updates\Exchange 2019"
    }

    $RegKey = Get-RemoteRegistrySubKey -MachineName $Script:Server `
        -SubKey $RegLocation `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    if ($null -ne $RegKey) {
        $IU = $RegKey.GetSubKeyNames()
        if ($null -ne $IU) {
            Write-Verbose "Detected fixes installed on the server"
            $fixes = @()
            foreach ($key in $IU) {
                $IUKey = $RegKey.OpenSubKey($key)
                $IUName = $IUKey.GetValue("PackageName")
                Write-Verbose "Found: $IUName"
                $fixes += $IUName
            }
            return $fixes
        } else {
            Write-Verbose "No IUs found in the registry"
        }
    } else {
        Write-Verbose "No RegKey returned"
    }

    Write-Verbose "Exiting: Get-ExchangeUpdates"
    return $null
}

Function Get-ExSetupDetails {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exSetupDetails = [string]::Empty
    Function Get-ExSetupDetailsScriptBlock {
        Get-Command ExSetup | ForEach-Object { $_.FileVersionInfo }
    }

    $exSetupDetails = Invoke-ScriptBlockHandler -ComputerName $Script:Server -ScriptBlock ${Function:Get-ExSetupDetailsScriptBlock} -ScriptBlockDescription "Getting ExSetup remotely" -CatchActionFunction ${Function:Invoke-CatchActions}
    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $exSetupDetails
}

Function Get-ServerRole {
    param(
        [Parameter(Mandatory = $true)][object]$ExchangeServerObj
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $roles = $ExchangeServerObj.ServerRole.ToString()
    Write-Verbose "Roll: $roles"
    #Need to change this to like because of Exchange 2010 with AIO with the hub role.
    if ($roles -like "Mailbox, ClientAccess*") {
        return [HealthChecker.ExchangeServerRole]::MultiRole
    } elseif ($roles -eq "Mailbox") {
        return [HealthChecker.ExchangeServerRole]::Mailbox
    } elseif ($roles -eq "Edge") {
        return [HealthChecker.ExchangeServerRole]::Edge
    } elseif ($roles -like "*ClientAccess*") {
        return [HealthChecker.ExchangeServerRole]::ClientAccess
    } else {
        return [HealthChecker.ExchangeServerRole]::None
    }
}
Function Get-ExchangeInformation {
    param(
        [HealthChecker.OSServerVersion]$OSMajorVersion
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand) Passed: OSMajorVersion: $OSMajorVersion"
    [HealthChecker.ExchangeInformation]$exchangeInformation = New-Object -TypeName HealthChecker.ExchangeInformation
    $exchangeInformation.GetExchangeServer = (Get-ExchangeServer -Identity $Script:Server -Status)
    $exchangeInformation.ExchangeCertificates = Get-ExchangeServerCertificates
    $buildInformation = $exchangeInformation.BuildInformation
    $buildVersionInfo = Get-ExchangeBuildVersionInformation -AdminDisplayVersion $exchangeInformation.GetExchangeServer.AdminDisplayVersion
    $buildInformation.MajorVersion = ([HealthChecker.ExchangeMajorVersion]$buildVersionInfo.MajorVersion)
    $buildInformation.BuildNumber = "{0}.{1}.{2}.{3}" -f $buildVersionInfo.Major, $buildVersionInfo.Minor, $buildVersionInfo.Build, $buildVersionInfo.Revision
    $buildInformation.ServerRole = (Get-ServerRole -ExchangeServerObj $exchangeInformation.GetExchangeServer)
    $buildInformation.ExchangeSetup = Get-ExSetupDetails

    if ($buildInformation.ServerRole -le [HealthChecker.ExchangeServerRole]::Mailbox ) {
        $exchangeInformation.GetMailboxServer = (Get-MailboxServer -Identity $Script:Server)
    }

    if (($buildInformation.MajorVersion -ge [HealthChecker.ExchangeMajorVersion]::Exchange2016 -and
            $buildInformation.ServerRole -le [HealthChecker.ExchangeServerRole]::Mailbox) -or
        ($buildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013 -and
            ($buildInformation.ServerRole -eq [HealthChecker.ExchangeServerRole]::ClientAccess -or
        $buildInformation.ServerRole -eq [HealthChecker.ExchangeServerRole]::MultiRole))) {
        $exchangeInformation.GetOwaVirtualDirectory = Get-OwaVirtualDirectory -Identity ("{0}\owa (Default Web Site)" -f $Script:Server) -ADPropertiesOnly
        $exchangeInformation.GetWebServicesVirtualDirectory = Get-WebServicesVirtualDirectory -Server $Script:Server
    }

    if ($Script:ExchangeShellComputer.ToolsOnly) {
        $buildInformation.LocalBuildNumber = "{0}.{1}.{2}.{3}" -f $Script:ExchangeShellComputer.Major, $Script:ExchangeShellComputer.Minor, `
            $Script:ExchangeShellComputer.Build, `
            $Script:ExchangeShellComputer.Revision
    }

    #Exchange 2013 or greater
    if ($buildInformation.MajorVersion -ge [HealthChecker.ExchangeMajorVersion]::Exchange2013) {
        $netFrameworkExchange = $exchangeInformation.NETFramework
        $buildAndRevision = $buildVersionInfo.BuildVersion
        Write-Verbose "The build and revision number: $buildAndRevision"
        #Build Numbers: https://docs.microsoft.com/en-us/Exchange/new-features/build-numbers-and-release-dates?view=exchserver-2019
        if ($buildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) {
            Write-Verbose "Exchange 2019 is detected. Checking build number..."
            $buildInformation.FriendlyName = "Exchange 2019 "

            #Exchange 2019 Information
            if ($buildAndRevision -lt 330.5) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::RTM
                $buildInformation.FriendlyName += "RTM"
                $buildInformation.ReleaseDate = "10/22/2018"
            } elseif ($buildAndRevision -lt 397.3) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU1
                $buildInformation.FriendlyName += "CU1"
                $buildInformation.ReleaseDate = "02/12/2019"
            } elseif ($buildAndRevision -lt 464.5) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU2
                $buildInformation.FriendlyName += "CU2"
                $buildInformation.ReleaseDate = "06/18/2019"
            } elseif ($buildAndRevision -lt 529.5) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU3
                $buildInformation.FriendlyName += "CU3"
                $buildInformation.ReleaseDate = "09/17/2019"
            } elseif ($buildAndRevision -lt 595.3) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU4
                $buildInformation.FriendlyName += "CU4"
                $buildInformation.ReleaseDate = "12/17/2019"
            } elseif ($buildAndRevision -lt 659.4) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU5
                $buildInformation.FriendlyName += "CU5"
                $buildInformation.ReleaseDate = "03/17/2020"
            } elseif ($buildAndRevision -lt 721.2) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU6
                $buildInformation.FriendlyName += "CU6"
                $buildInformation.ReleaseDate = "06/16/2020"
            } elseif ($buildAndRevision -lt 792.3) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU7
                $buildInformation.FriendlyName += "CU7"
                $buildInformation.ReleaseDate = "09/15/2020"
            } elseif ($buildAndRevision -lt 858.5) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU8
                $buildInformation.FriendlyName += "CU8"
                $buildInformation.ReleaseDate = "12/15/2020"
            } elseif ($buildAndRevision -lt 922.7) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU9
                $buildInformation.FriendlyName += "CU9"
                $buildInformation.ReleaseDate = "03/16/2021"
            } elseif ($buildAndRevision -lt 986.5) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU10
                $buildInformation.FriendlyName += "CU10"
                $buildInformation.ReleaseDate = "06/29/2021"
                $buildInformation.SupportedBuild = $true
            } elseif ($buildAndRevision -ge 986.5) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU11
                $buildInformation.FriendlyName += "CU11"
                $buildInformation.ReleaseDate = "09/28/2021"
                $buildInformation.SupportedBuild = $true
            }

            #Exchange 2019 .NET Information
            if ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU2) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU4) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
            } else {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
            }
        } elseif ($buildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) {
            Write-Verbose "Exchange 2016 is detected. Checking build number..."
            $buildInformation.FriendlyName = "Exchange 2016 "

            #Exchange 2016 Information
            if ($buildAndRevision -lt 466.34) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU1
                $buildInformation.FriendlyName += "CU1"
                $buildInformation.ReleaseDate = "03/15/2016"
            } elseif ($buildAndRevision -lt 544.27) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU2
                $buildInformation.FriendlyName += "CU2"
                $buildInformation.ReleaseDate = "06/21/2016"
            } elseif ($buildAndRevision -lt 669.32) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU3
                $buildInformation.FriendlyName += "CU3"
                $buildInformation.ReleaseDate = "09/20/2016"
            } elseif ($buildAndRevision -lt 845.34) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU4
                $buildInformation.FriendlyName += "CU4"
                $buildInformation.ReleaseDate = "12/13/2016"
            } elseif ($buildAndRevision -lt 1034.26) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU5
                $buildInformation.FriendlyName += "CU5"
                $buildInformation.ReleaseDate = "03/21/2017"
            } elseif ($buildAndRevision -lt 1261.35) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU6
                $buildInformation.FriendlyName += "CU6"
                $buildInformation.ReleaseDate = "06/24/2017"
            } elseif ($buildAndRevision -lt 1415.2) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU7
                $buildInformation.FriendlyName += "CU7"
                $buildInformation.ReleaseDate = "09/16/2017"
            } elseif ($buildAndRevision -lt 1466.3) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU8
                $buildInformation.FriendlyName += "CU8"
                $buildInformation.ReleaseDate = "12/19/2017"
            } elseif ($buildAndRevision -lt 1531.3) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU9
                $buildInformation.FriendlyName += "CU9"
                $buildInformation.ReleaseDate = "03/20/2018"
            } elseif ($buildAndRevision -lt 1591.10) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU10
                $buildInformation.FriendlyName += "CU10"
                $buildInformation.ReleaseDate = "06/19/2018"
            } elseif ($buildAndRevision -lt 1713.5) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU11
                $buildInformation.FriendlyName += "CU11"
                $buildInformation.ReleaseDate = "10/16/2018"
            } elseif ($buildAndRevision -lt 1779.2) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU12
                $buildInformation.FriendlyName += "CU12"
                $buildInformation.ReleaseDate = "02/12/2019"
            } elseif ($buildAndRevision -lt 1847.3) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU13
                $buildInformation.FriendlyName += "CU13"
                $buildInformation.ReleaseDate = "06/18/2019"
            } elseif ($buildAndRevision -lt 1913.5) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU14
                $buildInformation.FriendlyName += "CU14"
                $buildInformation.ReleaseDate = "09/17/2019"
            } elseif ($buildAndRevision -lt 1979.3) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU15
                $buildInformation.FriendlyName += "CU15"
                $buildInformation.ReleaseDate = "12/17/2019"
            } elseif ($buildAndRevision -lt 2044.4) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU16
                $buildInformation.FriendlyName += "CU16"
                $buildInformation.ReleaseDate = "03/17/2020"
            } elseif ($buildAndRevision -lt 2106.2) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU17
                $buildInformation.FriendlyName += "CU17"
                $buildInformation.ReleaseDate = "06/16/2020"
            } elseif ($buildAndRevision -lt 2176.2) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU18
                $buildInformation.FriendlyName += "CU18"
                $buildInformation.ReleaseDate = "09/15/2020"
            } elseif ($buildAndRevision -lt 2242.4) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU19
                $buildInformation.FriendlyName += "CU19"
                $buildInformation.ReleaseDate = "12/15/2020"
            } elseif ($buildAndRevision -lt 2308.8) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU20
                $buildInformation.FriendlyName += "CU20"
                $buildInformation.ReleaseDate = "03/16/2021"
            } elseif ($buildAndRevision -lt 2375.7) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU21
                $buildInformation.FriendlyName += "CU21"
                $buildInformation.ReleaseDate = "06/29/2021"
                $buildInformation.SupportedBuild = $true
            } elseif ($buildAndRevision -ge 2375.7) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU22
                $buildInformation.FriendlyName += "CU22"
                $buildInformation.ReleaseDate = "09/28/2021"
                $buildInformation.SupportedBuild = $true
            }

            #Exchange 2016 .NET Information
            if ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU2) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
            } elseif ($buildInformation.CU -eq [HealthChecker.ExchangeCULevel]::CU2) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d1wFix
            } elseif ($buildInformation.CU -eq [HealthChecker.ExchangeCULevel]::CU3) {

                if ($OSMajorVersion -eq [HealthChecker.OSServerVersion]::Windows2016) {
                    $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
                    $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
                } else {
                    $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
                    $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d1wFix
                }
            } elseif ($buildInformation.CU -eq [HealthChecker.ExchangeCULevel]::CU4) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU8) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU10) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d1
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU11) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d1
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d1
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU13) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d1
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU15) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
            } else {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
            }
        } else {
            Write-Verbose "Exchange 2013 is detected. Checking build number..."
            $buildInformation.FriendlyName = "Exchange 2013"

            #Exchange 2013 Information
            if ($buildAndRevision -lt 712.24) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU1
                $buildInformation.FriendlyName += "CU1"
                $buildInformation.ReleaseDate = "04/02/2013"
            } elseif ($buildAndRevision -lt 775.38) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU2
                $buildInformation.FriendlyName += "CU2"
                $buildInformation.ReleaseDate = "07/09/2013"
            } elseif ($buildAndRevision -lt 847.32) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU3
                $buildInformation.FriendlyName += "CU3"
                $buildInformation.ReleaseDate = "11/25/2013"
            } elseif ($buildAndRevision -lt 913.22) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU4
                $buildInformation.FriendlyName += "CU4"
                $buildInformation.ReleaseDate = "02/25/2014"
            } elseif ($buildAndRevision -lt 995.29) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU5
                $buildInformation.FriendlyName += "CU5"
                $buildInformation.ReleaseDate = "05/27/2014"
            } elseif ($buildAndRevision -lt 1044.25) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU6
                $buildInformation.FriendlyName += "CU6"
                $buildInformation.ReleaseDate = "08/26/2014"
            } elseif ($buildAndRevision -lt 1076.9) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU7
                $buildInformation.FriendlyName += "CU7"
                $buildInformation.ReleaseDate = "12/09/2014"
            } elseif ($buildAndRevision -lt 1104.5) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU8
                $buildInformation.FriendlyName += "CU8"
                $buildInformation.ReleaseDate = "03/17/2015"
            } elseif ($buildAndRevision -lt 1130.7) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU9
                $buildInformation.FriendlyName += "CU9"
                $buildInformation.ReleaseDate = "06/17/2015"
            } elseif ($buildAndRevision -lt 1156.6) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU10
                $buildInformation.FriendlyName += "CU10"
                $buildInformation.ReleaseDate = "09/15/2015"
            } elseif ($buildAndRevision -lt 1178.4) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU11
                $buildInformation.FriendlyName += "CU11"
                $buildInformation.ReleaseDate = "12/15/2015"
            } elseif ($buildAndRevision -lt 1210.3) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU12
                $buildInformation.FriendlyName += "CU12"
                $buildInformation.ReleaseDate = "03/15/2016"
            } elseif ($buildAndRevision -lt 1236.3) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU13
                $buildInformation.FriendlyName += "CU13"
                $buildInformation.ReleaseDate = "06/21/2016"
            } elseif ($buildAndRevision -lt 1263.5) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU14
                $buildInformation.FriendlyName += "CU14"
                $buildInformation.ReleaseDate = "09/20/2016"
            } elseif ($buildAndRevision -lt 1293.2) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU15
                $buildInformation.FriendlyName += "CU15"
                $buildInformation.ReleaseDate = "12/13/2016"
            } elseif ($buildAndRevision -lt 1320.4) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU16
                $buildInformation.FriendlyName += "CU16"
                $buildInformation.ReleaseDate = "03/21/2017"
            } elseif ($buildAndRevision -lt 1347.2) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU17
                $buildInformation.FriendlyName += "CU17"
                $buildInformation.ReleaseDate = "06/24/2017"
            } elseif ($buildAndRevision -lt 1365.1) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU18
                $buildInformation.FriendlyName += "CU18"
                $buildInformation.ReleaseDate = "09/16/2017"
            } elseif ($buildAndRevision -lt 1367.3) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU19
                $buildInformation.FriendlyName += "CU19"
                $buildInformation.ReleaseDate = "12/19/2017"
            } elseif ($buildAndRevision -lt 1395.4) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU20
                $buildInformation.FriendlyName += "CU20"
                $buildInformation.ReleaseDate = "03/20/2018"
            } elseif ($buildAndRevision -lt 1473.3) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU21
                $buildInformation.FriendlyName += "CU21"
                $buildInformation.ReleaseDate = "06/19/2018"
            } elseif ($buildAndRevision -lt 1497.2) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU22
                $buildInformation.FriendlyName += "CU22"
                $buildInformation.ReleaseDate = "02/12/2019"
            } elseif ($buildAndRevision -ge 1497.2) {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU23
                $buildInformation.FriendlyName += "CU23"
                $buildInformation.ReleaseDate = "06/18/2019"
                $buildInformation.SupportedBuild = $true
            }

            #Exchange 2013 .NET Information
            if ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU4) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU13) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU15) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d1wFix
            } elseif ($buildInformation.CU -eq [HealthChecker.ExchangeCULevel]::CU15) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d1
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU19) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU21) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d1
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU23) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d1
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
            } else {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
            }
        }

        try {
            $organizationConfig = Get-OrganizationConfig -ErrorAction Stop
            $exchangeInformation.GetOrganizationConfig = $organizationConfig
        } catch {
            Write-Yellow "Failed to run Get-OrganizationConfig."
            Invoke-CatchActions
        }

        $mitigationsEnabled = $null
        if ($null -ne $organizationConfig) {
            $mitigationsEnabled = $organizationConfig.MitigationsEnabled
        }

        $exchangeInformation.ExchangeEmergencyMitigationService = Get-ExchangeEmergencyMitigationServiceState `
            -RequiredInformation ([PSCustomObject]@{
                ComputerName       = $Script:Server
                MitigationsEnabled = $mitigationsEnabled
                GetExchangeServer  = $exchangeInformation.GetExchangeServer
            }) `
            -CatchActionFunction ${Function:Invoke-CatchActions}

        if ($null -ne $organizationConfig) {
            $exchangeInformation.MapiHttpEnabled = $organizationConfig.MapiHttpEnabled
            if ($null -ne $organizationConfig.EnableDownloadDomains) {
                $exchangeInformation.EnableDownloadDomains = $organizationConfig.EnableDownloadDomains
            }
        } else {
            Write-Verbose "MAPI HTTP Enabled and Download Domains Enabled results not accurate"
        }

        if ($buildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {
            $exchangeInformation.ApplicationPools = Get-ExchangeAppPoolsInformation
            try {
                $exchangeInformation.GetHybridConfiguration = Get-HybridConfiguration -ErrorAction Stop
            } catch {
                Write-Yellow "Failed to run Get-HybridConfiguration"
                Invoke-CatchActions
            }
        }

        $serverExchangeBinDirectory = Invoke-ScriptBlockHandler -ComputerName $Script:Server `
            -ScriptBlockDescription "Getting Exchange Bin Directory" `
            -CatchActionFunction ${Function:Invoke-CatchActions} `
            -ScriptBlock {
            "{0}Bin\" -f (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath
        }
        Write-Verbose "Found Exchange Bin: $serverExchangeBinDirectory"
        $exchangeInformation.ApplicationConfigFileStatus = Get-ExchangeApplicationConfigurationFileValidation -ConfigFileLocation ("{0}EdgeTransport.exe.config" -f $serverExchangeBinDirectory)

        $buildInformation.KBsInstalled = Get-ExchangeUpdates -ExchangeMajorVersion $buildInformation.MajorVersion
        if (($null -ne $buildInformation.KBsInstalled) -and ($buildInformation.KBsInstalled -like "*KB5000871*")) {
            Write-Verbose "March 2021 SU: KB5000871 was detected on the system"
            $buildInformation.March2021SUInstalled = $true
        } else {
            Write-Verbose "March 2021 SU: KB5000871 was not detected on the system"
            $buildInformation.March2021SUInstalled = $false
        }

        Write-Verbose "Query schema class information for CVE-2021-34470 testing"
        try {
            $exchangeInformation.msExchStorageGroup = Get-ExchangeAdSchemaClass -SchemaClassName "ms-Exch-Storage-Group"
        } catch {
            Write-Verbose "Failed to run Get-ExchangeAdSchemaClass"
            Invoke-CatchActions
        }

        $exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage = Get-RemoteRegistryValue -MachineName $Script:Server `
            -SubKey "SOFTWARE\Microsoft\ExchangeServer\v15\Search\SystemParameters" `
            -GetValue "CtsProcessorAffinityPercentage" `
            -CatchActionFunction ${Function:Invoke-CatchActions}
        $exchangeInformation.RegistryValues.FipsAlgorithmPolicyEnabled = Get-RemoteRegistryValue -MachineName $Script:Server `
            -SubKey "SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy" `
            -GetValue "Enabled" `
            -CatchActionFunction ${Function:Invoke-CatchActions}
        $exchangeInformation.ServerMaintenance = Get-ExchangeServerMaintenanceState -ComponentsToSkip "ForwardSyncDaemon", "ProvisioningRps"

        if (($buildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::ClientAccess) -and
            ($buildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::None)) {
            try {
                $testServiceHealthResults = Test-ServiceHealth -Server $Script:Server -ErrorAction Stop
                foreach ($notRunningService in $testServiceHealthResults.ServicesNotRunning) {
                    if ($exchangeInformation.ExchangeServicesNotRunning -notcontains $notRunningService) {
                        $exchangeInformation.ExchangeServicesNotRunning += $notRunningService
                    }
                }
            } catch {
                Write-Verbose "Failed to run Test-ServiceHealth"
                Invoke-CatchActions
            }
        }
    } elseif ($buildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2010) {
        Write-Verbose "Exchange 2010 detected."
        $buildInformation.FriendlyName = "Exchange 2010"
        $buildInformation.BuildNumber = $exchangeInformation.GetExchangeServer.AdminDisplayVersion.ToString()
    }

    Write-Verbose "Exiting: Get-ExchangeInformation"
    return $exchangeInformation
}




Function Get-WmiObjectHandler {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWMICmdlet', '', Justification = 'This is what this function is for')]
    [CmdletBinding()]
    param(
        [string]
        $ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory = $true)]
        [string]
        $Class,

        [string]
        $Filter,

        [string]
        $Namespace,

        [scriptblock]
        $CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - ComputerName: '$ComputerName' | Class: '$Class' | Filter: '$Filter' | Namespace: '$Namespace'"

        $execute = @{
            ComputerName = $ComputerName
            Class        = $Class
        }

        if (-not ([string]::IsNullOrEmpty($Filter))) {
            $execute.Add("Filter", $Filter)
        }

        if (-not ([string]::IsNullOrEmpty($Namespace))) {
            $execute.Add("Namespace", $Namespace)
        }
    }
    process {
        try {
            $wmi = Get-WmiObject @execute -ErrorAction Stop
            return $wmi
        } catch {
            Write-Verbose "Failed to run Get-WmiObject on class '$class'"

            if ($null -ne $CatchActionFunction) {
                & $CatchActionFunction
            }
        }
    }
}
Function Get-ProcessorInformation {
    [CmdletBinding()]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $wmiObject = $null
        $processorName = [string]::Empty
        $maxClockSpeed = 0
        $numberOfLogicalCores = 0
        $numberOfPhysicalCores = 0
        $numberOfProcessors = 0
        $currentClockSpeed = 0
        $processorIsThrottled = $false
        $differentProcessorCoreCountDetected = $false
        $differentProcessorsDetected = $false
        $presentedProcessorCoreCount = 0
        $previousProcessor = $null
    }
    process {
        $wmiObject = @(Get-WmiObjectHandler -ComputerName $MachineName -Class "Win32_Processor" -CatchActionFunction $CatchActionFunction)
        $processorName = $wmiObject[0].Name
        $maxClockSpeed = $wmiObject[0].MaxClockSpeed
        Write-Verbose "Evaluating processor results"

        foreach ($processor in $wmiObject) {
            $numberOfPhysicalCores += $processor.NumberOfCores
            $numberOfLogicalCores += $processor.NumberOfLogicalProcessors
            $numberOfProcessors++

            if ($processor.CurrentClockSpeed -lt $processor.MaxClockSpeed) {
                Write-Verbose "Processor is being throttled"
                $processorIsThrottled = $true
                $currentClockSpeed = $processor.CurrentClockSpeed
            }

            if ($null -ne $previousProcessor) {

                if ($processor.Name -ne $previousProcessor.Name -or
                    $processor.MaxClockSpeed -ne $previousProcessor.MaxClockSpeed) {
                    Write-Verbose "Different Processors are detected!!! This is an issue."
                    $differentProcessorsDetected = $true
                }

                if ($processor.NumberOfLogicalProcessors -ne $previousProcessor.NumberOfLogicalProcessors) {
                    Write-Verbose "Different Processor core count per processor socket detected. This is an issue."
                    $differentProcessorCoreCountDetected = $true
                }
            }
            $previousProcessor = $processor
        }

        $presentedProcessorCoreCount = Invoke-ScriptBlockHandler -ComputerName $MachineName `
            -ScriptBlock { [System.Environment]::ProcessorCount } `
            -ScriptBlockDescription "Trying to get the System.Environment ProcessorCount" `
            -CatchActionFunction $CatchActionFunction

        if ($null -eq $presentedProcessorCoreCount) {
            Write-Verbose "Wasn't able to get Presented Processor Core Count on the Server. Setting to -1."
            $presentedProcessorCoreCount = -1
        }
    }
    end {
        Write-Verbose "PresentedProcessorCoreCount: $presentedProcessorCoreCount"
        Write-Verbose "NumberOfPhysicalCores: $numberOfPhysicalCores | NumberOfLogicalCores: $numberOfLogicalCores | NumberOfProcessors: $numberOfProcessors"
        Write-Verbose "ProcessorIsThrottled: $processorIsThrottled | CurrentClockSpeed: $currentClockSpeed"
        Write-Verbose "DifferentProcessorsDetected: $differentProcessorsDetected | DifferentProcessorCoreCountDetected: $differentProcessorCoreCountDetected"
        return [PSCustomObject]@{
            Name                                = $processorName
            MaxMegacyclesPerCore                = $maxClockSpeed
            NumberOfPhysicalCores               = $numberOfPhysicalCores
            NumberOfLogicalCores                = $numberOfLogicalCores
            NumberOfProcessors                  = $numberOfProcessors
            CurrentMegacyclesPerCore            = $currentClockSpeed
            ProcessorIsThrottled                = $processorIsThrottled
            DifferentProcessorsDetected         = $differentProcessorsDetected
            DifferentProcessorCoreCountDetected = $differentProcessorCoreCountDetected
            EnvironmentProcessorCount           = $presentedProcessorCoreCount
            ProcessorClassObject                = $wmiObject
        }
    }
}

Function Get-ServerType {
    [CmdletBinding()]
    [OutputType("System.String")]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $ServerType
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - ServerType: $ServerType"
        $returnServerType = [string]::Empty
    }
    process {
        if ($ServerType -like "VMWare*") { $returnServerType = "VMware" }
        elseif ($ServerType -like "*Amazon EC2*") { $returnServerType = "AmazonEC2" }
        elseif ($ServerType -like "*Microsoft Corporation*") { $returnServerType = "HyperV" }
        elseif ($ServerType.Length -gt 0) { $returnServerType = "Physical" }
        else { $returnServerType = "Unknown" }
    }
    end {
        Write-Verbose "Returning: $returnServerType"
        return $returnServerType
    }
}
Function Get-HardwareInformation {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    [HealthChecker.HardwareInformation]$hardware_obj = New-Object HealthChecker.HardwareInformation
    $system = Get-WmiObjectHandler -ComputerName $Script:Server -Class "Win32_ComputerSystem" -CatchActionFunction ${Function:Invoke-CatchActions}
    $hardware_obj.MemoryInformation = Get-WmiObjectHandler -ComputerName $Script:Server -Class "Win32_PhysicalMemory" -CatchActionFunction ${Function:Invoke-CatchActions}
    $hardware_obj.Manufacturer = $system.Manufacturer
    $hardware_obj.System = $system
    $hardware_obj.AutoPageFile = $system.AutomaticManagedPagefile
    ForEach ($memory in $hardware_obj.MemoryInformation) {
        $hardware_obj.TotalMemory += $memory.Capacity
    }
    $hardware_obj.ServerType = (Get-ServerType -ServerType $system.Manufacturer)
    $processorInformation = Get-ProcessorInformation -MachineName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}

    #Need to do it this way because of Windows 2012R2
    $processor = New-Object HealthChecker.ProcessorInformation
    $processor.Name = $processorInformation.Name
    $processor.NumberOfPhysicalCores = $processorInformation.NumberOfPhysicalCores
    $processor.NumberOfLogicalCores = $processorInformation.NumberOfLogicalCores
    $processor.NumberOfProcessors = $processorInformation.NumberOfProcessors
    $processor.MaxMegacyclesPerCore = $processorInformation.MaxMegacyclesPerCore
    $processor.CurrentMegacyclesPerCore = $processorInformation.CurrentMegacyclesPerCore
    $processor.ProcessorIsThrottled = $processorInformation.ProcessorIsThrottled
    $processor.DifferentProcessorsDetected = $processorInformation.DifferentProcessorsDetected
    $processor.DifferentProcessorCoreCountDetected = $processorInformation.DifferentProcessorCoreCountDetected
    $processor.EnvironmentProcessorCount = $processorInformation.EnvironmentProcessorCount
    $processor.ProcessorClassObject = $processorInformation.ProcessorClassObject

    $hardware_obj.Processor = $processor
    $hardware_obj.Model = $system.Model

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $hardware_obj
}



Function Get-ServerRebootPending {
    [CmdletBinding()]
    param(
        [string]$ServerName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {

        Function Get-PendingFileReboot {
            try {
                if ((Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\" -Name PendingFileRenameOperations -ErrorAction Stop)) {
                    return $true
                }
                return $false
            } catch {
                throw
            }
        }

        Function Get-PendingCCMReboot {
            try {
                return (Invoke-CimMethod -Namespace 'Root\ccm\clientSDK' -ClassName 'CCM_ClientUtilities' -Name 'DetermineIfRebootPending' -ErrorAction Stop)
            } catch {
                throw
            }
        }

        Function Get-PathTestingReboot {
            param(
                [string]$TestingPath
            )

            return (Test-Path $TestingPath)
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $pendingRebootLocations = New-Object 'System.Collections.Generic.List[string]'
    }
    process {
        $pendingFileRenameOperationValue = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-PendingFileReboot} `
            -ScriptBlockDescription "Get-PendingFileReboot" `
            -CatchActionFunction $CatchActionFunction

        if ($null -eq $pendingFileRenameOperationValue) {
            $pendingFileRenameOperationValue = $false
        }

        $componentBasedServicingPendingRebootValue = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-PathTestingReboot} `
            -ScriptBlockDescription "Component Based Servicing Reboot Pending" `
            -ArgumentList "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" `
            -CatchActionFunction $CatchActionFunction

        $ccmReboot = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-PendingCCMReboot} `
            -ScriptBlockDescription "Get-PendingSCCMReboot" `
            -CatchActionFunction $CatchActionFunction

        $autoUpdatePendingRebootValue = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-PathTestingReboot} `
            -ScriptBlockDescription "Auto Update Pending Reboot" `
            -ArgumentList "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" `
            -CatchActionFunction $CatchActionFunction

        $ccmRebootPending = $ccmReboot -and ($ccmReboot.RebootPending -or $ccmReboot.IsHardRebootPending)
        $pendingReboot = $ccmRebootPending -or $pendingFileRenameOperationValue -or $componentBasedServicingPendingRebootValue -or $autoUpdatePendingRebootValue

        if ($ccmRebootPending) {
            Write-Verbose "RebootPending in CCM_ClientUtilities"
            $pendingRebootLocations.Add("CCM_ClientUtilities Showing Reboot Pending")
        }

        if ($pendingFileRenameOperationValue) {
            Write-Verbose "RebootPending at HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations"
            $pendingRebootLocations.Add("HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations")
        }

        if ($componentBasedServicingPendingRebootValue) {
            Write-Verbose "RebootPending at HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
            $pendingRebootLocations.Add("HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")
        }

        if ($autoUpdatePendingRebootValue) {
            Write-Verbose "RebootPending at HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
            $pendingRebootLocations.Add("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired")
        }
    }
    end {
        return [PSCustomObject]@{
            PendingFileRenameOperations          = $pendingFileRenameOperationValue
            ComponentBasedServicingPendingReboot = $componentBasedServicingPendingRebootValue
            AutoUpdatePendingReboot              = $autoUpdatePendingRebootValue
            CcmRebootPending                     = $ccmRebootPending
            PendingReboot                        = $pendingReboot
            PendingRebootLocations               = $pendingRebootLocations
        }
    }
}


Function Get-VisualCRedistributableInstalledVersion {
    [CmdletBinding()]
    param(
        [string]$ComputerName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $softwareList = New-Object 'System.Collections.Generic.List[object]'
    }
    process {
        $installedSoftware = Invoke-ScriptBlockHandler -ComputerName $ComputerName `
            -ScriptBlock { Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* } `
            -ScriptBlockDescription "Querying for software" `
            -CatchActionFunction $CatchActionFunction

        foreach ($software in $installedSoftware) {

            if ($software.PSObject.Properties.Name -contains "DisplayName" -and $software.DisplayName -like "Microsoft Visual C++ *") {
                Write-Verbose "Microsoft Visual C++ Found: $($software.DisplayName)"
                $softwareList.Add([PSCustomObject]@{
                        DisplayName       = $software.DisplayName
                        DisplayVersion    = $software.DisplayVersion
                        InstallDate       = $software.InstallDate
                        VersionIdentifier = $software.Version
                    })
            }
        }
    }
    end {
        Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
        return $softwareList
    }
}

Function Get-VisualCRedistributableInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet(2012, 2013)]
        [int]
        $Year
    )

    if ($Year -eq 2012) {
        return [PSCustomObject]@{
            VersionNumber = 184610406
            DownloadUrl   = "https://www.microsoft.com/en-us/download/details.aspx?id=30679"
            DisplayName   = "Microsoft Visual C++ 2012*"
        }
    } else {
        return [PSCustomObject]@{
            VersionNumber = 201367256
            DownloadUrl   = "https://support.microsoft.com/en-us/topic/update-for-visual-c-2013-redistributable-package-d8ccd6a5-4e26-c290-517b-8da6cfdf4f10"
            DisplayName   = "Microsoft Visual C++ 2013*"
        }
    }
}

Function Test-VisualCRedistributableInstalled {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet(2012, 2013)]
        [int]
        $Year,

        [Parameter(Mandatory = $true, Position = 1)]
        [object]
        $Installed
    )

    $desired = Get-VisualCRedistributableInfo $Year

    return ($null -ne ($Installed | Where-Object { $_.DisplayName -like $desired.DisplayName }))
}

Function Test-VisualCRedistributableUpToDate {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet(2012, 2013)]
        [int]
        $Year,

        [Parameter(Mandatory = $true, Position = 1)]
        [object]
        $Installed
    )

    $desired = Get-VisualCRedistributableInfo $Year

    return ($null -ne ($Installed | Where-Object {
                $_.DisplayName -like $desired.DisplayName -and $_.VersionIdentifier -eq $desired.VersionNumber
            }))
}

Function Get-AllTlsSettingsFromRegistry {
    [CmdletBinding()]
    [OutputType("System.Collections.Hashtable")]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {

        Function Get-TLSMemberValue {
            param(
                [Parameter(Mandatory = $true)]
                [string]
                $GetKeyType,

                [Parameter(Mandatory = $false)]
                [object]
                $KeyValue,

                [Parameter(Mandatory = $true)]
                [string]
                $ServerClientType,

                [Parameter(Mandatory = $true)]
                [string]
                $TlsVersion
            )
            Write-Verbose "KeyValue is null: '$($null -eq $KeyValue)' | KeyValue: '$KeyValue' | ServerClientType: $ServerClientType | TLSVersion: $tlsVersion | GetKeyType: $GetKeyType"
            switch ($GetKeyType) {
                "Enabled" {
                    return $null -eq $KeyValue -or $KeyValue -eq 1
                }
                "DisabledByDefault" {
                    return $null -ne $KeyValue -and $KeyValue -eq 1
                }
            }
        }

        Function Get-NETDefaultTLSValue {
            param(
                [Parameter(Mandatory = $false)]
                [object]
                $KeyValue,

                [Parameter(Mandatory = $true)]
                [string]
                $NetVersion,

                [Parameter(Mandatory = $true)]
                [string]
                $KeyName
            )
            Write-Verbose "KeyValue is null: '$($null -eq $KeyValue)' | KeyValue: '$KeyValue' | NetVersion: '$NetVersion' | KeyName: '$KeyName'"
            return $null -ne $KeyValue -and $KeyValue -eq 1
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - MachineName: '$MachineName'"
        $registryBase = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS {0}\{1}"
        $tlsVersions = @("1.0", "1.1", "1.2")
        $keyValues = ("Enabled", "DisabledByDefault")
        $netVersions = @("v2.0.50727", "v4.0.30319")
        $netRegistryBase = "SOFTWARE\{0}\.NETFramework\{1}"
        [HashTable]$allTlsObjects = @{}
    }
    process {
        foreach ($tlsVersion in $tlsVersions) {
            $registryServer = $registryBase -f $tlsVersion, "Server"
            $registryClient = $registryBase -f $tlsVersion, "Client"
            $currentTLSObject = [PSCustomObject]@{
                TLSVersion = $tlsVersion
            }

            foreach ($getKey in $keyValues) {
                $memberServerName = "Server$getKey"
                $memberClientName = "Client$getKey"

                $serverValue = Get-RemoteRegistryValue `
                    -MachineName $MachineName `
                    -SubKey $registryServer `
                    -GetValue $getKey `
                    -CatchActionFunction $CatchActionFunction
                $clientValue = Get-RemoteRegistryValue `
                    -MachineName $MachineName `
                    -SubKey $registryClient `
                    -GetValue $getKey `
                    -CatchActionFunction $CatchActionFunction

                $currentTLSObject | Add-Member -MemberType NoteProperty `
                    -Name $memberServerName `
                    -Value (Get-TLSMemberValue -GetKeyType $getKey -KeyValue $serverValue -ServerClientType "Server" -TlsVersion $tlsVersion)
                $currentTLSObject | Add-Member -MemberType NoteProperty `
                    -Name $memberClientName `
                    -Value (Get-TLSMemberValue -GetKeyType $getKey -KeyValue $clientValue -ServerClientType "Client" -TlsVersion $tlsVersion)
            }
            $allTlsObjects.Add($TlsVersion, $currentTLSObject)
        }

        foreach ($netVersion in $netVersions) {
            $currentNetTlsDefaultVersionObject = New-Object PSCustomObject
            $currentNetTlsDefaultVersionObject | Add-Member -MemberType NoteProperty -Name "NetVersion" -Value $netVersion

            $SystemDefaultTlsVersions = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey ($netRegistryBase -f "Microsoft", $netVersion) `
                -GetValue "SystemDefaultTlsVersions" `
                -CatchActionFunction $CatchActionFunction
            $SchUseStrongCrypto = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey ($netRegistryBase -f "Microsoft", $netVersion) `
                -GetValue "SchUseStrongCrypto" `
                -CatchActionFunction $CatchActionFunction
            $WowSystemDefaultTlsVersions = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey ($netRegistryBase -f "Wow6432Node\Microsoft", $netVersion) `
                -GetValue "SystemDefaultTlsVersions" `
                -CatchActionFunction $CatchActionFunction
            $WowSchUseStrongCrypto = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey ($netRegistryBase -f "Wow6432Node\Microsoft", $netVersion) `
                -GetValue "SchUseStrongCrypto" `
                -CatchActionFunction $CatchActionFunction

            $currentNetTlsDefaultVersionObject = [PSCustomObject]@{
                NetVersion                  = $netVersion
                SystemDefaultTlsVersions    = (Get-NETDefaultTLSValue -KeyValue $SystemDefaultTlsVersions -NetVersion $netVersion -KeyName "SystemDefaultTlsVersions")
                SchUseStrongCrypto          = (Get-NETDefaultTLSValue -KeyValue $SchUseStrongCrypto -NetVersion $netVersion -KeyName "SchUseStrongCrypto")
                WowSystemDefaultTlsVersions = (Get-NETDefaultTLSValue -KeyValue $WowSystemDefaultTlsVersions -NetVersion $netVersion -KeyName "WowSystemDefaultTlsVersions")
                WowSchUseStrongCrypto       = (Get-NETDefaultTLSValue -KeyValue $WowSchUseStrongCrypto -NetVersion $netVersion -KeyName "WowSchUseStrongCrypto")
                SecurityProtocol            = (Invoke-ScriptBlockHandler -ComputerName $MachineName -ScriptBlock { ([System.Net.ServicePointManager]::SecurityProtocol).ToString() } -CatchActionFunction $CatchActionFunction)
            }

            $hashKeyName = "NET{0}" -f ($netVersion.Split(".")[0])
            $allTlsObjects.Add($hashKeyName, $currentNetTlsDefaultVersionObject)
        }
        return $allTlsObjects
    }
}

Function Invoke-CatchActionErrorLoop {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [int]$CurrentErrors,
        [Parameter(Mandatory = $false, Position = 1)]
        [scriptblock]$CatchActionFunction
    )
    process {
        if ($null -ne $CatchActionFunction -and
            $Error.Count -ne $CurrentErrors) {
            $i = 0
            while ($i -lt ($Error.Count - $currentErrors)) {
                & $CatchActionFunction $Error[$i]
                $i++
            }
        }
    }
}
Function Get-AllNicInformation {
    [CmdletBinding()]
    param(
        [string]$ComputerName = $env:COMPUTERNAME,
        [string]$ComputerFQDN,
        [scriptblock]$CatchActionFunction
    )
    begin {

        Function Get-NicPnpCapabilitiesSetting {
            [CmdletBinding()]
            param(
                [ValidateNotNullOrEmpty()]
                [string]$NicAdapterComponentId
            )
            begin {
                $nicAdapterBasicPath = "SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}"
                [int]$i = 0
                [int]$retryCounter = 0
                Write-Verbose "Probing started to detect NIC adapter registry path"
                [int]$retryLimit = 3
            }
            process {
                do {
                    $nicAdapterPnPCapabilitiesProbingKey = "$nicAdapterBasicPath\$($i.ToString().PadLeft(4, "0"))"
                    $netCfgInstanceId = Get-RemoteRegistryValue -MachineName $ComputerName `
                        -SubKey $nicAdapterPnPCapabilitiesProbingKey `
                        -GetValue "NetCfgInstanceId" `
                        -CatchActionFunction $CatchActionFunction

                    if ($netCfgInstanceId -eq $NicAdapterComponentId) {
                        Write-Verbose "Matching ComponentId found - now checking for PnPCapabilitiesValue"
                        $nicAdapterPnPCapabilitiesValue = Get-RemoteRegistryValue -MachineName $ComputerName `
                            -SubKey $nicAdapterPnPCapabilitiesProbingKey `
                            -GetValue "PnPCapabilities" `
                            -CatchActionFunction $CatchActionFunction
                        break
                    } else {
                        Write-Verbose "No matching ComponentId found"

                        if ($null -eq $netCfgInstanceId) {
                            $retryCounter++
                            Write-Verbose "Enumeration possibly interrupted. Attempt: $retryCounter/$retryLimit"
                        } else {
                            Write-Verbose "Reset the retry counter"
                            $retryCounter = 0
                        }
                        $i++
                    }
                } while ($retryCounter -lt $retryLimit)
            }
            end {
                return [PSCustomObject]@{
                    PnPCapabilities   = $nicAdapterPnPCapabilitiesValue
                    SleepyNicDisabled = ($nicAdapterPnPCapabilitiesValue -eq 24 -or $nicAdapterPnPCapabilitiesValue -eq 280)
                }
            }
        }

        Function Get-NetworkConfiguration {
            [CmdletBinding()]
            param(
                [string]$ComputerName
            )
            begin {
                $currentErrors = $Error.Count
                $params = @{
                    ErrorAction = "Stop"
                }
            }
            process {
                try {
                    if (($ComputerName).Split(".")[0] -ne $env:COMPUTERNAME) {
                        $cimSession = New-CimSession -ComputerName $ComputerName -ErrorAction Stop
                        $params.Add("CimSession", $cimSession)
                    }
                    $networkIpConfiguration = Get-NetIPConfiguration @params | Where-Object { $_.NetAdapter.MediaConnectionState -eq "Connected" }
                    Invoke-CatchActionErrorLoop -CurrentErrors $currentErrors -CatchActionFunction $CatchActionFunction
                    return $networkIpConfiguration
                } catch {
                    Write-Verbose "Failed to run Get-NetIPConfiguration. Error $($_.Exception)"
                    #just rethrow as caller will handle the catch
                    throw
                }
            }
        }

        Function Get-NicInformation {
            [CmdletBinding()]
            param(
                [array]$NetworkConfiguration,
                [bool]$WmiObject
            )
            begin {

                Function Get-IpvAddresses {
                    return [PSCustomObject]@{
                        Address        = ([string]::Empty)
                        Subnet         = ([string]::Empty)
                        DefaultGateway = ([string]::Empty)
                    }
                }

                if ($null -eq $NetworkConfiguration) {
                    Write-Verbose "NetworkConfiguration are null in New-NicInformation. Returning a null object."
                    return $null
                }

                $nicObjects = New-Object 'System.Collections.Generic.List[object]'
            }
            process {
                if ($WmiObject) {
                    $networkAdapterConfigurations = Get-WmiObjectHandler -ComputerName $ComputerName `
                        -Class "Win32_NetworkAdapterConfiguration" `
                        -Filter "IPEnabled = True" `
                        -CatchActionFunction $CatchActionFunction
                }

                foreach ($networkConfig in $NetworkConfiguration) {
                    $dnsClient = $null
                    $rssEnabledValue = 2
                    $netAdapterRss = $null
                    $mtuSize = 0
                    $driverDate = [DateTime]::MaxValue
                    $driverVersion = [string]::Empty
                    $description = [string]::Empty
                    $ipv4Address = @()
                    $ipv6Address = @()
                    $ipv6Enabled = $false

                    if (-not ($WmiObject)) {
                        Write-Verbose "Working on NIC: $($networkConfig.InterfaceDescription)"
                        $adapter = $networkConfig.NetAdapter

                        if ($adapter.DriverFileName -ne "NdisImPlatform.sys") {
                            $nicPnpCapabilitiesSetting = Get-NicPnpCapabilitiesSetting -NicAdapterComponentId $adapter.DeviceID
                        } else {
                            Write-Verbose "Multiplexor adapter detected. Going to skip PnpCapabilities check"
                            $nicPnpCapabilitiesSetting = [PSCustomObject]@{
                                PnPCapabilities = "MultiplexorNoPnP"
                            }
                        }

                        try {
                            $dnsClient = $adapter | Get-DnsClient -ErrorAction Stop
                            $isRegisteredInDns = $dnsClient.RegisterThisConnectionsAddress
                            Write-Verbose "Got DNS Client information"
                        } catch {
                            Write-Verbose "Failed to get the DNS client information"
                            Invoke-CatchActionError $CatchActionFunction
                        }

                        try {
                            $netAdapterRss = $adapter | Get-NetAdapterRss -ErrorAction Stop
                            Write-Verbose "Got Net Adapter RSS Information"

                            if ($null -ne $netAdapterRss) {
                                [int]$rssEnabledValue = $netAdapterRss.Enabled
                            }
                        } catch {
                            Write-Verbose "Failed to get RSS Information"
                            Invoke-CatchActionError $CatchActionFunction
                        }

                        foreach ($ipAddress in $networkConfig.AllIPAddresses.IPAddress) {
                            if ($ipAddress.Contains(":")) {
                                $ipv6Enabled = $true
                            }
                        }

                        for ($i = 0; $i -lt $networkConfig.IPv4Address.Count; $i++) {
                            $newIpvAddress = Get-IpvAddresses

                            if ($null -ne $networkConfig.IPv4Address -and
                                $i -lt $networkConfig.IPv4Address.Count) {
                                $newIpvAddress.Address = $networkConfig.IPv4Address[$i].IPAddress
                                $newIpvAddress.Subnet = $networkConfig.IPv4Address[$i].PrefixLength
                            }

                            if ($null -ne $networkConfig.IPv4DefaultGateway -and
                                $i -lt $networkConfig.IPv4Address.Count) {
                                $newIpvAddress.DefaultGateway = $networkConfig.IPv4DefaultGateway[$i].NextHop
                            }
                            $ipv4Address += $newIpvAddress
                        }

                        for ($i = 0; $i -lt $networkConfig.IPv6Address.Count; $i++) {
                            $newIpvAddress = Get-IpvAddresses

                            if ($null -ne $networkConfig.IPv6Address -and
                                $i -lt $networkConfig.IPv6Address.Count) {
                                $newIpvAddress.Address = $networkConfig.IPv6Address[$i].IPAddress
                                $newIpvAddress.Subnet = $networkConfig.IPv6Address[$i].PrefixLength
                            }

                            if ($null -ne $networkConfig.IPv6DefaultGateway -and
                                $i -lt $networkConfig.IPv6DefaultGateway.Count) {
                                $newIpvAddress.DefaultGateway = $networkConfig.IPv6DefaultGateway[$i].NextHop
                            }
                            $ipv6Address += $newIpvAddress
                        }

                        $mtuSize = $adapter.MTUSize
                        $driverDate = $adapter.DriverDate
                        $driverVersion = $adapter.DriverVersionString
                        $description = $adapter.InterfaceDescription
                        $dnsServerToBeUsed = $networkConfig.DNSServer.ServerAddresses
                    } else {
                        Write-Verbose "Working on NIC: $($networkConfig.Description)"
                        $adapter = $networkConfig
                        $description = $adapter.Description

                        if ($adapter.ServiceName -ne "NdisImPlatformMp") {
                            $nicPnpCapabilitiesSetting = Get-NicPnpCapabilitiesSetting -NicAdapterComponentId $adapter.Guid
                        } else {
                            Write-Verbose "Multiplexor adapter detected. Going to skip PnpCapabilities check"
                            $nicPnpCapabilitiesSetting = [PSCustomObject]@{
                                PnPCapabilities = "MultiplexorNoPnP"
                            }
                        }

                        #set the correct $adapterConfiguration to link to the correct $networkConfig that we are on
                        $adapterConfiguration = $networkAdapterConfigurations |
                            Where-Object { $_.SettingID -eq $networkConfig.GUID -or
                                $_.SettingID -eq $networkConfig.InterfaceGuid }

                        if ($null -eq $adapterConfiguration) {
                            Write-Verbose "Failed to find correct adapterConfiguration for this networkConfig."
                            Write-Verbose "GUID: $($networkConfig.GUID) | InterfaceGuid: $($networkConfig.InterfaceGuid)"
                        }

                        $ipv6Enabled = ($adapterConfiguration.IPAddress | Where-Object { $_.Contains(":") }).Count -ge 1

                        $ipv4Gateway = $adapterConfiguration.DefaultIPGateway | Where-Object { $_.Contains(".") }
                        $ipv6Gateway = $adapterConfiguration.DefaultIPGateway | Where-Object { $_.Contains(":") }

                        for ($i = 0; $i -lt $adapterConfiguration.IPAddress.Count; $i++) {

                            if ($adapterConfiguration.IPAddress[$i].Contains(":")) {
                                $newIpv6Address = Get-IpvAddresses
                                if ($i -lt $adapterConfiguration.IPAddress.Count) {
                                    $newIpv6Address.Address = $adapterConfiguration.IPAddress[$i]
                                    $newIpv6Address.Subnet = $adapterConfiguration.IPSubnet[$i]
                                }

                                $newIpv6Address.DefaultGateway = $ipv6Gateway
                                $ipv6Address += $newIpv6Address
                            } else {
                                $newIpv4Address = Get-IpvAddresses
                                if ($i -lt $adapterConfiguration.IPAddress.Count) {
                                    $newIpv4Address.Address = $adapterConfiguration.IPAddress[$i]
                                    $newIpv4Address.Subnet = $adapterConfiguration.IPSubnet[$i]
                                }

                                $newIpv4Address.DefaultGateway = $ipv4Gateway
                                $ipv4Address += $newIpv4Address
                            }
                        }

                        $isRegisteredInDns = $adapterConfiguration.FullDNSRegistrationEnabled
                        $dnsServerToBeUsed = $adapterConfiguration.DNSServerSearchOrder
                    }

                    $nicObjects.Add([PSCustomObject]@{
                            WmiObject         = $WmiObject
                            Name              = $adapter.Name
                            LinkSpeed         = ((($adapter.Speed) / 1000000).ToString() + " Mbps")
                            DriverDate        = $driverDate
                            NetAdapterRss     = $netAdapterRss
                            RssEnabledValue   = $rssEnabledValue
                            IPv6Enabled       = $ipv6Enabled
                            Description       = $description
                            DriverVersion     = $driverVersion
                            MTUSize           = $mtuSize
                            PnPCapabilities   = $nicPnpCapabilitiesSetting.PnpCapabilities
                            SleepyNicDisabled = $nicPnpCapabilitiesSetting.SleepyNicDisabled
                            IPv4Addresses     = $ipv4Address
                            IPv6Addresses     = $ipv6Address
                            RegisteredInDns   = $isRegisteredInDns
                            DnsServer         = $dnsServerToBeUsed
                            DnsClient         = $dnsClient
                        })
                }
            }
            end {
                Write-Verbose "Found $($nicObjects.Count) active adapters on the computer."
                Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
                return $nicObjects
            }
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - ComputerName: '$ComputerName' | ComputerFQDN: '$ComputerFQDN'"
    }
    process {
        try {
            try {
                $networkConfiguration = Get-NetworkConfiguration -ComputerName $ComputerName
            } catch {
                Invoke-CatchActionError $CatchActionFunction

                try {
                    if (-not ([string]::IsNullOrEmpty($ComputerFQDN))) {
                        $networkConfiguration = Get-NetworkConfiguration -ComputerName $ComputerFQDN
                    } else {
                        $bypassCatchActions = $true
                        Write-Verbose "No FQDN was passed, going to rethrow error."
                        throw
                    }
                } catch {
                    #Just throw again
                    throw
                }
            }

            if ([String]::IsNullOrEmpty($networkConfiguration)) {
                # Throw if nothing was returned by previous calls.
                # Can be caused when executed on Server 2008 R2 where CIM namespace ROOT/StandardCimv2 is invalid.
                Write-Verbose "No value was returned by 'Get-NetworkConfiguration'. Fallback to WMI."
                throw
            }

            return (Get-NicInformation -NetworkConfiguration $networkConfiguration)
        } catch {
            if (-not $bypassCatchActions) {
                Invoke-CatchActionError $CatchActionFunction
            }

            $wmiNetworkCards = Get-WmiObjectHandler -ComputerName $ComputerName `
                -Class "Win32_NetworkAdapter" `
                -Filter "NetConnectionStatus ='2'" `
                -CatchActionFunction $CatchActionFunction

            return (Get-NicInformation -NetworkConfiguration $wmiNetworkCards -WmiObject $true)
        }
    }
}

Function Get-CredentialGuardEnabled {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $registryValue = Get-RemoteRegistryValue -MachineName $Script:Server `
        -SubKey "SYSTEM\CurrentControlSet\Control\LSA" `
        -GetValue "LsaCfgFlags" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    if ($null -ne $registryValue -and
        $registryValue -ne 0) {
        return $true
    }

    return $false
}

Function Get-HttpProxySetting {

    $httpProxy32 = [String]::Empty
    $httpProxy64 = [String]::Empty
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    Function Get-WinHttpSettings {
        param(
            [Parameter(Mandatory = $true)][string]$RegistryLocation
        )
        $connections = Get-ItemProperty -Path $RegistryLocation
        $Proxy = [string]::Empty
        if (($null -ne $connections) -and
            ($Connections | Get-Member).Name -contains "WinHttpSettings") {
            foreach ($Byte in $Connections.WinHttpSettings) {
                if ($Byte -ge 48) {
                    $Proxy += [CHAR]$Byte
                }
            }
        }
        return $(if ($Proxy -eq [string]::Empty) { "<None>" } else { $Proxy })
    }

    $httpProxy32 = Invoke-ScriptBlockHandler -ComputerName $Script:Server -ScriptBlock ${Function:Get-WinHttpSettings} -ArgumentList "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections" -ScriptBlockDescription "Getting 32 Http Proxy Value" -CatchActionFunction ${Function:Invoke-CatchActions}
    $httpProxy64 = Invoke-ScriptBlockHandler -ComputerName $Script:Server -ScriptBlock ${Function:Get-WinHttpSettings} -ArgumentList "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\Connections" -ScriptBlockDescription "Getting 64 Http Proxy Value" -CatchActionFunction ${Function:Invoke-CatchActions}

    Write-Verbose "Http Proxy 32: $httpProxy32"
    Write-Verbose "Http Proxy 64: $httpProxy64"
    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"

    if ($httpProxy32 -ne "<None>") {
        return $httpProxy32
    } else {
        return $httpProxy64
    }
}

Function Get-LmCompatibilityLevelInformation {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    [HealthChecker.LmCompatibilityLevelInformation]$ServerLmCompatObject = New-Object -TypeName HealthChecker.LmCompatibilityLevelInformation
    $registryValue = Get-RemoteRegistryValue -RegistryHive "LocalMachine" `
        -MachineName $Script:Server `
        -SubKey "SYSTEM\CurrentControlSet\Control\Lsa" `
        -GetValue "LmCompatibilityLevel" `
        -ValueType "DWord" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    if ($null -eq $registryValue) {
        $registryValue = 3
    }

    $ServerLmCompatObject.RegistryValue = $registryValue
    Write-Verbose "LmCompatibilityLevel Registry Value: $registryValue"

    Switch ($ServerLmCompatObject.RegistryValue) {
        0 { $ServerLmCompatObject.Description = "Clients use LM and NTLM authentication, but they never use NTLMv2 session security. Domain controllers accept LM, NTLM, and NTLMv2 authentication." }
        1 { $ServerLmCompatObject.Description = "Clients use LM and NTLM authentication, and they use NTLMv2 session security if the server supports it. Domain controllers accept LM, NTLM, and NTLMv2 authentication." }
        2 { $ServerLmCompatObject.Description = "Clients use only NTLM authentication, and they use NTLMv2 session security if the server supports it. Domain controller accepts LM, NTLM, and NTLMv2 authentication." }
        3 { $ServerLmCompatObject.Description = "Clients use only NTLMv2 authentication, and they use NTLMv2 session security if the server supports it. Domain controllers accept LM, NTLM, and NTLMv2 authentication." }
        4 { $ServerLmCompatObject.Description = "Clients use only NTLMv2 authentication, and they use NTLMv2 session security if the server supports it. Domain controller refuses LM authentication responses, but it accepts NTLM and NTLMv2." }
        5 { $ServerLmCompatObject.Description = "Clients use only NTLMv2 authentication, and they use NTLMv2 session security if the server supports it. Domain controller refuses LM and NTLM authentication responses, but it accepts NTLMv2." }
    }

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    Return $ServerLmCompatObject
}

Function Get-PageFileInformation {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    [HealthChecker.PageFileInformation]$page_obj = New-Object HealthChecker.PageFileInformation
    $pageFile = Get-WmiObjectHandler -ComputerName $Script:Server -Class "Win32_PageFileSetting" -CatchActionFunction ${Function:Invoke-CatchActions}

    if ($null -ne $pageFile) {
        if ($pageFile.GetType().Name -eq "ManagementObject") {
            $page_obj.MaxPageSize = $pageFile.MaximumSize
        }
        $page_obj.PageFile = $pageFile
    } else {
        Write-Verbose "Return Null value"
    }

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $page_obj
}

Function Get-ServerOperatingSystemVersion {
    [CmdletBinding()]
    [OutputType("System.String")]
    param(
        [string]$OsCaption
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $osReturnValue = [string]::Empty
    }
    process {
        if ([string]::IsNullOrEmpty($OsCaption)) {
            Write-Verbose "Getting the local machine version build number"
            $OsCaption = (Get-WmiObjectHandler -Class "Win32_OperatingSystem").Caption
        }
        Write-Verbose "OsCaption: '$OsCaption'"

        switch -Wildcard ($OsCaption) {
            "*Server 2008 R2*" { $osReturnValue = "Windows2008R2"; break }
            "*Server 2008*" { $osReturnValue = "Windows2008" }
            "*Server 2012 R2*" { $osReturnValue = "Windows2012R2"; break }
            "*Server 2012*" { $osReturnValue = "Windows2012" }
            "*Server 2016*" { $osReturnValue = "Windows2016" }
            "*Server 2019*" { $osReturnValue = "Windows2019" }
            "Microsoft Windows Server Standard" { $osReturnValue = "WindowsCore" }
            "Microsoft Windows Server Datacenter" { $osReturnValue = "WindowsCore" }
            default { $osReturnValue = "Unknown" }
        }
    }
    end {
        Write-Verbose "Returned: '$osReturnValue'"
        return [string]$osReturnValue
    }
}

Function Get-Smb1ServerSettings {
    [CmdletBinding()]
    param(
        [string]$ServerName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $smbServerConfiguration = $null
        $windowsFeature = $null
    }
    process {
        $smbServerConfiguration = Invoke-ScriptBlockHandler -ComputerName $ServerName `
            -ScriptBlock { Get-SmbServerConfiguration } `
            -CatchActionFunction $CatchActionFunction `
            -ScriptBlockDescription "Get-SmbServerConfiguration"

        try {
            $windowsFeature = Get-WindowsFeature "FS-SMB1" -ComputerName $ServerName -ErrorAction Stop
        } catch {
            Write-Verbose "Failed to Get-WindowsFeature for FS-SMB1"
            Invoke-CatchActionError $CatchActionFunction
        }
    }
    end {
        return [PSCustomObject]@{
            SmbServerConfiguration = $smbServerConfiguration
            WindowsFeature         = $windowsFeature
            SuccessfulGetInstall   = $null -ne $windowsFeature
            SuccessfulGetBlocked   = $null -ne $smbServerConfiguration
            Installed              = $windowsFeature.Installed -eq $true
            IsBlocked              = $smbServerConfiguration.EnableSMB1Protocol -eq $false
        }
    }
}

Function Get-TimeZoneInformationRegistrySettings {
    [CmdletBinding()]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $timeZoneInformationSubKey = "SYSTEM\CurrentControlSet\Control\TimeZoneInformation"
        $actionsToTake = @()
    }
    process {
        $dynamicDaylightTimeDisabled = Get-RemoteRegistryValue -MachineName $MachineName -SubKey $timeZoneInformationSubKey -GetValue "DynamicDaylightTimeDisabled" -CatchActionFunction $CatchActionFunction
        $timeZoneKeyName = Get-RemoteRegistryValue -MachineName $MachineName -SubKey $timeZoneInformationSubKey -GetValue "TimeZoneKeyName" -CatchActionFunction $CatchActionFunction
        $standardStart = Get-RemoteRegistryValue -MachineName $MachineName -SubKey $timeZoneInformationSubKey -GetValue "StandardStart" -CatchActionFunction $CatchActionFunction
        $daylightStart = Get-RemoteRegistryValue -MachineName $MachineName -SubKey $timeZoneInformationSubKey -GetValue "DaylightStart" -CatchActionFunction $CatchActionFunction

        if ([string]::IsNullOrEmpty($timeZoneKeyName)) {
            Write-Verbose "TimeZoneKeyName is null or empty. Action should be taken to address this."
            $actionsToTake += "TimeZoneKeyName is blank. Need to switch your current time zone to a different value, then switch it back to have this value populated again."
        }

        $standardStartNonZeroValue = ($null -ne ($standardStart | Where-Object { $_ -ne 0 }))
        $daylightStartNonZeroValue = ($null -ne ($daylightStart | Where-Object { $_ -ne 0 }))

        if ($dynamicDaylightTimeDisabled -ne 0 -and
            ($standardStartNonZeroValue -or
            $daylightStartNonZeroValue)) {
            Write-Verbose "Determined that there is a chance the settings set could cause a DST issue."
            $dstIssueDetected = $true
            $actionsToTake += "High Warning: DynamicDaylightTimeDisabled is set, Windows can not properly detect any DST rule changes in your time zone. `
            It is possible that you could be running into this issue. Set 'Adjust for daylight saving time automatically to on'"
        } elseif ($dynamicDaylightTimeDisabled -ne 0) {
            Write-Verbose "Daylight savings auto adjustment is disabled."
            $actionsToTake += "Warning: DynamicDaylightTimeDisabled is set, Windows can not properly detect any DST rule changes in your time zone."
        }
    }
    end {
        return [PSCustomObject]@{
            DynamicDaylightTimeDisabled = $dynamicDaylightTimeDisabled
            TimeZoneKeyName             = $timeZoneKeyName
            StandardStart               = $standardStart
            DaylightStart               = $daylightStart
            DstIssueDetected            = $dstIssueDetected
            ActionsToTake               = $actionsToTake
        }
    }
}


Function Get-CounterSamples {
    param(
        [Parameter(Mandatory = $true)][array]$MachineNames,
        [Parameter(Mandatory = $true)][array]$Counters
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    try {
        $counterSamples = (Get-Counter -ComputerName $MachineNames -Counter $Counters -ErrorAction Stop).CounterSamples
    } catch {
        Invoke-CatchActions
        Write-Verbose "Failed to get counter samples"
    }
    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $counterSamples
}
Function Get-OperatingSystemInformation {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    [HealthChecker.OperatingSystemInformation]$osInformation = New-Object HealthChecker.OperatingSystemInformation
    $win32_OperatingSystem = Get-WmiObjectHandler -ComputerName $Script:Server -Class Win32_OperatingSystem -CatchActionFunction ${Function:Invoke-CatchActions}
    $win32_PowerPlan = Get-WmiObjectHandler -ComputerName $Script:Server -Class Win32_PowerPlan -Namespace 'root\cimv2\power' -Filter "isActive='true'" -CatchActionFunction ${Function:Invoke-CatchActions}
    $currentDateTime = Get-Date
    $lastBootUpTime = [Management.ManagementDateTimeConverter]::ToDateTime($win32_OperatingSystem.lastbootuptime)
    $osInformation.BuildInformation.VersionBuild = $win32_OperatingSystem.Version
    $osInformation.BuildInformation.MajorVersion = (Get-ServerOperatingSystemVersion -OsCaption $win32_OperatingSystem.Caption)
    $osInformation.BuildInformation.FriendlyName = $win32_OperatingSystem.Caption
    $osInformation.BuildInformation.OperatingSystem = $win32_OperatingSystem
    $osInformation.ServerBootUp.Days = ($currentDateTime - $lastBootUpTime).Days
    $osInformation.ServerBootUp.Hours = ($currentDateTime - $lastBootUpTime).Hours
    $osInformation.ServerBootUp.Minutes = ($currentDateTime - $lastBootUpTime).Minutes
    $osInformation.ServerBootUp.Seconds = ($currentDateTime - $lastBootUpTime).Seconds

    if ($null -ne $win32_PowerPlan) {

        if ($win32_PowerPlan.InstanceID -eq "Microsoft:PowerPlan\{8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c}") {
            Write-Verbose "High Performance Power Plan is set to true"
            $osInformation.PowerPlan.HighPerformanceSet = $true
        } else { Write-Verbose "High Performance Power Plan is NOT set to true" }
        $osInformation.PowerPlan.PowerPlanSetting = $win32_PowerPlan.ElementName
    } else {
        Write-Verbose "Power Plan Information could not be read"
        $osInformation.PowerPlan.PowerPlanSetting = "N/A"
    }
    $osInformation.PowerPlan.PowerPlan = $win32_PowerPlan
    $osInformation.PageFile = Get-PageFileInformation
    $osInformation.NetworkInformation.NetworkAdapters = (Get-AllNicInformation -ComputerName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions} -ComputerFQDN $Script:ServerFQDN)
    foreach ($adapter in $osInformation.NetworkInformation.NetworkAdapters) {

        if (!$adapter.IPv6Enabled) {
            $osInformation.NetworkInformation.IPv6DisabledOnNICs = $true
            break
        }
    }

    $osInformation.NetworkInformation.IPv6DisabledComponents = Get-RemoteRegistryValue -MachineName $Script:Server `
        -SubKey "SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters" `
        -GetValue "DisabledComponents" `
        -ValueType "DWord" `
        -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.NetworkInformation.TCPKeepAlive = Get-RemoteRegistryValue -MachineName $Script:Server `
        -SubKey "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" `
        -GetValue "KeepAliveTime" `
        -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.NetworkInformation.RpcMinConnectionTimeout = Get-RemoteRegistryValue -MachineName $Script:Server `
        -SubKey "Software\Policies\Microsoft\Windows NT\RPC\" `
        -GetValue "MinimumConnectionTimeout" `
        -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.NetworkInformation.HttpProxy = Get-HttpProxySetting
    $osInformation.InstalledUpdates.HotFixes = (Get-HotFix -ComputerName $Script:Server -ErrorAction SilentlyContinue) #old school check still valid and faster and a failsafe
    $osInformation.LmCompatibility = Get-LmCompatibilityLevelInformation
    $counterSamples = (Get-CounterSamples -MachineNames $Script:Server -Counters "\Network Interface(*)\Packets Received Discarded")

    if ($null -ne $counterSamples) {
        $osInformation.NetworkInformation.PacketsReceivedDiscarded = $counterSamples
    }
    $serverReboot = (Get-ServerRebootPending -ServerName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions})
    $osInformation.ServerPendingReboot.PendingFileRenameOperations = $serverReboot.PendingFileRenameOperations
    $osInformation.ServerPendingReboot.SccmReboot = $serverReboot.SccmReboot
    $osInformation.ServerPendingReboot.SccmRebootPending = $serverReboot.SccmRebootPending
    $osInformation.ServerPendingReboot.ComponentBasedServicingRebootPending = $serverReboot.ComponentBasedServicingRebootPending
    $osInformation.ServerPendingReboot.AutoUpdatePendingReboot = $serverReboot.AutoUpdatePendingReboot
    $osInformation.ServerPendingReboot.PendingReboot = $serverReboot.PendingReboot
    $timeZoneInformation = Get-TimeZoneInformationRegistrySettings -MachineName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.TimeZone.DynamicDaylightTimeDisabled = $timeZoneInformation.DynamicDaylightTimeDisabled
    $osInformation.TimeZone.TimeZoneKeyName = $timeZoneInformation.TimeZoneKeyName
    $osInformation.TimeZone.StandardStart = $timeZoneInformation.StandardStart
    $osInformation.TimeZone.DaylightStart = $timeZoneInformation.DaylightStart
    $osInformation.TimeZone.DstIssueDetected = $timeZoneInformation.DstIssueDetected
    $osInformation.TimeZone.ActionsToTake = $timeZoneInformation.ActionsToTake
    $osInformation.TimeZone.CurrentTimeZone = Invoke-ScriptBlockHandler -ComputerName $Script:Server `
        -ScriptBlock { ([System.TimeZone]::CurrentTimeZone).StandardName } `
        -ScriptBlockDescription "Getting Current Time Zone" `
        -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.TLSSettings = Get-AllTlsSettingsFromRegistry -MachineName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.VcRedistributable = Get-VisualCRedistributableInstalledVersion -ComputerName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.CredentialGuardEnabled = Get-CredentialGuardEnabled
    $osInformation.RegistryValues.CurrentVersionUbr = Get-RemoteRegistryValue `
        -MachineName $Script:Server `
        -SubKey "SOFTWARE\Microsoft\Windows NT\CurrentVersion" `
        -GetValue "UBR" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    $osInformation.RegistryValues.LanManServerDisabledCompression = Get-RemoteRegistryValue `
        -MachineName $Script:Server `
        -SubKey "SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" `
        -GetValue "DisableCompression" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    $osInformation.Smb1ServerSettings = Get-Smb1ServerSettings -ServerName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $osInformation
}

Function Get-DotNetDllFileVersions {
    [CmdletBinding()]
    [OutputType("System.Collections.Hashtable")]
    param(
        [string]$ComputerName,
        [array]$FileNames,
        [scriptblock]$CatchActionFunction
    )

    begin {
        Function Invoke-ScriptBlockGetItem {
            param(
                [string]$FilePath
            )
            $getItem = Get-Item $FilePath

            $returnObject = ([PSCustomObject]@{
                    GetItem          = $getItem
                    LastWriteTimeUtc = $getItem.LastWriteTimeUtc
                    VersionInfo      = ([PSCustomObject]@{
                            FileMajorPart   = $getItem.VersionInfo.FileMajorPart
                            FileMinorPart   = $getItem.VersionInfo.FileMinorPart
                            FileBuildPart   = $getItem.VersionInfo.FileBuildPart
                            FilePrivatePart = $getItem.VersionInfo.FilePrivatePart
                        })
                })

            return $returnObject
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $dotNetInstallPath = [string]::Empty
        $files = @{}
    }
    process {
        $dotNetInstallPath = Get-RemoteRegistryValue -MachineName $ComputerName `
            -SubKey "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" `
            -GetValue "InstallPath" `
            -CatchActionFunction $CatchActionFunction

        if ([string]::IsNullOrEmpty($dotNetInstallPath)) {
            Write-Verbose "Failed to determine .NET install path"
            return
        }

        foreach ($fileName in $FileNames) {
            Write-Verbose "Querying for .NET DLL File $fileName"
            $getItem = Invoke-ScriptBlockHandler -ComputerName $ComputerName `
                -ScriptBlock ${Function:Invoke-ScriptBlockGetItem} `
                -ArgumentList ("{0}\{1}" -f $dotNetInstallPath, $filename) `
                -CatchActionFunction $CatchActionFunction
            $files.Add($fileName, $getItem)
        }
    }
    end {
        return $files
    }
}


Function Get-NETFrameworkVersion {
    [CmdletBinding()]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [int]$NetVersionKey = -1,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $friendlyName = [string]::Empty
        $minValue = -1
    }
    process {

        if ($NetVersionKey -eq -1) {
            [int]$NetVersionKey = Get-RemoteRegistryValue -MachineName $MachineName `
                -SubKey "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" `
                -GetValue "Release" `
                -CatchActionFunction $CatchActionFunction
        }

        #Using Minimum Version as per https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed?redirectedfrom=MSDN#minimum-version
        if ($NetVersionKey -lt 378389) {
            $friendlyName = "Unknown"
            $minValue = -1
        } elseif ($NetVersionKey -lt 378675) {
            $friendlyName = "4.5"
            $minValue = 378389
        } elseif ($NetVersionKey -lt 379893) {
            $friendlyName = "4.5.1"
            $minValue = 378675
        } elseif ($NetVersionKey -lt 393295) {
            $friendlyName = "4.5.2"
            $minValue = 379893
        } elseif ($NetVersionKey -lt 394254) {
            $friendlyName = "4.6"
            $minValue = 393295
        } elseif ($NetVersionKey -lt 394802) {
            $friendlyName = "4.6.1"
            $minValue = 394254
        } elseif ($NetVersionKey -lt 460798) {
            $friendlyName = "4.6.2"
            $minValue = 394802
        } elseif ($NetVersionKey -lt 461308) {
            $friendlyName = "4.7"
            $minValue = 460798
        } elseif ($NetVersionKey -lt 461808) {
            $friendlyName = "4.7.1"
            $minValue = 461308
        } elseif ($NetVersionKey -lt 528040) {
            $friendlyName = "4.7.2"
            $minValue = 461808
        } elseif ($NetVersionKey -ge 528040) {
            $friendlyName = "4.8"
            $minValue = 528040
        }
    }
    end {
        Write-Verbose "FriendlyName: $friendlyName | RegistryValue: $netVersionKey | MinimumValue: $minValue"
        return [PSCustomObject]@{
            FriendlyName  = $friendlyName
            RegistryValue = $NetVersionKey
            MinimumValue  = $minValue
        }
    }
}
Function Get-HealthCheckerExchangeServer {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    [HealthChecker.HealthCheckerExchangeServer]$HealthExSvrObj = New-Object -TypeName HealthChecker.HealthCheckerExchangeServer
    $HealthExSvrObj.ServerName = $Script:Server
    $HealthExSvrObj.HardwareInformation = Get-HardwareInformation
    $HealthExSvrObj.OSInformation = Get-OperatingSystemInformation
    $HealthExSvrObj.ExchangeInformation = Get-ExchangeInformation -OSMajorVersion $HealthExSvrObj.OSInformation.BuildInformation.MajorVersion

    if ($HealthExSvrObj.ExchangeInformation.BuildInformation.MajorVersion -ge [HealthChecker.ExchangeMajorVersion]::Exchange2013) {
        $netFrameworkVersion = Get-NETFrameworkVersion -MachineName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}
        $HealthExSvrObj.OSInformation.NETFramework.FriendlyName = $netFrameworkVersion.FriendlyName
        $HealthExSvrObj.OSInformation.NETFramework.RegistryValue = $netFrameworkVersion.RegistryValue
        $HealthExSvrObj.OSInformation.NETFramework.NetMajorVersion = $netFrameworkVersion.MinimumValue
        $HealthExSvrObj.OSInformation.NETFramework.FileInformation = Get-DotNetDllFileVersions -ComputerName $Script:Server -FileNames @("System.Data.dll", "System.Configuration.dll") -CatchActionFunction ${Function:Invoke-CatchActions}

        if ($netFrameworkVersion.MinimumValue -eq $HealthExSvrObj.ExchangeInformation.NETFramework.MaxSupportedVersion) {
            $HealthExSvrObj.ExchangeInformation.NETFramework.OnRecommendedVersion = $true
        }
    }
    $HealthExSvrObj.HealthCheckerVersion = $BuildVersion
    $HealthExSvrObj.GenerationTime = [datetime]::Now
    Write-Verbose "Finished building health Exchange Server Object for server: $Script:Server"
    return $HealthExSvrObj
}

Function Get-ErrorsThatOccurred {

    if ($Error.Count -gt 0) {
        Write-Grey(" "); Write-Grey(" ")
        Function Write-Errors {
            Write-Verbose "`r`n`r`nErrors that occurred that wasn't handled"

            $index = 0
            $Error |
                ForEach-Object {
                    $index++
                    $currentError = $_
                    $handledError = $Script:ErrorsExcluded |
                        Where-Object { $_.Equals($currentError) }

                        if ($null -eq $handledError) {
                            Write-Verbose "Error Index: $index"
                            Write-Verbose $currentError

                            if ($null -ne $currentError.ScriptStackTrace) {
                                Write-Verbose $currentError.ScriptStackTrace
                            }
                            Write-Verbose "-----------------------------------`r`n`r`n"
                        }
                    }

            Write-Verbose "`r`n`r`nErrors that were handled"
            $index = 0
            $Error |
                ForEach-Object {
                    $index++
                    $currentError = $_
                    $handledError = $Script:ErrorsExcluded |
                        Where-Object { $_.Equals($currentError) }

                        if ($null -ne $handledError) {
                            Write-Verbose "Error Index: $index"
                            Write-Verbose $handledError

                            if ($null -ne $handledError.ScriptStackTrace) {
                                Write-Verbose $handledError.ScriptStackTrace
                            }
                            Write-Verbose "-----------------------------------`r`n`r`n"
                        }
                    }
        }

        if ($Error.Count -ne $Script:ErrorsExcludedCount) {
            Write-Red("There appears to have been some errors in the script. To assist with debugging of the script, please send the HealthChecker-Debug_*.txt, HealthChecker-Errors.json, and .xml file to ExToolsFeedback@microsoft.com.")
            $Script:Logger.PreventLogCleanup = $true
            Write-Errors
            #Need to convert Error to Json because running into odd issues with trying to export $Error out in my lab. Got StackOverflowException for one of the errors i always see there.
            try {
                $Error |
                    ConvertTo-Json |
                    Out-File ("$OutputFilePath\HealthChecker-Errors.json")
            } catch {
                Write-Red("Failed to export the HealthChecker-Errors.json")
                Invoke-CatchActions
            }
        } elseif ($Script:VerboseEnabled -or
            $SaveDebugLog) {
            Write-Verbose "All errors that occurred were in try catch blocks and was handled correctly."
            $Script:Logger.PreventLogCleanup = $true
            Write-Errors
        }
    } else {
        Write-Verbose "No errors occurred in the script."
    }
}

Function Get-HealthCheckFilesItemsFromLocation {
    $items = Get-ChildItem $XMLDirectoryPath | Where-Object { $_.Name -like "HealthChecker-*-*.xml" }

    if ($null -eq $items) {
        Write-Host("Doesn't appear to be any Health Check XML files here....stopping the script")
        exit
    }
    return $items
}

Function Get-OnlyRecentUniqueServersXMLs {
    param(
        [Parameter(Mandatory = $true)][array]$FileItems
    )
    $aObject = @()

    foreach ($item in $FileItems) {
        $obj = New-Object PSCustomObject
        [string]$itemName = $item.Name
        $ServerName = $itemName.Substring(($itemName.IndexOf("-") + 1), ($itemName.LastIndexOf("-") - $itemName.IndexOf("-") - 1))
        $obj | Add-Member -MemberType NoteProperty -Name ServerName -Value $ServerName
        $obj | Add-Member -MemberType NoteProperty -Name FileName -Value $itemName
        $obj | Add-Member -MemberType NoteProperty -Name FileObject -Value $item
        $aObject += $obj
    }

    $grouped = $aObject | Group-Object ServerName
    $FilePathList = @()

    foreach ($gServer in $grouped) {

        if ($gServer.Count -gt 1) {
            #going to only use the most current file for this server providing that they are using the newest updated version of Health Check we only need to sort by name
            $groupData = $gServer.Group #because of win2008
            $FilePathList += ($groupData | Sort-Object FileName -Descending | Select-Object -First 1).FileObject.VersionInfo.FileName
        } else {
            $FilePathList += ($gServer.Group).FileObject.VersionInfo.FileName
        }
    }
    return $FilePathList
}

Function Import-MyData {
    param(
        [Parameter(Mandatory = $true)][array]$FilePaths
    )
    [System.Collections.Generic.List[System.Object]]$myData = New-Object -TypeName System.Collections.Generic.List[System.Object]

    foreach ($filePath in $FilePaths) {
        $importData = Import-Clixml -Path $filePath
        $myData.Add($importData)
    }
    return $myData
}



Function Confirm-ExchangeShell {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $false)]
        [bool]$LoadExchangeShell = $true,

        [Parameter(Mandatory = $false)]
        [scriptblock]$CatchActionFunction
    )

    begin {
        $passed = $false
        $edgeTransportKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'
        $setupKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed: LoadExchangeShell: $LoadExchangeShell | Identity: $Identity"
        $params = @{
            Identity    = $Identity
            ErrorAction = "Stop"
        }
    }
    process {
        try {
            $currentErrors = $Error.Count
            Get-ExchangeServer @params | Out-Null
            Write-Verbose "Exchange PowerShell Module already loaded."
            $passed = $true
            Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction
        } catch {
            Write-Verbose "Failed to run Get-ExchangeServer"
            Invoke-CatchActionError $CatchActionFunction

            if (-not ($LoadExchangeShell)) {
                return
            }

            #Test 32 bit process, as we can't see the registry if that is the case.
            if (-not ([System.Environment]::Is64BitProcess)) {
                Write-Warning "Open a 64 bit PowerShell process to continue"
                return
            }

            if (Test-Path "$setupKey") {
                $currentErrors = $Error.Count
                Write-Verbose "We are on Exchange 2013 or newer"

                try {
                    if (Test-Path $edgeTransportKey) {
                        Write-Verbose "We are on Exchange Edge Transport Server"
                        [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exshell.psc1" -ErrorAction Stop

                        foreach ($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn) {
                            Write-Verbose "Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name
                            Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                        }

                        Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop
                    } else {
                        Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                        Connect-ExchangeServer -Auto -ClientApplication:ManagementShell
                    }

                    Write-Verbose "Imported Module. Trying Get-Exchange Server Again"
                    Get-ExchangeServer @params | Out-Null
                    $passed = $true
                    Write-Verbose "Successfully loaded Exchange Management Shell"
                    Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction
                } catch {
                    Write-Warning "Failed to Load Exchange PowerShell Module..."
                    Invoke-CatchActionError $CatchActionFunction
                }
            } else {
                Write-Verbose "Not on an Exchange or Tools server"
            }
        }
    }
    end {

        $currentErrors = $Error.Count
        $returnObject = [PSCustomObject]@{
            ShellLoaded = $passed
            Major       = ((Get-ItemProperty -Path $setupKey -Name "MsiProductMajor" -ErrorAction SilentlyContinue).MsiProductMajor)
            Minor       = ((Get-ItemProperty -Path $setupKey -Name "MsiProductMinor" -ErrorAction SilentlyContinue).MsiProductMinor)
            Build       = ((Get-ItemProperty -Path $setupKey -Name "MsiBuildMajor" -ErrorAction SilentlyContinue).MsiBuildMajor)
            Revision    = ((Get-ItemProperty -Path $setupKey -Name "MsiBuildMinor" -ErrorAction SilentlyContinue).MsiBuildMinor)
            EdgeServer  = $passed -and (Test-Path $setupKey) -and (Test-Path $edgeTransportKey)
            ToolsOnly   = $passed -and (Test-Path $setupKey) -and (!(Test-Path $edgeTransportKey)) -and `
            ($null -eq (Get-ItemProperty -Path $setupKey -Name "Services" -ErrorAction SilentlyContinue))
            RemoteShell = $passed -and (!(Test-Path $setupKey))
        }

        Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction

        return $returnObject
    }
}
Function Invoke-ScriptLogFileLocation {
    param(
        [Parameter(Mandatory = $true)][string]$FileName,
        [Parameter(Mandatory = $false)][bool]$IncludeServerName = $false
    )
    $endName = "-{0}.txt" -f $dateTimeStringFormat

    if ($IncludeServerName) {
        $endName = "-{0}{1}" -f $Script:Server, $endName
    }

    $Script:OutputFullPath = "{0}\{1}{2}" -f $OutputFilePath, $FileName, $endName
    $Script:OutXmlFullPath = $Script:OutputFullPath.Replace(".txt", ".xml")

    if ($AnalyzeDataOnly -or
        $BuildHtmlServersReport -or
        $ScriptUpdateOnly) {
        return
    }

    $Script:ExchangeShellComputer = Confirm-ExchangeShell -Identity $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}

    if (!($Script:ExchangeShellComputer.ShellLoaded)) {
        Write-Yellow("Failed to load Exchange Shell... stopping script")
        exit
    }

    if ($Script:ExchangeShellComputer.ToolsOnly -and
        $env:COMPUTERNAME -eq $Script:Server -and
        !($LoadBalancingReport)) {
        Write-Yellow("Can't run Exchange Health Checker Against a Tools Server. Use the -Server Parameter and provide the server you want to run the script against.")
        exit
    }

    Write-VerboseWriter("Script Executing on Server $env:COMPUTERNAME")
    Write-VerboseWriter("ToolsOnly: $($Script:ExchangeShellComputer.ToolsOnly) | RemoteShell $($Script:ExchangeShellComputer.RemoteShell)")
}

Function Test-RequiresServerFqdn {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $tempServerName = ($Script:Server).Split(".")

    if ($tempServerName[0] -eq $env:COMPUTERNAME) {
        Write-Verbose "Executed against the local machine. No need to pass '-ComputerName' parameter."
        return
    } else {
        try {
            $Script:ServerFQDN = (Get-ExchangeServer $Script:Server -ErrorAction Stop).FQDN
        } catch {
            Invoke-CatchActions
            Write-Verbose "Unable to query Fqdn via 'Get-ExchangeServer'"
        }
    }

    try {
        Invoke-Command -ComputerName $Script:Server -ScriptBlock { Get-Date | Out-Null } -ErrorAction Stop
        Write-Verbose "Connected successfully using: $($Script:Server)."
    } catch {
        Invoke-CatchActions
        if ($tempServerName.Count -gt 1) {
            $Script:Server = $tempServerName[0]
        } else {
            $Script:Server = $Script:ServerFQDN
        }

        try {
            Invoke-Command -ComputerName $Script:Server -ScriptBlock { Get-Date | Out-Null } -ErrorAction Stop
            Write-Verbose "Fallback to: $($Script:Server) Connection was successfully established."
        } catch {
            Write-Red("Failed to run against: {0}. Please try to run the script locally on: {0} for results. " -f $Script:Server)
            exit
        }
    }
}

try {
    #Enums and custom data types
    Add-Type -TypeDefinition @"
    using System;
    using System.Collections;
        namespace HealthChecker
        {
            public class HealthCheckerExchangeServer
            {
                public string ServerName;        //String of the server that we are working with
                public HardwareInformation HardwareInformation;  // Hardware Object Information
                public OperatingSystemInformation  OSInformation; // OS Version Object Information
                public ExchangeInformation ExchangeInformation; //Detailed Exchange Information
                public string HealthCheckerVersion; //To determine the version of the script on the object.
                public DateTime GenerationTime; //Time stamp of running the script
            }

            // ExchangeInformation
            public class ExchangeInformation
            {
                public ExchangeBuildInformation BuildInformation = new ExchangeBuildInformation();   //Exchange build information
                public object GetExchangeServer;      //Stores the Get-ExchangeServer Object
                public object GetMailboxServer;       //Stores the Get-MailboxServer Object
                public object GetOwaVirtualDirectory; //Stores the Get-OwaVirtualDirectory Object
                public object GetWebServicesVirtualDirectory; //stores the Get-WebServicesVirtualDirectory object
                public object GetOrganizationConfig; //Stores the result from Get-OrganizationConfig
                public object msExchStorageGroup;   //Stores the properties of the 'ms-Exch-Storage-Group' Schema class
                public object GetHybridConfiguration; //Stores the Get-HybridConfiguration Object
                public bool EnableDownloadDomains = new bool(); //True if Download Domains are enabled on org level
                public ExchangeNetFrameworkInformation NETFramework = new ExchangeNetFrameworkInformation();
                public bool MapiHttpEnabled; //Stored from organization config
                public System.Array ExchangeServicesNotRunning; //Contains the Exchange services not running by Test-ServiceHealth
                public Hashtable ApplicationPools = new Hashtable();
                public ExchangeRegistryValues RegistryValues = new ExchangeRegistryValues();
                public ExchangeServerMaintenance ServerMaintenance;
                public System.Array ExchangeCertificates;           //stores all the Exchange certificates on the servers.
                public object ExchangeEmergencyMitigationService;   //stores the Exchange Emergency Mitigation Service (EEMS) object
                public Hashtable ApplicationConfigFileStatus = new Hashtable();
            }

            public class ExchangeBuildInformation
            {
                public ExchangeServerRole ServerRole; //Roles that are currently set and installed.
                public ExchangeMajorVersion MajorVersion; //Exchange Version (Exchange 2010/2013/2019)
                public ExchangeCULevel CU;             // Exchange CU Level
                public string FriendlyName;     //Exchange Friendly Name is provided
                public string BuildNumber;      //Exchange Build Number
                public string LocalBuildNumber; //Local Build Number. Is only populated if from a Tools Machine
                public string ReleaseDate;      // Exchange release date for which the CU they are currently on
                public bool SupportedBuild;     //Determines if we are within the correct build of Exchange.
                public object ExchangeSetup;    //Stores the Get-Command ExSetup object
                public System.Array KBsInstalled;  //Stored object IU or Security KB fixes
                public bool March2021SUInstalled;    //True if March 2021 SU is installed
            }

            public class ExchangeNetFrameworkInformation
            {
                public NetMajorVersion MinSupportedVersion; //Min Supported .NET Framework version
                public NetMajorVersion MaxSupportedVersion; //Max (Recommended) Supported .NET version.
                public bool OnRecommendedVersion; //RecommendedNetVersion Info includes all the factors. Windows Version & CU.
                public string DisplayWording; //Display if we are in Support or not
            }

            public class ExchangeServerMaintenance
            {
                public System.Array InactiveComponents;
                public object GetServerComponentState;
                public object GetClusterNode;
                public object GetMailboxServer;
            }

            //enum for CU levels of Exchange
            public enum ExchangeCULevel
            {
                Unknown,
                Preview,
                RTM,
                CU1,
                CU2,
                CU3,
                CU4,
                CU5,
                CU6,
                CU7,
                CU8,
                CU9,
                CU10,
                CU11,
                CU12,
                CU13,
                CU14,
                CU15,
                CU16,
                CU17,
                CU18,
                CU19,
                CU20,
                CU21,
                CU22,
                CU23
            }

            //enum for the server roles that the computer is
            public enum ExchangeServerRole
            {
                MultiRole,
                Mailbox,
                ClientAccess,
                Hub,
                Edge,
                None
            }

            //enum for the Exchange version
            public enum ExchangeMajorVersion
            {
                Unknown,
                Exchange2010,
                Exchange2013,
                Exchange2016,
                Exchange2019
            }

            public class ExchangeRegistryValues
            {
                public int CtsProcessorAffinityPercentage;    //Stores the CtsProcessorAffinityPercentage registry value from HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ExchangeServer\v15\Search\SystemParameters
                public int FipsAlgorithmPolicyEnabled;       //Stores the Enabled value from HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy
            }
            // End ExchangeInformation

            // OperatingSystemInformation
            public class OperatingSystemInformation
            {
                public OSBuildInformation BuildInformation = new OSBuildInformation(); // contains build information
                public NetworkInformation NetworkInformation = new NetworkInformation(); //stores network information and settings
                public PowerPlanInformation PowerPlan = new PowerPlanInformation(); //stores the power plan information
                public PageFileInformation PageFile;             //stores the page file information
                public LmCompatibilityLevelInformation LmCompatibility; // stores Lm Compatibility Level Information
                public ServerPendingReboot ServerPendingReboot = new ServerPendingReboot();// determines if the server is pending a reboot. TODO: Adjust to contain the registry values that we are looking at.
                public TimeZoneInformation TimeZone = new TimeZoneInformation();    //stores time zone information
                public Hashtable TLSSettings;            // stores the TLS settings on the server.
                public InstalledUpdatesInformation InstalledUpdates = new InstalledUpdatesInformation();  //store the install update
                public ServerBootUpInformation ServerBootUp = new ServerBootUpInformation();   // stores the server boot up time information
                public System.Array VcRedistributable;            //stores the Visual C++ Redistributable
                public OSNetFrameworkInformation NETFramework = new OSNetFrameworkInformation();          //stores OS Net Framework
                public bool CredentialGuardEnabled;
                public OSRegistryValues RegistryValues = new OSRegistryValues();
                public object Smb1ServerSettings;
            }

            public class ServerPendingReboot
            {
                public bool PendingFileRenameOperations;
                public object SccmReboot;
                public bool SccmRebootPending;
                public bool ComponentBasedServicingRebootPending;
                public bool AutoUpdatePendingReboot;
                public bool PendingReboot;
            }

            public class OSBuildInformation
            {
                public OSServerVersion MajorVersion; //OS Major Version
                public string VersionBuild;           //hold the build number
                public string FriendlyName;           //string holder of the Windows Server friendly name
                public object OperatingSystem;        // holds Win32_OperatingSystem
            }

            public class NetworkInformation
            {
                public double TCPKeepAlive;           // value used for the TCP/IP keep alive value in the registry
                public double RpcMinConnectionTimeout;  //holds the value for the RPC minimum connection timeout.
                public string HttpProxy;                // holds the setting for HttpProxy if one is set.
                public object PacketsReceivedDiscarded;   //hold all the packets received discarded on the server.
                public double IPv6DisabledComponents;    //value stored in the registry HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\DisabledComponents
                public bool IPv6DisabledOnNICs;          //value that determines if we have IPv6 disabled on some NICs or not.
                public System.Array NetworkAdapters;           //stores all the NICs on the servers.
                public string PnPCapabilities;      //Value from PnPCapabilities registry
                public bool SleepyNicDisabled;     //If the NIC can be in power saver mode by the OS.
            }

            public class PowerPlanInformation
            {
                public bool HighPerformanceSet;      // If the power plan is High Performance
                public string PowerPlanSetting;      //value for the power plan that is set
                public object PowerPlan;            //object to store the power plan information
            }

            public class PageFileInformation
            {
                public object PageFile;       //store the information that we got for the page file
                public double MaxPageSize;    //holds the information of what our page file is set to
            }

            public class OSRegistryValues
            {
                public int CurrentVersionUbr; // stores SOFTWARE\Microsoft\Windows NT\CurrentVersion\UBR
                public int LanManServerDisabledCompression; // stores SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\DisabledCompression
            }

            public class LmCompatibilityLevelInformation
            {
                public int RegistryValue;       //The LmCompatibilityLevel for the server (INT 1 - 5)
                public string Description;      //description of the LmCompat that the server is set to
            }

            public class TimeZoneInformation
            {
                public string CurrentTimeZone; //stores the value for the current time zone of the server.
                public int DynamicDaylightTimeDisabled; // the registry value for DynamicDaylightTimeDisabled.
                public string TimeZoneKeyName; // the registry value TimeZoneKeyName.
                public string StandardStart;   // the registry value for StandardStart.
                public string DaylightStart;   // the registry value for DaylightStart.
                public bool DstIssueDetected;  // Determines if there is a high chance of an issue.
                public System.Array ActionsToTake; //array of verbage of the issues detected.
            }

            public class ServerRebootInformation
            {
                public bool PendingFileRenameOperations;            //bool "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\" item PendingFileRenameOperations.
                public object SccmReboot;                           // object to store CimMethod for class name CCM_ClientUtilities
                public bool SccmRebootPending;                      // SccmReboot has either PendingReboot or IsHardRebootPending is set to true.
                public bool ComponentBasedServicingPendingReboot;   // bool HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending
                public bool AutoUpdatePendingReboot;                // bool HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired
                public bool PendingReboot;                         // bool if reboot types are set to true
            }

            public class InstalledUpdatesInformation
            {
                public System.Array HotFixes;     //array to keep all the hotfixes of the server
                public System.Array HotFixInfo;   //object to store hotfix information
                public System.Array InstalledUpdates; //store the install updates
            }

            public class ServerBootUpInformation
            {
                public string Days;
                public string Hours;
                public string Minutes;
                public string Seconds;
            }

            public class OSNetFrameworkInformation
            {
                public NetMajorVersion NetMajorVersion; //NetMajorVersion value
                public string FriendlyName;  //string of the friendly name
                public int RegistryValue; //store the registry value
                public Hashtable FileInformation; //stores Get-Item information for .NET Framework
            }

            //enum for the OSServerVersion that we are
            public enum OSServerVersion
            {
                Unknown,
                Windows2008,
                Windows2008R2,
                Windows2012,
                Windows2012R2,
                Windows2016,
                Windows2019,
                WindowsCore
            }

            //enum for the dword value of the .NET frame 4 that we are on
            public enum NetMajorVersion
            {
                Unknown = 0,
                Net4d5 = 378389,
                Net4d5d1 = 378675,
                Net4d5d2 = 379893,
                Net4d5d2wFix = 380035,
                Net4d6 = 393295,
                Net4d6d1 = 394254,
                Net4d6d1wFix = 394294,
                Net4d6d2 = 394802,
                Net4d7 = 460798,
                Net4d7d1 = 461308,
                Net4d7d2 = 461808,
                Net4d8 = 528040
            }
            // End OperatingSystemInformation

            // HardwareInformation
            public class HardwareInformation
            {
                public string Manufacturer; //String to display the hardware information
                public ServerType ServerType; //Enum to determine if the hardware is VMware, HyperV, Physical, or Unknown
                public System.Array MemoryInformation; //Detailed information about the installed memory
                public UInt64 TotalMemory; //Stores the total memory cooked value
                public object System;   //object to store the system information that we have collected
                public ProcessorInformation Processor;   //Detailed processor Information
                public bool AutoPageFile; //True/False if we are using a page file that is being automatically set
                public string Model; //string to display Model
            }

            //enum for the type of computer that we are
            public enum ServerType
            {
                VMWare,
                AmazonEC2,
                HyperV,
                Physical,
                Unknown
            }

            public class ProcessorInformation
            {
                public string Name;    //String of the processor name
                public int NumberOfPhysicalCores;    //Number of Physical cores that we have
                public int NumberOfLogicalCores;  //Number of Logical cores that we have presented to the os
                public int NumberOfProcessors; //Total number of processors that we have in the system
                public int MaxMegacyclesPerCore; //Max speed that we can get out of the cores
                public int CurrentMegacyclesPerCore; //Current speed that we are using the cores at
                public bool ProcessorIsThrottled;  //True/False if we are throttling our processor
                public bool DifferentProcessorsDetected; //true/false to detect if we have different processor types detected
                public bool DifferentProcessorCoreCountDetected; //detect if there are a different number of core counts per Processor CPU socket
                public int EnvironmentProcessorCount; //[system.environment]::processorcount
                public object ProcessorClassObject;        // object to store the processor information
            }

            //HTML & display classes
            public class HtmlServerValues
            {
                public System.Array OverviewValues;
                public System.Array ActionItems;   //use HtmlServerActionItemRow
                public System.Array ServerDetails;    // use HtmlServerInformationRow
            }

            public class HtmlServerActionItemRow
            {
                public string Setting;
                public string DetailValue;
                public string RecommendedDetails;
                public string MoreInformation;
                public string Class;
            }

            public class HtmlServerInformationRow
            {
                public string Name;
                public string DetailValue;
                public string Class;
            }

            public class DisplayResultsLineInfo
            {
                public string DisplayValue;
                public string Name;
                public int TabNumber;
                public object TestingValue; //Used for pester testing down the road.
                public object OutColumns; //used for colorized format table option.
                public string WriteType;

                public string Line
                {
                    get
                    {
                        if (String.IsNullOrEmpty(this.Name))
                        {
                            return this.DisplayValue;
                        }

                        return String.Concat(this.Name, ": ", this.DisplayValue);
                    }
                }
            }

            public class DisplayResultsGroupingKey
            {
                public string Name;
                public int DefaultTabNumber;
                public bool DisplayGroupName;
                public int DisplayOrder;
            }

            public class AnalyzedInformation
            {
                public HealthCheckerExchangeServer HealthCheckerExchangeServer;
                public Hashtable HtmlServerValues = new Hashtable();
                public Hashtable DisplayResults = new Hashtable();
            }
        }
"@ -ErrorAction Stop
} catch {
    Write-Warning "There was an error trying to add custom classes to the current PowerShell session. You need to close this session and open a new one to have the script properly work."
    exit
}

Function Write-ResultsToScreen {
    param(
        [Hashtable]$ResultsToWrite
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $indexOrderGroupingToKey = @{}

    foreach ($keyGrouping in $ResultsToWrite.Keys) {
        $indexOrderGroupingToKey[$keyGrouping.DisplayOrder] = $keyGrouping
    }

    $sortedIndexOrderGroupingToKey = $indexOrderGroupingToKey.Keys | Sort-Object

    foreach ($key in $sortedIndexOrderGroupingToKey) {
        Write-Verbose "Working on Key: $key"
        $keyGrouping = $indexOrderGroupingToKey[$key]
        Write-Verbose "Working on Key Group: $($keyGrouping.Name)"
        Write-Verbose "Total lines to write: $($ResultsToWrite[$keyGrouping].Count)"

        if ($keyGrouping.DisplayGroupName) {
            Write-Grey($keyGrouping.Name)
            $dashes = [string]::empty
            1..($keyGrouping.Name.Length) | ForEach-Object { $dashes = $dashes + "-" }
            Write-Grey($dashes)
        }

        foreach ($line in $ResultsToWrite[$keyGrouping]) {
            $tab = [string]::Empty

            if ($line.TabNumber -ne 0) {
                1..($line.TabNumber) | ForEach-Object { $tab = $tab + "`t" }
            }

            $writeValue = "{0}{1}" -f $tab, $line.Line
            switch ($line.WriteType) {
                "Grey" { Write-Grey($writeValue) }
                "Yellow" { Write-Yellow($writeValue) }
                "Green" { Write-Green($writeValue) }
                "Red" { Write-Red($writeValue) }
                "OutColumns" { Write-OutColumns($line.OutColumns) }
            }
        }

        Write-Grey("")
    }
}

Function Write-Verbose {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'In order to log Write-Verbose from Shared functions')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [string]$Message
    )

    process {
        #write to the debug log and call Write-Verbose normally
        Write-DebugLog $Message
        Microsoft.PowerShell.Utility\Write-Verbose $Message
    }
}


<#
.SYNOPSIS
    Outputs a table of objects with certain values colorized.
.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>
function Out-Columns {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object[]]
        $InputObject,

        [Parameter(Mandatory = $false, Position = 0)]
        [string[]]
        $Properties,

        [Parameter(Mandatory = $false, Position = 1)]
        [scriptblock[]]
        $ColorizerFunctions = @(),

        [Parameter(Mandatory = $false)]
        [int]
        $IndentSpaces = 0,

        [Parameter(Mandatory = $false)]
        [int]
        $LinesBetweenObjects = 0,

        [Parameter(Mandatory = $false)]
        [ref]
        $StringOutput
    )

    begin {
        function WrapLine {
            param([string]$line, [int]$width)
            if ($line.Length -le $width -and $line.IndexOf("`n") -lt 0) {
                return $line
            }

            $lines = New-Object System.Collections.ArrayList

            $noLF = $line.Replace("`r", "")
            $lineSplit = $noLF.Split("`n")
            foreach ($l in $lineSplit) {
                if ($l.Length -le $width) {
                    [void]$lines.Add($l)
                } else {
                    $split = $l.Split(" ")
                    $sb = New-Object System.Text.StringBuilder
                    for ($i = 0; $i -lt $split.Length; $i++) {
                        if ($sb.Length -eq 0 -and $sb.Length + $split[$i].Length -lt $width) {
                            [void]$sb.Append($split[$i])
                        } elseif ($sb.Length -gt 0 -and $sb.Length + $split[$i].Length + 1 -lt $width) {
                            [void]$sb.Append(" " + $split[$i])
                        } elseif ($sb.Length -gt 0) {
                            [void]$lines.Add($sb.ToString())
                            [void]$sb.Clear()
                            $i--
                        } else {
                            if ($split[$i].Length -lt $width) {
                                [void]$lines.Add($split[$i])
                            } else {
                                [void]$lines.Add($split[$i].Substring(0, $width))
                                $split[$i] = $split[$i].Substring($width + 1)
                                $i--
                            }
                        }
                    }

                    if ($sb.Length -gt 0) {
                        [void]$lines.Add($sb.ToString())
                    }
                }
            }

            return $lines
        }

        function GetLineObjects {
            param($obj, $props, $colWidths)
            $linesNeededForThisObject = 1
            $multiLineProps = @{}
            for ($i = 0; $i -lt $props.Length; $i++) {
                $p = $props[$i]
                $val = $obj."$p"

                if ($val -isnot [array]) {
                    $val = WrapLine -line $val -width $colWidths[$i]
                } elseif ($val -is [array]) {
                    $val = $val | Where-Object { $null -ne $_ }
                    $val = $val | ForEach-Object { WrapLine -line $_ -width $colWidths[$i] }
                }

                if ($val -is [array]) {
                    $multiLineProps[$p] = $val
                    if ($val.Length -gt $linesNeededForThisObject) {
                        $linesNeededForThisObject = $val.Length
                    }
                }
            }

            if ($linesNeededForThisObject -eq 1) {
                $obj
            } else {
                for ($i = 0; $i -lt $linesNeededForThisObject; $i++) {
                    $lineProps = @{}
                    foreach ($p in $props) {
                        if ($null -ne $multiLineProps[$p] -and $multiLineProps[$p].Length -gt $i) {
                            $lineProps[$p] = $multiLineProps[$p][$i]
                        } elseif ($i -eq 0) {
                            $lineProps[$p] = $o."$p"
                        } else {
                            $lineProps[$p] = $null
                        }
                    }

                    [PSCustomObject]$lineProps
                }
            }
        }

        function GetColumnColors {
            param($obj, $props, $funcs)

            $consoleHost = (Get-Host).Name -eq "ConsoleHost"
            $colColors = New-Object string[] $props.Count
            for ($i = 0; $i -lt $props.Count; $i++) {
                if ($consoleHost) {
                    $fgColor = (Get-Host).ui.rawui.ForegroundColor
                } else {
                    $fgColor = "White"
                }
                foreach ($func in $funcs) {
                    $result = $func.Invoke($o, $props[$i])
                    if (-not [string]::IsNullOrEmpty($result)) {
                        $fgColor = $result
                        break # The first colorizer that takes action wins
                    }
                }

                $colColors[$i] = $fgColor
            }

            $colColors
        }

        function GetColumnWidths {
            param($objects, $props)

            $colWidths = New-Object int[] $props.Count

            # Start with the widths of the property names
            for ($i = 0; $i -lt $props.Count; $i++) {
                $colWidths[$i] = $props[$i].Length
            }

            # Now check the widths of the widest values
            foreach ($thing in $objects) {
                for ($i = 0; $i -lt $props.Count; $i++) {
                    $val = $thing."$($props[$i])"
                    if ($null -ne $val) {
                        $width = 0
                        if ($val -isnot [array]) {
                            $val = $val.ToString().Split("`n")
                        }

                        $width = ($val | ForEach-Object {
                                if ($null -ne $_) { $_.ToString() } else { "" }
                            } | Sort-Object Length -Descending | Select-Object -First 1).Length

                        if ($width -gt $colWidths[$i]) {
                            $colWidths[$i] = $width
                        }
                    }
                }
            }

            # If we're within the window width, we're done
            $totalColumnWidth = $colWidths.Length * $padding + ($colWidths | Measure-Object -Sum).Sum + $IndentSpaces
            $windowWidth = (Get-Host).UI.RawUI.WindowSize.Width
            if ($windowWidth -lt 1 -or $totalColumnWidth -lt $windowWidth) {
                return $colWidths
            }

            # Take size away from one or more columns to make them fit
            while ($totalColumnWidth -ge $windowWidth) {
                $startingTotalWidth = $totalColumnWidth
                $widest = $colWidths | Sort-Object -Descending | Select-Object -First 1
                $newWidest = [Math]::Floor($widest * 0.95)
                for ($i = 0; $i -lt $colWidths.Length; $i++) {
                    if ($colWidths[$i] -eq $widest) {
                        $colWidths[$i] = $newWidest
                        break
                    }
                }

                $totalColumnWidth = $colWidths.Length * $padding + ($colWidths | Measure-Object -Sum).Sum + $IndentSpaces
                if ($totalColumnWidth -ge $startingTotalWidth) {
                    # Somehow we didn't reduce the size at all, so give up
                    break
                }
            }

            return $colWidths
        }

        $objects = New-Object System.Collections.ArrayList
        $padding = 2
        $stb = New-Object System.Text.StringBuilder
    }

    process {
        foreach ($thing in $InputObject) {
            [void]$objects.Add($thing)
        }
    }

    end {
        if ($objects.Count -gt 0) {
            $props = $null

            if ($null -ne $Properties) {
                $props = $Properties
            } else {
                $props = $objects[0].PSObject.Properties.Name
            }

            $colWidths = GetColumnWidths $objects $props

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            Write-Host (" " * $IndentSpaces) -NoNewline
            [void]$stb.Append(" " * $IndentSpaces)

            for ($i = 0; $i -lt $props.Count; $i++) {
                Write-Host ("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $props[$i]) -NoNewline
                [void]$stb.Append("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $props[$i])
            }

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            Write-Host (" " * $IndentSpaces) -NoNewline
            [void]$stb.Append(" " * $IndentSpaces)

            for ($i = 0; $i -lt $props.Count; $i++) {
                Write-Host ("{0,$(-1 * ($colWidths[$i] + $padding))}" -f ("-" * $props[$i].Length)) -NoNewline
                [void]$stb.Append("{0,$(-1 * ($colWidths[$i] + $padding))}" -f ("-" * $props[$i].Length))
            }

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            foreach ($o in $objects) {
                $colColors = GetColumnColors -obj $o -props $props -funcs $ColorizerFunctions
                $lineObjects = @(GetLineObjects -obj $o -props $props -colWidths $colWidths)
                foreach ($lineObj in $lineObjects) {
                    Write-Host (" " * $IndentSpaces) -NoNewline
                    [void]$stb.Append(" " * $IndentSpaces)
                    for ($i = 0; $i -lt $props.Count; $i++) {
                        $val = $o."$($props[$i])"
                        Write-Host ("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $lineObj."$($props[$i])") -NoNewline -ForegroundColor $colColors[$i]
                        [void]$stb.Append("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $lineObj."$($props[$i])")
                    }

                    Write-Host
                    [void]$stb.Append([System.Environment]::NewLine)
                }

                for ($i = 0; $i -lt $LinesBetweenObjects; $i++) {
                    Write-Host
                    [void]$stb.Append([System.Environment]::NewLine)
                }
            }

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            if ($null -ne $StringOutput) {
                $StringOutput.Value = $stb.ToString()
            }
        }
    }
}
function Write-Red($message) {
    Write-DebugLog $message
    Write-Host $message -ForegroundColor Red
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Yellow($message) {
    Write-DebugLog $message
    Write-Host $message -ForegroundColor Yellow
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Green($message) {
    Write-DebugLog $message
    Write-Host $message -ForegroundColor Green
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Grey($message) {
    Write-DebugLog $message
    Write-Host $message
    $message | Out-File ($OutputFullPath) -Append
}

function Write-DebugLog($message) {
    if (![string]::IsNullOrEmpty($message)) {
        $Script:Logger.WriteToFileOnly($message)
    }
}

function Write-OutColumns($OutColumns) {
    if ($null -ne $OutColumns) {
        $stringOutput = $null
        $OutColumns.DisplayObject |
            Out-Columns -Properties $OutColumns.SelectProperties `
                -ColorizerFunctions $OutColumns.ColorizerFunctions `
                -IndentSpaces $OutColumns.IndentSpaces `
                -StringOutput ([ref]$stringOutput)
        $stringOutput | Out-File ($OutputFullPath) -Append
        Write-DebugLog $stringOutput
    }
}

Function Write-Break {
    Write-Host ""
}

Function Get-HtmlServerReport {
    param(
        [Parameter(Mandatory = $true)][array]$AnalyzedHtmlServerValues
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $htmlHeader = "<html>
        <style>
        BODY{font-family: Arial; font-size: 8pt;}
        H1{font-size: 16px;}
        H2{font-size: 14px;}
        H3{font-size: 12px;}
        TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
        TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
        TD{border: 1px solid black; padding: 5px; }
        td.Green{background: #7FFF00;}
        td.Yellow{background: #FFE600;}
        td.Red{background: #FF0000; color: #ffffff;}
        td.Info{background: #85D4FF;}
        </style>
        <body>
        <h1 align=""center"">Exchange Health Checker v$($BuildVersion)</h1><br>
        <h2>Servers Overview</h2>"

    [array]$htmlOverviewTable += "<p>
        <table>
        <tr>"

    foreach ($tableHeaderName in $AnalyzedHtmlServerValues[0]["OverviewValues"].Name) {
        $htmlOverviewTable += "<th>{0}</th>" -f $tableHeaderName
    }

    $htmlOverviewTable += "</tr>"

    foreach ($serverHtmlServerValues in $AnalyzedHtmlServerValues) {
        $htmlTableRow = @()
        [array]$htmlTableRow += "<tr>"
        foreach ($htmlTableDataRow in $serverHtmlServerValues["OverviewValues"]) {
            $htmlTableRow += "<td class=`"{0}`">{1}</td>" -f $htmlTableDataRow.Class, `
                $htmlTableDataRow.DetailValue
        }

        $htmlTableRow += "</tr>"
        $htmlOverviewTable += $htmlTableRow
    }

    $htmlOverviewTable += "</table></p>"

    [array]$htmlServerDetails += "<p><h2>Server Details</h2><table>"

    foreach ($serverHtmlServerValues in $AnalyzedHtmlServerValues) {
        foreach ($htmlTableDataRow in $serverHtmlServerValues["ServerDetails"]) {
            if ($htmlTableDataRow.Name -eq "Server Name") {
                $htmlServerDetails += "<tr><th>{0}</th><th>{1}</th><tr>" -f $htmlTableDataRow.Name, `
                    $htmlTableDataRow.DetailValue
            } else {
                $htmlServerDetails += "<tr><td class=`"{0}`">{1}</td><td class=`"{0}`">{2}</td><tr>" -f $htmlTableDataRow.Class, `
                    $htmlTableDataRow.Name, `
                    $htmlTableDataRow.DetailValue
            }
        }
    }
    $htmlServerDetails += "</table></p>"

    $htmlReport = $htmlHeader + $htmlOverviewTable + $htmlServerDetails + "</body></html>"

    $htmlReport | Out-File $HtmlReportFile -Encoding UTF8
}

Function Get-CASLoadBalancingReport {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    #Connection and requests per server and client type values
    $CASConnectionStats = @{}
    $TotalCASConnectionCount = 0
    $AutoDStats = @{}
    $TotalAutoDRequests = 0
    $EWSStats = @{}
    $TotalEWSRequests = 0
    $MapiHttpStats = @{}
    $TotalMapiHttpRequests = 0
    $EASStats = @{}
    $TotalEASRequests = 0
    $OWAStats = @{}
    $TotalOWARequests = 0
    $RpcHttpStats = @{}
    $TotalRpcHttpRequests = 0
    $CASServers = @()

    if ($null -ne $CasServerList) {
        Write-Grey("Custom CAS server list is being used.  Only servers specified after the -CasServerList parameter will be used in the report.")
        foreach ($cas in $CasServerList) {
            $CASServers += (Get-ExchangeServer $cas)
        }
    } elseif ($SiteName -ne [string]::Empty) {
        Write-Grey("Site filtering ON.  Only Exchange 2013/2016 CAS servers in {0} will be used in the report." -f $SiteName)
        $CASServers = Get-ExchangeServer | Where-Object { `
            ($_.IsClientAccessServer -eq $true) -and `
            ($_.AdminDisplayVersion -Match "^Version 15") -and `
            ([System.Convert]::ToString($_.Site).Split("/")[-1] -eq $SiteName) }
    } else {
        Write-Grey("Site filtering OFF.  All Exchange 2013/2016 CAS servers will be used in the report.")
        $CASServers = Get-ExchangeServer | Where-Object { ($_.IsClientAccessServer -eq $true) -and ($_.AdminDisplayVersion -Match "^Version 15") }
    }

    if ($CASServers.Count -eq 0) {
        Write-Red("Error: No CAS servers found using the specified search criteria.")
        Exit
    }

    #Request stats from perfmon for all CAS
    $PerformanceCounters = @()
    $PerformanceCounters += "\Web Service(Default Web Site)\Current Connections"
    $PerformanceCounters += "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_Autodiscover)\Requests Executing"
    $PerformanceCounters += "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_EWS)\Requests Executing"
    $PerformanceCounters += "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_mapi)\Requests Executing"
    $PerformanceCounters += "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_Microsoft-Server-ActiveSync)\Requests Executing"
    $PerformanceCounters += "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_owa)\Requests Executing"
    $PerformanceCounters += "\ASP.NET Apps v4.0.30319(_LM_W3SVC_1_ROOT_Rpc)\Requests Executing"
    $currentErrors = $Error.Count
    $AllCounterResults = Get-Counter -ComputerName $CASServers -Counter $PerformanceCounters -ErrorAction SilentlyContinue

    if ($currentErrors -ne $Error.Count) {
        $i = 0
        while ($i -lt ($Error.Count - $currentErrors)) {
            Invoke-CatchActions -CopyThisError $Error[$i]
            $i++
        }

        Write-VerboseWriter("Failed to get some counters")
    }

    foreach ($Result in $AllCounterResults.CounterSamples) {
        $CasName = ($Result.Path).Split("\\", [System.StringSplitOptions]::RemoveEmptyEntries)[0]
        $ResultCookedValue = $Result.CookedValue

        if ($Result.Path -like "*{0}*{1}" -f $CasName, $PerformanceCounters[0]) {
            #Total connections
            $CASConnectionStats.Add($CasName, $ResultCookedValue)
            $TotalCASConnectionCount += $ResultCookedValue
        } elseif ($Result.Path -like "*{0}*{1}" -f $CasName, $PerformanceCounters[1]) {
            #AutoD requests
            $AutoDStats.Add($CasName, $ResultCookedValue)
            $TotalAutoDRequests += $ResultCookedValue
        } elseif ($Result.Path -like "*{0}*{1}" -f $CasName, $PerformanceCounters[2]) {
            #EWS requests
            $EWSStats.Add($CasName, $ResultCookedValue)
            $TotalEWSRequests += $ResultCookedValue
        } elseif ($Result.Path -like "*{0}*{1}" -f $CasName, $PerformanceCounters[3]) {
            #MapiHttp requests
            $MapiHttpStats.Add($CasName, $ResultCookedValue)
            $TotalMapiHttpRequests += $ResultCookedValue
        } elseif ($Result.Path -like "*{0}*{1}" -f $CasName, $PerformanceCounters[4]) {
            #EAS requests
            $EASStats.Add($CasName, $ResultCookedValue)
            $TotalEASRequests += $ResultCookedValue
        } elseif ($Result.Path -like "*{0}*{1}" -f $CasName, $PerformanceCounters[5]) {
            #OWA requests
            $OWAStats.Add($CasName, $ResultCookedValue)
            $TotalOWARequests += $ResultCookedValue
        } elseif ($Result.Path -like "*{0}*{1}" -f $CasName, $PerformanceCounters[6]) {
            #RPCHTTP requests
            $RpcHttpStats.Add($CasName, $ResultCookedValue)
            $TotalRpcHttpRequests += $ResultCookedValue
        }
    }


    #Report the results for connection count
    Write-Grey("")
    Write-Grey("Connection Load Distribution Per Server")
    Write-Grey("Total Connections: {0}" -f $TotalCASConnectionCount)
    #Calculate percentage of connection load
    $CASConnectionStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
        Write-Grey($_.Key + ": " + $_.Value + " Connections = " + [math]::Round((([int]$_.Value / $TotalCASConnectionCount) * 100)) + "% Distribution")
    }

    #Same for each client type.  These are request numbers not connection numbers.
    #AutoD
    if ($TotalAutoDRequests -gt 0) {
        Write-Grey("")
        Write-Grey("Current AutoDiscover Requests Per Server")
        Write-Grey("Total Requests: {0}" -f $TotalAutoDRequests)
        $AutoDStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
            Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value / $TotalAutoDRequests) * 100)) + "% Distribution")
        }
    }

    #EWS
    if ($TotalEWSRequests -gt 0) {
        Write-Grey("")
        Write-Grey("Current EWS Requests Per Server")
        Write-Grey("Total Requests: {0}" -f $TotalEWSRequests)
        $EWSStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
            Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value / $TotalEWSRequests) * 100)) + "% Distribution")
        }
    }

    #MapiHttp
    if ($TotalMapiHttpRequests -gt 0) {
        Write-Grey("")
        Write-Grey("Current MapiHttp Requests Per Server")
        Write-Grey("Total Requests: {0}" -f $TotalMapiHttpRequests)
        $MapiHttpStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
            Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value / $TotalMapiHttpRequests) * 100)) + "% Distribution")
        }
    }

    #EAS
    if ($TotalEASRequests -gt 0) {
        Write-Grey("")
        Write-Grey("Current EAS Requests Per Server")
        Write-Grey("Total Requests: {0}" -f $TotalEASRequests)
        $EASStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
            Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value / $TotalEASRequests) * 100)) + "% Distribution")
        }
    }

    #OWA
    if ($TotalOWARequests -gt 0) {
        Write-Grey("")
        Write-Grey("Current OWA Requests Per Server")
        Write-Grey("Total Requests: {0}" -f $TotalOWARequests)
        $OWAStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
            Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value / $TotalOWARequests) * 100)) + "% Distribution")
        }
    }

    #RpcHttp
    if ($TotalRpcHttpRequests -gt 0) {
        Write-Grey("")
        Write-Grey("Current RpcHttp Requests Per Server")
        Write-Grey("Total Requests: {0}" -f $TotalRpcHttpRequests)
        $RpcHttpStats.GetEnumerator() | Sort-Object -Descending | ForEach-Object {
            Write-Grey($_.Key + ": " + $_.Value + " Requests = " + [math]::Round((([int]$_.Value / $TotalRpcHttpRequests) * 100)) + "% Distribution")
        }
    }
    Write-Grey("")
}

Function Get-ComputerCoresObject {
    param(
        [Parameter(Mandatory = $true)][string]$Machine_Name
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand) Passed: $Machine_Name"

    $returnObj = New-Object PSCustomObject
    $returnObj | Add-Member -MemberType NoteProperty -Name Error -Value $false
    $returnObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Machine_Name
    $returnObj | Add-Member -MemberType NoteProperty -Name NumberOfCores -Value ([int]::empty)
    $returnObj | Add-Member -MemberType NoteProperty -Name Exception -Value ([string]::empty)
    $returnObj | Add-Member -MemberType NoteProperty -Name ExceptionType -Value ([string]::empty)

    try {
        $wmi_obj_processor = Get-WmiObjectHandler -ComputerName $Machine_Name -Class "Win32_Processor" -CatchActionFunction ${Function:Invoke-CatchActions}

        foreach ($processor in $wmi_obj_processor) {
            $returnObj.NumberOfCores += $processor.NumberOfCores
        }

        Write-Grey("Server {0} Cores: {1}" -f $Machine_Name, $returnObj.NumberOfCores)
    } catch {
        Invoke-CatchActions
        $thisError = $Error[0]

        if ($thisError.Exception.Gettype().FullName -eq "System.UnauthorizedAccessException") {
            Write-Yellow("Unable to get processor information from server {0}. You do not have the correct permissions to get this data from that server. Exception: {1}" -f $Machine_Name, $thisError.ToString())
        } else {
            Write-Yellow("Unable to get processor information from server {0}. Reason: {1}" -f $Machine_Name, $thisError.ToString())
        }
        $returnObj.Exception = $thisError.ToString()
        $returnObj.ExceptionType = $thisError.Exception.Gettype().FullName
        $returnObj.Error = $true
    }

    return $returnObj
}

Function Get-ExchangeDCCoreRatio {

    Invoke-ScriptLogFileLocation -FileName "HealthChecker-ExchangeDCCoreRatio"
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    Write-Grey("Exchange Server Health Checker Report - AD GC Core to Exchange Server Core Ratio - v{0}" -f $BuildVersion)
    $coreRatioObj = New-Object PSCustomObject

    try {
        Write-Verbose "Attempting to load Active Directory Module"
        Import-Module ActiveDirectory
        Write-Verbose "Successfully loaded"
    } catch {
        Write-Red("Failed to load Active Directory Module. Stopping the script")
        exit
    }

    $ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name
    [array]$DomainControllers = (Get-ADForest).Domains | ForEach-Object { Get-ADDomainController -Filter { isGlobalCatalog -eq $true -and Site -eq $ADSite } -Server $_ }

    [System.Collections.Generic.List[System.Object]]$DCList = New-Object System.Collections.Generic.List[System.Object]
    $DCCoresTotal = 0
    Write-Break
    Write-Grey("Collecting data for the Active Directory Environment in Site: {0}" -f $ADSite)
    $iFailedDCs = 0

    foreach ($DC in $DomainControllers) {
        $DCCoreObj = Get-ComputerCoresObject -Machine_Name $DC.Name
        $DCList.Add($DCCoreObj)

        if (-not ($DCCoreObj.Error)) {
            $DCCoresTotal += $DCCoreObj.NumberOfCores
        } else {
            $iFailedDCs++
        }
    }

    $coreRatioObj | Add-Member -MemberType NoteProperty -Name DCList -Value $DCList

    if ($iFailedDCs -eq $DomainControllers.count) {
        #Core count is going to be 0, no point to continue the script
        Write-Red("Failed to collect data from your DC servers in site {0}." -f $ADSite)
        Write-Yellow("Because we can't determine the ratio, we are going to stop the script. Verify with the above errors as to why we failed to collect the data and address the issue, then run the script again.")
        exit
    }

    [array]$ExchangeServers = Get-ExchangeServer | Where-Object { $_.Site -match $ADSite }
    $EXCoresTotal = 0
    [System.Collections.Generic.List[System.Object]]$EXList = New-Object System.Collections.Generic.List[System.Object]
    Write-Break
    Write-Grey("Collecting data for the Exchange Environment in Site: {0}" -f $ADSite)
    foreach ($svr in $ExchangeServers) {
        $EXCoreObj = Get-ComputerCoresObject -Machine_Name $svr.Name
        $EXList.Add($EXCoreObj)

        if (-not ($EXCoreObj.Error)) {
            $EXCoresTotal += $EXCoreObj.NumberOfCores
        }
    }
    $coreRatioObj | Add-Member -MemberType NoteProperty -Name ExList -Value $EXList

    Write-Break
    $CoreRatio = $EXCoresTotal / $DCCoresTotal
    Write-Grey("Total DC/GC Cores: {0}" -f $DCCoresTotal)
    Write-Grey("Total Exchange Cores: {0}" -f $EXCoresTotal)
    Write-Grey("You have {0} Exchange Cores for every Domain Controller Global Catalog Server Core" -f $CoreRatio)

    if ($CoreRatio -gt 8) {
        Write-Break
        Write-Red("Your Exchange to Active Directory Global Catalog server's core ratio does not meet the recommended guidelines of 8:1")
        Write-Red("Recommended guidelines for Exchange 2013/2016 for every 8 Exchange cores you want at least 1 Active Directory Global Catalog Core.")
        Write-Yellow("Documentation:")
        Write-Yellow("`thttps://aka.ms/HC-PerfSize")
        Write-Yellow("`thttps://aka.ms/HC-ADCoreCount")
    } else {
        Write-Break
        Write-Green("Your Exchange Environment meets the recommended core ratio of 8:1 guidelines.")
    }

    $XMLDirectoryPath = $OutputFullPath.Replace(".txt", ".xml")
    $coreRatioObj | Export-Clixml $XMLDirectoryPath
    Write-Grey("Output file written to {0}" -f $OutputFullPath)
    Write-Grey("Output XML Object file written to {0}" -f $XMLDirectoryPath)
}

Function Get-MailboxDatabaseAndMailboxStatistics {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $AllDBs = Get-MailboxDatabaseCopyStatus -server $Script:Server -ErrorAction SilentlyContinue
    $MountedDBs = $AllDBs | Where-Object { $_.ActiveCopy -eq $true }

    if ($MountedDBs.Count -gt 0) {
        Write-Grey("`tActive Database:")
        foreach ($db in $MountedDBs) {
            Write-Grey("`t`t" + $db.Name)
        }
        $MountedDBs.DatabaseName | ForEach-Object { Write-Verbose "Calculating User Mailbox Total for Active Database: $_"; $TotalActiveUserMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited).Count }
        Write-Grey("`tTotal Active User Mailboxes on server: " + $TotalActiveUserMailboxCount)
        $MountedDBs.DatabaseName | ForEach-Object { Write-Verbose "Calculating Public Mailbox Total for Active Database: $_"; $TotalActivePublicFolderMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited -PublicFolder).Count }
        Write-Grey("`tTotal Active Public Folder Mailboxes on server: " + $TotalActivePublicFolderMailboxCount)
        Write-Grey("`tTotal Active Mailboxes on server " + $Script:Server + ": " + ($TotalActiveUserMailboxCount + $TotalActivePublicFolderMailboxCount).ToString())
    } else {
        Write-Grey("`tNo Active Mailbox Databases found on server " + $Script:Server + ".")
    }

    $HealthyDbs = $AllDBs | Where-Object { $_.Status -match 'Healthy' }

    if ($HealthyDbs.count -gt 0) {
        Write-Grey("`r`n`tPassive Databases:")
        foreach ($db in $HealthyDbs) {
            Write-Grey("`t`t" + $db.Name)
        }
        $HealthyDbs.DatabaseName | ForEach-Object { Write-Verbose "`tCalculating User Mailbox Total for Passive Healthy Databases: $_"; $TotalPassiveUserMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited).Count }
        Write-Grey("`tTotal Passive user Mailboxes on Server: " + $TotalPassiveUserMailboxCount)
        $HealthyDbs.DatabaseName | ForEach-Object { Write-Verbose "`tCalculating Passive Mailbox Total for Passive Healthy Databases: $_"; $TotalPassivePublicFolderMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited -PublicFolder).Count }
        Write-Grey("`tTotal Passive Public Mailboxes on server: " + $TotalPassivePublicFolderMailboxCount)
        Write-Grey("`tTotal Passive Mailboxes on server: " + ($TotalPassiveUserMailboxCount + $TotalPassivePublicFolderMailboxCount).ToString())
    } else {
        Write-Grey("`tNo Passive Mailboxes found on server " + $Script:Server + ".")
    }
}

#TODO: Address this

#https://github.com/dpaulson45/PublicPowerShellFunctions/blob/master/src/Common/Write-HostWriters/Write-ScriptMethodHostWriter.ps1
#v21.01.22.2212
Function Write-ScriptMethodHostWriter {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Need to use Write Host')]
    param(
        [Parameter(Mandatory = $true)][string]$WriteString
    )
    if ($null -ne $this.LoggerObject) {
        $this.LoggerObject.WriteHost($WriteString)
    } elseif ($null -eq $this.HostFunctionCaller) {
        Write-Host $WriteString
    } else {
        $this.HostFunctionCaller($WriteString)
    }
}

Function Write-VerboseWriter {
    param(
        [Parameter(Mandatory = $true)][string]$WriteString
    )
    if ($null -ne $Script:Logger) {
        $Script:Logger.WriteVerbose($WriteString)
    } elseif ($null -eq $VerboseFunctionCaller) {
        Write-Verbose $WriteString
    } else {
        &$VerboseFunctionCaller $WriteString
    }
}

#https://github.com/dpaulson45/PublicPowerShellFunctions/blob/master/src/Common/Write-VerboseWriters/Write-ScriptMethodVerboseWriter.ps1
#v21.01.22.2212
Function Write-ScriptMethodVerboseWriter {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Need to use Write Host')]
    param(
        [Parameter(Mandatory = $true)][string]$WriteString
    )
    if ($null -ne $this.LoggerObject) {
        $this.LoggerObject.WriteVerbose($WriteString)
    } elseif ($null -eq $this.VerboseFunctionCaller -and
        $this.WriteVerboseData) {
        Write-Host $WriteString -ForegroundColor Cyan
    } elseif ($this.WriteVerboseData) {
        $this.VerboseFunctionCaller($WriteString)
    }
}


#https://github.com/dpaulson45/PublicPowerShellFunctions/blob/master/src/Common/Confirm-Administrator/Confirm-Administrator.ps1
#v21.01.22.2212
Function Confirm-Administrator {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

    if ($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
        return $true
    } else {
        return $false
    }
}

#https://github.com/dpaulson45/PublicPowerShellFunctions/blob/master/src/Common/New-LoggerObject/New-LoggerObject.ps1
#v21.01.22.2234
Function New-LoggerObject {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'I prefer New here')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$LogDirectory = ".",
        [Parameter(Mandatory = $false)][string]$LogName = "Script_Logging",
        [Parameter(Mandatory = $false)][bool]$EnableDateTime = $true,
        [Parameter(Mandatory = $false)][bool]$IncludeDateTimeToFileName = $true,
        [Parameter(Mandatory = $false)][int]$MaxFileSizeInMB = 10,
        [Parameter(Mandatory = $false)][int]$CheckSizeIntervalMinutes = 10,
        [Parameter(Mandatory = $false)][int]$NumberOfLogsToKeep = 10,
        [Parameter(Mandatory = $false)][bool]$VerboseEnabled,
        [Parameter(Mandatory = $false)][scriptblock]$HostFunctionCaller,
        [Parameter(Mandatory = $false)][scriptblock]$VerboseFunctionCaller
    )
    #Function Version #v21.01.22.2234

    ########################
    #
    # Template Functions
    #
    ########################

    Function Write-ToLog {
        param(
            [object]$WriteString,
            [string]$LogLocation
        )
        $WriteString | Out-File ($LogLocation) -Append
    }

    ########################
    #
    # End Template Functions
    #
    ########################

    ########## Parameter Binding Exceptions ##############
    if ($LogDirectory -eq ".") {
        $LogDirectory = (Get-Location).Path
    }

    if ([string]::IsNullOrWhiteSpace($LogName)) {
        throw [System.Management.Automation.ParameterBindingException] "Failed to provide valid LogName"
    }

    if (!(Test-Path $LogDirectory)) {

        try {
            New-Item -Path $LogDirectory -ItemType Directory | Out-Null
        } catch {
            throw [System.Management.Automation.ParameterBindingException] "Failed to provide valid LogDirectory"
        }
    }

    $loggerObject = New-Object PSCustomObject
    $loggerObject | Add-Member -MemberType NoteProperty -Name "FileDirectory" -Value $LogDirectory
    $loggerObject | Add-Member -MemberType NoteProperty -Name "FileName" -Value $LogName
    $loggerObject | Add-Member -MemberType NoteProperty -Name "FullPath" -Value $fullLogPath
    $loggerObject | Add-Member -MemberType NoteProperty -Name "InstanceBaseName" -Value ([string]::Empty)
    $loggerObject | Add-Member -MemberType NoteProperty -Name "EnableDateTime" -Value $EnableDateTime
    $loggerObject | Add-Member -MemberType NoteProperty -Name "IncludeDateTimeToFileName" -Value $IncludeDateTimeToFileName
    $loggerObject | Add-Member -MemberType NoteProperty -Name "MaxFileSizeInMB" -Value $MaxFileSizeInMB
    $loggerObject | Add-Member -MemberType NoteProperty -Name "CheckSizeIntervalMinutes" -Value $CheckSizeIntervalMinutes
    $loggerObject | Add-Member -MemberType NoteProperty -Name "NextFileCheckTime" -Value ((Get-Date).AddMinutes($CheckSizeIntervalMinutes))
    $loggerObject | Add-Member -MemberType NoteProperty -Name "InstanceNumber" -Value 1
    $loggerObject | Add-Member -MemberType NoteProperty -Name "NumberOfLogsToKeep" -Value $NumberOfLogsToKeep
    $loggerObject | Add-Member -MemberType NoteProperty -Name "WriteVerboseData" -Value $VerboseEnabled
    $loggerObject | Add-Member -MemberType NoteProperty -Name "PreventLogCleanup" -Value $false
    $loggerObject | Add-Member -MemberType NoteProperty -Name "LoggerDisabled" -Value $false
    $loggerObject | Add-Member -MemberType ScriptMethod -Name "ToLog" -Value ${Function:Write-ToLog}
    $loggerObject | Add-Member -MemberType ScriptMethod -Name "WriteHostWriter" -Value ${Function:Write-ScriptMethodHostWriter}
    $loggerObject | Add-Member -MemberType ScriptMethod -Name "WriteVerboseWriter" -Value ${Function:Write-ScriptMethodVerboseWriter}

    if ($null -ne $HostFunctionCaller) {
        $loggerObject | Add-Member -MemberType ScriptMethod -Name "HostFunctionCaller" -Value $HostFunctionCaller
    }

    if ($null -ne $VerboseFunctionCaller) {
        $loggerObject | Add-Member -MemberType ScriptMethod -Name "VerboseFunctionCaller" -Value $VerboseFunctionCaller
    }

    $loggerObject | Add-Member -MemberType ScriptMethod -Name "WriteHost" -Value {
        param(
            [object]$LoggingString
        )

        if ($this.LoggerDisabled) {
            return
        }

        if ($null -eq $LoggingString) {
            throw [System.Management.Automation.ParameterBindingException] "Failed to provide valid LoggingString"
        }

        if ($this.EnableDateTime) {
            $LoggingString = "[{0}] : {1}" -f [System.DateTime]::Now, $LoggingString
        }

        $this.WriteHostWriter($LoggingString)
        $this.ToLog($LoggingString, $this.FullPath)
        $this.LogUpKeep()
    }

    $loggerObject | Add-Member -MemberType ScriptMethod -Name "WriteVerbose" -Value {
        param(
            [object]$LoggingString
        )

        if ($this.LoggerDisabled) {
            return
        }

        if ($null -eq $LoggingString) {
            throw [System.Management.Automation.ParameterBindingException] "Failed to provide valid LoggingString"
        }

        if ($this.EnableDateTime) {
            $LoggingString = "[{0}] : {1}" -f [System.DateTime]::Now, $LoggingString
        }
        $this.WriteVerboseWriter($LoggingString)
        $this.ToLog($LoggingString, $this.FullPath)
        $this.LogUpKeep()
    }

    $loggerObject | Add-Member -MemberType ScriptMethod -Name "WriteToFileOnly" -Value {
        param(
            [object]$LoggingString
        )

        if ($this.LoggerDisabled) {
            return
        }

        if ($null -eq $LoggingString) {
            throw [System.Management.Automation.ParameterBindingException] "Failed to provide valid LoggingString"
        }

        if ($this.EnableDateTime) {
            $LoggingString = "[{0}] : {1}" -f [System.DateTime]::Now, $LoggingString
        }
        $this.ToLog($LoggingString, $this.FullPath)
        $this.LogUpKeep()
    }

    $loggerObject | Add-Member -MemberType ScriptMethod -Name "UpdateFileLocation" -Value {

        if ($this.LoggerDisabled) {
            return
        }

        if ($null -eq $this.FullPath) {

            if ($this.IncludeDateTimeToFileName) {
                $this.InstanceBaseName = "{0}_{1}" -f $this.FileName, ((Get-Date).ToString('yyyyMMddHHmmss'))
                $this.FullPath = "{0}\{1}.txt" -f $this.FileDirectory, $this.InstanceBaseName
            } else {
                $this.InstanceBaseName = "{0}" -f $this.FileName
                $this.FullPath = "{0}\{1}.txt" -f $this.FileDirectory, $this.InstanceBaseName
            }
        } else {

            do {
                $this.FullPath = "{0}\{1}_{2}.txt" -f $this.FileDirectory, $this.InstanceBaseName, $this.InstanceNumber
                $this.InstanceNumber++
            }while (Test-Path $this.FullPath)
            $this.WriteVerbose("Updated to New Log")
        }
    }

    $loggerObject | Add-Member -MemberType ScriptMethod -Name "LogUpKeep" -Value {

        if ($this.LoggerDisabled) {
            return
        }

        if ($this.NextFileCheckTime -gt [System.DateTime]::Now) {
            return
        }
        $this.NextFileCheckTime = (Get-Date).AddMinutes($this.CheckSizeIntervalMinutes)
        $this.CheckFileSize()
        $this.CheckNumberOfFiles()
        $this.WriteVerbose("Did Log Object Up Keep")
    }

    $loggerObject | Add-Member -MemberType ScriptMethod -Name "CheckFileSize" -Value {

        if ($this.LoggerDisabled) {
            return
        }

        $item = Get-ChildItem $this.FullPath

        if (($item.Length / 1MB) -gt $this.MaxFileSizeInMB) {
            $this.UpdateFileLocation()
        }
    }

    $loggerObject | Add-Member -MemberType ScriptMethod -Name "CheckNumberOfFiles" -Value {

        if ($this.LoggerDisabled) {
            return
        }

        $filter = "{0}*" -f $this.InstanceBaseName
        $items = Get-ChildItem -Path $this.FileDirectory | Where-Object { $_.Name -like $filter }

        if ($items.Count -gt $this.NumberOfLogsToKeep) {
            do {
                $items | Sort-Object LastWriteTime | Select-Object -First 1 | Remove-Item -Force
                $items = Get-ChildItem -Path $this.FileDirectory | Where-Object { $_.Name -like $filter }
            }while ($items.Count -gt $this.NumberOfLogsToKeep)
        }
    }

    $loggerObject | Add-Member -MemberType ScriptMethod -Name "RemoveLatestLogFile" -Value {

        if ($this.LoggerDisabled) {
            return
        }

        if (!$this.PreventLogCleanup) {
            $item = Get-ChildItem $this.FullPath
            Remove-Item $item -Force
        }
    }

    $loggerObject | Add-Member -MemberType ScriptMethod -Name "DisableLogger" -Value {
        $this.LoggerDisabled = $true
    }

    $loggerObject | Add-Member -MemberType ScriptMethod -Name "EnableLogger" -Value {
        $this.LoggerDisabled = $false
    }

    $loggerObject.UpdateFileLocation()
    try {
        "[{0}] : Creating a new logger instance" -f [System.DAteTime]::Now | Out-File ($loggerObject.FullPath) -Append
    } catch {
        throw
    }

    return $loggerObject
}

<#
    Determines if the script has an update available. Use the optional
    -AutoUpdate switch to make it update itself. Returns $true if an
    update was downloaded, $false otherwise. The result will always
    be $false if the -AutoUpdate switch is not used.
#>
function Test-ScriptVersion {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter()]
        [switch]
        $AutoUpdate
    )

    function Confirm-ProxyServer {
        [CmdletBinding()]
        [OutputType([bool])]
        param (
            [Parameter(Mandatory = $true)]
            [string]
            $TargetUri
        )

        try {
            $proxyObject = ([System.Net.WebRequest]::GetSystemWebproxy()).GetProxy($TargetUri)
            if ($TargetUri -ne $proxyObject.OriginalString) {
                return $true
            } else {
                return $false
            }
        } catch {
            return $false
        }
    }

    function Confirm-Signature {
        [CmdletBinding()]
        [OutputType([bool])]
        param (
            [Parameter(Mandatory = $true)]
            [string]
            $File
        )

        $IsValid = $false
        $MicrosoftSigningRoot2010 = 'CN=Microsoft Root Certificate Authority 2010, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'
        $MicrosoftSigningRoot2011 = 'CN=Microsoft Root Certificate Authority 2011, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'

        try {
            $sig = Get-AuthenticodeSignature -FilePath $File

            if ($sig.Status -ne 'Valid') {
                Write-Warning "Signature is not trusted by machine as Valid, status: $($sig.Status)."
                throw
            }

            $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
            $chain.ChainPolicy.VerificationFlags = "IgnoreNotTimeValid"

            if (-not $chain.Build($sig.SignerCertificate)) {
                Write-Warning "Signer certificate doesn't chain correctly."
                throw
            }

            if ($chain.ChainElements.Count -le 1) {
                Write-Warning "Certificate Chain shorter than expected."
                throw
            }

            $rootCert = $chain.ChainElements[$chain.ChainElements.Count - 1]

            if ($rootCert.Certificate.Subject -ne $rootCert.Certificate.Issuer) {
                Write-Warning "Top-level certifcate in chain is not a root certificate."
                throw
            }

            if ($rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2010 -and $rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2011) {
                Write-Warning "Unexpected root cert. Expected $MicrosoftSigningRoot2010 or $MicrosoftSigningRoot2011, but found $($rootCert.Certificate.Subject)."
                throw
            }

            Write-Host "File signed by $($sig.SignerCertificate.Subject)"

            $IsValid = $true
        } catch {
            $IsValid = $false
        }

        $IsValid
    }

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    if ((Get-AuthenticodeSignature -FilePath $scriptFullName).Status -eq "NotSigned") {
        Write-Warning "This script appears to be an unsigned test build. Skipping version check."
        return $false
    }

    $oldName = [IO.Path]::GetFileNameWithoutExtension($scriptName) + ".old"
    $oldFullName = (Join-Path $scriptPath $oldName)

    $tempFullName = (Join-Path $env:TEMP $scriptName)

    $BuildVersion = "21.09.28.1529"
    try {
        $versionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        if (Confirm-ProxyServer -TargetUri "https://github.com") {
            $webClient = New-Object System.Net.WebClient
            $webClient.Headers.Add("User-Agent", "PowerShell")
            $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
        }
        $versionData = [Text.Encoding]::UTF8.GetString((Invoke-WebRequest $versionsUrl -UseBasicParsing).Content) | ConvertFrom-Csv
        $latestVersion = ($versionData | Where-Object { $_.File -eq $scriptName }).Version
        if ($null -ne $latestVersion -and $latestVersion -ne $BuildVersion) {
            if ($AutoUpdate) {
                if (Test-Path $tempFullName) {
                    Remove-Item $tempFullName -Force -Confirm:$false -ErrorAction Stop
                }
                Write-Host "AutoUpdate: Downloading update."
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Invoke-WebRequest "https://github.com/microsoft/CSS-Exchange/releases/latest/download/$scriptName" -OutFile $tempFullName -UseBasicParsing
                if (Confirm-Signature -File $tempFullName) {
                    Write-Host "AutoUpdate: Signature validated."
                    if (Test-Path $oldFullName) {
                        Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                    }
                    Move-Item $scriptFullName $oldFullName
                    Move-Item $tempFullName $scriptFullName
                    Write-Host "AutoUpdate: Succeeded."
                    return $true
                } else {
                    Write-Warning "AutoUpdate: Signature could not be verified: $tempFullName."
                    Write-Warning "AutoUpdate: Update was not applied."
                }
            } else {
                Write-Warning "$scriptName $BuildVersion is outdated. Please download the latest, version $latestVersion."
            }
        }
    } catch {
        # Work around empty catch block rule. The failure is intentionally silent.
        # For example, the script might be running on a computer with no internet access.
        "Version check failed" | Out-Null
    }

    return $false
}

Function Write-Host {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Proper handling of write host with colors')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1)]
        [object]$Object,
        [switch]$NoNewLine,
        [string]$ForegroundColor
    )
    begin {
        $consoleHost = $host.Name -eq "ConsoleHost"
        $params = @{
            Object    = $Object
            NoNewLine = $NoNewLine
        }
    }
    process {

        if ([string]::IsNullOrEmpty($ForegroundColor)) {
            if ($null -ne $host.UI.RawUI.ForegroundColor -and
                $consoleHost) {
                $params.Add("ForegroundColor", $host.UI.RawUI.ForegroundColor)
            }
        } elseif ($ForegroundColor -eq "Yellow" -and
            $consoleHost -and
            $null -ne $host.PrivateData.WarningForegroundColor) {
            $params.Add("ForegroundColor", $host.PrivateData.WarningForegroundColor)
        } elseif ($ForegroundColor -eq "Red" -and
            $consoleHost -and
            $null -ne $host.PrivateData.ErrorForegroundColor) {
            $params.Add("ForegroundColor", $host.PrivateData.ErrorForegroundColor)
        } else {
            $params.Add("ForegroundColor", $ForegroundColor)
        }

        Microsoft.PowerShell.Utility\Write-Host @params
    }
}

Function SetProperForegroundColor {
    $Script:OriginalConsoleForegroundColor = $host.UI.RawUI.ForegroundColor

    if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.WarningForegroundColor) {
        Write-Verbose "Foreground Color matches warning's color"

        if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
            $Host.UI.RawUI.ForegroundColor = "Gray"
        }
    }

    if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.ErrorForegroundColor) {
        Write-Verbose "Foreground Color matches error's color"

        if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
            $Host.UI.RawUI.ForegroundColor = "Gray"
        }
    }
}

Function RevertProperForegroundColor {
    $Host.UI.RawUI.ForegroundColor = $Script:OriginalConsoleForegroundColor
}

Function Main {

    if (-not (Confirm-Administrator) -and
        (-not $AnalyzeDataOnly -and
        -not $BuildHtmlServersReport -and
        -not $ScriptUpdateOnly)) {
        Write-Warning "The script needs to be executed in elevated mode. Start the Exchange Management Shell as an Administrator."
        $Error.Clear()
        Start-Sleep -Seconds 2;
        exit
    }

    $Error.Clear() #Always clear out the errors
    $Script:ErrorsExcludedCount = 0 #this is a way to determine if the only errors occurred were in try catch blocks. If there is a combination of errors in and out, then i will just dump it all out to avoid complex issues.
    $Script:ErrorsExcluded = @()
    $Script:date = (Get-Date)
    $Script:dateTimeStringFormat = $date.ToString("yyyyMMddHHmmss")

    if ($BuildHtmlServersReport) {
        Invoke-ScriptLogFileLocation -FileName "HealthChecker-HTMLServerReport"
        $files = Get-HealthCheckFilesItemsFromLocation
        $fullPaths = Get-OnlyRecentUniqueServersXMLs $files
        $importData = Import-MyData -FilePaths $fullPaths
        Get-HtmlServerReport -AnalyzedHtmlServerValues $importData.HtmlServerValues
        Start-Sleep 2;
        return
    }

    if ((Test-Path $OutputFilePath) -eq $false) {
        Write-Host "Invalid value specified for -OutputFilePath." -ForegroundColor Red
        return
    }

    if ($LoadBalancingReport) {
        Invoke-ScriptLogFileLocation -FileName "HealthChecker-LoadBalancingReport"
        Write-Green("Client Access Load Balancing Report on " + $date)
        Get-CASLoadBalancingReport
        Write-Grey("Output file written to " + $OutputFullPath)
        Write-Break
        Write-Break
        return
    }

    if ($DCCoreRatio) {
        $oldErrorAction = $ErrorActionPreference
        $ErrorActionPreference = "Stop"
        try {
            Get-ExchangeDCCoreRatio
            return
        } finally {
            $ErrorActionPreference = $oldErrorAction
        }
    }

    if ($MailboxReport) {
        Invoke-ScriptLogFileLocation -FileName "HealthChecker-MailboxReport" -IncludeServerName $true
        Get-MailboxDatabaseAndMailboxStatistics
        Write-Grey("Output file written to {0}" -f $Script:OutputFullPath)
        return
    }

    if ($AnalyzeDataOnly) {
        Invoke-ScriptLogFileLocation -FileName "HealthChecker-Analyzer"
        $files = Get-HealthCheckFilesItemsFromLocation
        $fullPaths = Get-OnlyRecentUniqueServersXMLs $files
        $importData = Import-MyData -FilePaths $fullPaths

        $analyzedResults = @()
        foreach ($serverData in $importData) {
            $analyzedServerResults = Invoke-AnalyzerEngine -HealthServerObject $serverData.HealthCheckerExchangeServer
            Write-ResultsToScreen -ResultsToWrite $analyzedServerResults.DisplayResults
            $analyzedResults += $analyzedServerResults
        }

        Get-HtmlServerReport -AnalyzedHtmlServerValues $analyzedResults.HtmlServerValues
        return
    }

    if ($ScriptUpdateOnly) {
        Invoke-ScriptLogFileLocation -FileName "HealthChecker-ScriptUpdateOnly"
        switch (Test-ScriptVersion -AutoUpdate) {
            ($true) { Write-Green("Script was successfully updated.") }
            ($false) { Write-Yellow("No update of the script performed.") }
            default { Write-Red("Unable to perform ScriptUpdateOnly operation.") }
        }
        return
    }

    Invoke-ScriptLogFileLocation -FileName "HealthChecker" -IncludeServerName $true
    $currentErrors = $Error.Count

    if ((-not $SkipVersionCheck) -and
        (Test-ScriptVersion -AutoUpdate)) {
        Write-Yellow "Script was updated. Please rerun the command."
        return
    } else {
        $Script:DisplayedScriptVersionAlready = $true
        Write-Green "Exchange Health Checker version $BuildVersion"
    }

    if ($currentErrors -ne $Error.Count) {
        $index = 0
        while ($index -lt ($Error.Count - $currentErrors)) {
            Invoke-CatchActions $Error[$index]
            $index++
        }
    }

    Test-RequiresServerFqdn
    [HealthChecker.HealthCheckerExchangeServer]$HealthObject = Get-HealthCheckerExchangeServer
    $analyzedResults = Invoke-AnalyzerEngine -HealthServerObject $HealthObject
    Write-ResultsToScreen -ResultsToWrite $analyzedResults.DisplayResults
    $currentErrors = $Error.Count

    try {
        $analyzedResults | Export-Clixml -Path $OutXmlFullPath -Encoding UTF8 -Depth 6 -ErrorAction SilentlyContinue
    } catch {
        Write-Verbose "Failed to Export-Clixml. Converting HealthCheckerExchangeServer to json"
        $jsonHealthChecker = $analyzedResults.HealthCheckerExchangeServer | ConvertTo-Json

        $testOuputxml = [PSCustomObject]@{
            HealthCheckerExchangeServer = $jsonHealthChecker | ConvertFrom-Json
            HtmlServerValues            = $analyzedResults.HtmlServerValues
            DisplayResults              = $analyzedResults.DisplayResults
        }

        $testOuputxml | Export-Clixml -Path $OutXmlFullPath -Encoding UTF8 -Depth 6 -ErrorAction Stop
    } finally {
        if ($currentErrors -ne $Error.Count) {
            $index = 0
            while ($index -lt ($Error.Count - $currentErrors)) {
                Invoke-CatchActions $Error[$index]
                $index++
            }
        }

        Write-Grey("Output file written to {0}" -f $Script:OutputFullPath)
        Write-Grey("Exported Data Object Written to {0} " -f $Script:OutXmlFullPath)
    }
}

try {
    $Script:Logger = New-LoggerObject -LogName "HealthChecker-$($Script:Server)-Debug" `
        -LogDirectory $OutputFilePath `
        -VerboseEnabled $Script:VerboseEnabled `
        -EnableDateTime $false `
        -ErrorAction SilentlyContinue
    SetProperForegroundColor
    Main
} finally {
    Get-ErrorsThatOccurred
    if ($Script:VerboseEnabled) {
        $Host.PrivateData.VerboseForegroundColor = $VerboseForeground
    }
    $Script:Logger.RemoveLatestLogFile()
    if ($Script:Logger.PreventLogCleanup) {
        Write-Host("Output Debug file written to {0}" -f $Script:Logger.FullPath)
    }
    RevertProperForegroundColor
}


# SIG # Begin signature block
# MIIjngYJKoZIhvcNAQcCoIIjjzCCI4sCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCqRz5cqHF1t8OP
# LUKpvSAVc+jd9EjQGbXnAoWQGhsHNKCCDYEwggX/MIID56ADAgECAhMzAAAB32vw
# LpKnSrTQAAAAAAHfMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjAxMjE1MjEzMTQ1WhcNMjExMjAyMjEzMTQ1WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC2uxlZEACjqfHkuFyoCwfL25ofI9DZWKt4wEj3JBQ48GPt1UsDv834CcoUUPMn
# s/6CtPoaQ4Thy/kbOOg/zJAnrJeiMQqRe2Lsdb/NSI2gXXX9lad1/yPUDOXo4GNw
# PjXq1JZi+HZV91bUr6ZjzePj1g+bepsqd/HC1XScj0fT3aAxLRykJSzExEBmU9eS
# yuOwUuq+CriudQtWGMdJU650v/KmzfM46Y6lo/MCnnpvz3zEL7PMdUdwqj/nYhGG
# 3UVILxX7tAdMbz7LN+6WOIpT1A41rwaoOVnv+8Ua94HwhjZmu1S73yeV7RZZNxoh
# EegJi9YYssXa7UZUUkCCA+KnAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUOPbML8IdkNGtCfMmVPtvI6VZ8+Mw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDYzMDA5MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAnnqH
# tDyYUFaVAkvAK0eqq6nhoL95SZQu3RnpZ7tdQ89QR3++7A+4hrr7V4xxmkB5BObS
# 0YK+MALE02atjwWgPdpYQ68WdLGroJZHkbZdgERG+7tETFl3aKF4KpoSaGOskZXp
# TPnCaMo2PXoAMVMGpsQEQswimZq3IQ3nRQfBlJ0PoMMcN/+Pks8ZTL1BoPYsJpok
# t6cql59q6CypZYIwgyJ892HpttybHKg1ZtQLUlSXccRMlugPgEcNZJagPEgPYni4
# b11snjRAgf0dyQ0zI9aLXqTxWUU5pCIFiPT0b2wsxzRqCtyGqpkGM8P9GazO8eao
# mVItCYBcJSByBx/pS0cSYwBBHAZxJODUqxSXoSGDvmTfqUJXntnWkL4okok1FiCD
# Z4jpyXOQunb6egIXvkgQ7jb2uO26Ow0m8RwleDvhOMrnHsupiOPbozKroSa6paFt
# VSh89abUSooR8QdZciemmoFhcWkEwFg4spzvYNP4nIs193261WyTaRMZoceGun7G
# CT2Rl653uUj+F+g94c63AhzSq4khdL4HlFIP2ePv29smfUnHtGq6yYFDLnT0q/Y+
# Di3jwloF8EWkkHRtSuXlFUbTmwr/lDDgbpZiKhLS7CBTDj32I0L5i532+uHczw82
# oZDmYmYmIUSMbZOgS65h797rj5JJ6OkeEUJoAVwwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVczCCFW8CAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAd9r8C6Sp0q00AAAAAAB3zAN
# BglghkgBZQMEAgEFAKCBxjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQg7DKwTzq0
# j2JcMQoKeDBGy7Pv03JfIE65mfmZxBhALwYwWgYKKwYBBAGCNwIBDDFMMEqgGoAY
# AEMAUwBTACAARQB4AGMAaABhAG4AZwBloSyAKmh0dHBzOi8vZ2l0aHViLmNvbS9t
# aWNyb3NvZnQvQ1NTLUV4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQCD6GalcAP9
# fc0gjz5S+W0YHGvRYuYcSntvYiV73YtQjlQWnA6Enczoy6RjVh7J2GI0MUu07ZhU
# dizpLhijvqELHK5fCQfq206LWGzeGGyVzWeoqe3K2A9p7refmLbR9zTo+Gv6XYUy
# +a8ir2jz9VkeC3GV/1qGQWijsMNZdVLabruJDnrXZSxzNDvZfVDz1Fkb4GEOPKSc
# 95apeuX3r4MSzo+HfT/KKO68j5zxlUVA5hox+mx0tjlCYTw/Ho29z9AFohCsKXF8
# VjecBqYTPyY+I6mC6Eq6G5w/aaXfEnJydUutoSVx7HQhkrLLeNC3z9N7//ZliFIZ
# dlTlYzL0L4VJoYIS5TCCEuEGCisGAQQBgjcDAwExghLRMIISzQYJKoZIhvcNAQcC
# oIISvjCCEroCAQMxDzANBglghkgBZQMEAgEFADCCAVEGCyqGSIb3DQEJEAEEoIIB
# QASCATwwggE4AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIIuCgMAW
# v5fLbS1ou3ySIailtN9Da2gPdahuIYBk6xYTAgZhQ6lGjHgYEzIwMjEwOTI5MTg1
# NzQxLjI0NFowBIACAfSggdCkgc0wgcoxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjdCRjEtRTNFQS1CODA4MSUwIwYD
# VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloIIOPDCCBPEwggPZoAMC
# AQICEzMAAAFRw1DnWWyqxqcAAAAAAVEwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UE
# BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
# BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0
# IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMjAxMTEyMTgyNjA0WhcNMjIwMjExMTgy
# NjA0WjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMG
# A1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEmMCQGA1UECxMdVGhh
# bGVzIFRTUyBFU046N0JGMS1FM0VBLUI4MDgxJTAjBgNVBAMTHE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQCf0ofvqoSuO+84iSNZsem0yRgOOYb4kSbOC7Kv9XGNmBn+KDwyTjuOpIk/lHEf
# +wPKqFi7uM9I7zqyJmHy7sMFf0vwj4AH7x88+8Pi6gsoPbYGmgWXgHwXDkrtK6Ju
# 9vEY3tp0vX/Nb6xZeVW+kOEQ8goMgK8R02MZMuGS19+2N5+D2W6YExQEnYbj+Dhp
# 3R0O9E2YqIxldd78uXhCD+g9LNcJQRihJKprkP7kxGKZV7n9hMuPSNWvyIXjlXSF
# PtUfw4k7hgiZydmGroPDUb7DoAJEZ48WY5apby0RnXdIyY6q4mtOTDLLzPI21W20
# kBft2IUttHRK8yVsllYrQod3AgMBAAGjggEbMIIBFzAdBgNVHQ4EFgQUxXf/42hQ
# YpM0aDo4zITp83VE6m0wHwYDVR0jBBgwFoAU1WM6XIoxkPNDe3xGG8UzaFqFbVUw
# VgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9j
# cmwvcHJvZHVjdHMvTWljVGltU3RhUENBXzIwMTAtMDctMDEuY3JsMFoGCCsGAQUF
# BwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aS9jZXJ0cy9NaWNUaW1TdGFQQ0FfMjAxMC0wNy0wMS5jcnQwDAYDVR0TAQH/BAIw
# ADATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQsFAAOCAQEAK/31wBWD
# mfHRKqO8t9DOa6AyPlwn00TrR25IfUunEdiKb0uzdR+Jh3u3Qm/ITD+tFMQodvOd
# XosUuVf76UckwYrNmce1N7Y4jpkcWc2IWG2DJa5gMmubspDKQ2LUbUtu5WJ70x6G
# agr6EGJmeetx9lKcFKiSu87ZARYcLXGdnnAzzZQSOmsVg6RyFT7pFygKOOYgUZ+B
# LM2PUwht/iVwnkWhXUyDoXAXjkKKM5cdVevOSKwxn2m4OkWOMRXpMBjog2AySEt6
# /8BWjDSwXwx9DO0kiUVh0USRnk0X8jLOgLZhv2LDhsIp0Gt0PcCzqa+gZI2MILqU
# 53PoR6skrc2EWDCCBnEwggRZoAMCAQICCmEJgSoAAAAAAAIwDQYJKoZIhvcNAQEL
# BQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAwBgNV
# BAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEwMB4X
# DTEwMDcwMTIxMzY1NVoXDTI1MDcwMTIxNDY1NVowfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCpHQ28
# dxGKOiDs/BOX9fp/aZRrdFQQ1aUKAIKF++18aEssX8XD5WHCdrc+Zitb8BVTJwQx
# H0EbGpUdzgkTjnxhMFmxMEQP8WCIhFRDDNdNuDgIs0Ldk6zWczBXJoKjRQ3Q6vVH
# gc2/JGAyWGBG8lhHhjKEHnRhZ5FfgVSxz5NMksHEpl3RYRNuKMYa+YaAu99h/EbB
# Jx0kZxJyGiGKr0tkiVBisV39dx898Fd1rL2KQk1AUdEPnAY+Z3/1ZsADlkR+79BL
# /W7lmsqxqPJ6Kgox8NpOBpG2iAg16HgcsOmZzTznL0S6p/TcZL2kAcEgCZN4zfy8
# wMlEXV4WnAEFTyJNAgMBAAGjggHmMIIB4jAQBgkrBgEEAYI3FQEEAwIBADAdBgNV
# HQ4EFgQU1WM6XIoxkPNDe3xGG8UzaFqFbVUwGQYJKwYBBAGCNxQCBAweCgBTAHUA
# YgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU
# 1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2Ny
# bC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIw
# MTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0w
# Ni0yMy5jcnQwgaAGA1UdIAEB/wSBlTCBkjCBjwYJKwYBBAGCNy4DMIGBMD0GCCsG
# AQUFBwIBFjFodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vUEtJL2RvY3MvQ1BTL2Rl
# ZmF1bHQuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAFAAbwBsAGkA
# YwB5AF8AUwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQAH
# 5ohRDeLG4Jg/gXEDPZ2joSFvs+umzPUxvs8F4qn++ldtGTCzwsVmyWrf9efweL3H
# qJ4l4/m87WtUVwgrUYJEEvu5U4zM9GASinbMQEBBm9xcF/9c+V4XNZgkVkt070IQ
# yK+/f8Z/8jd9Wj8c8pl5SpFSAK84Dxf1L3mBZdmptWvkx872ynoAb0swRCQiPM/t
# A6WWj1kpvLb9BOFwnzJKJ/1Vry/+tuWOM7tiX5rbV0Dp8c6ZZpCM/2pif93FSguR
# JuI57BlKcWOdeyFtw5yjojz6f32WapB4pm3S4Zz5Hfw42JT0xqUKloakvZ4argRC
# g7i1gJsiOCC1JeVk7Pf0v35jWSUPei45V3aicaoGig+JFrphpxHLmtgOR5qAxdDN
# p9DvfYPw4TtxCd9ddJgiCGHasFAeb73x4QDf5zEHpJM692VHeOj4qEir995yfmFr
# b3epgcunCaw5u+zGy9iCtHLNHfS4hQEegPsbiSpUObJb2sgNVZl6h3M7COaYLeqN
# 4DMuEin1wC9UJyH3yKxO2ii4sanblrKnQqLJzxlBTeCG+SqaoxFmMNO7dDJL32N7
# 9ZmKLxvHIa9Zta7cRDyXUHHXodLFVeNp3lfB0d4wwP3M5k37Db9dT+mdHhk4L7zP
# WAUu7w2gUDXa7wknHNWzfjUeCLraNtvTX4/edIhJEqGCAs4wggI3AgEBMIH4oYHQ
# pIHNMIHKMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYD
# VQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMSYwJAYDVQQLEx1UaGFs
# ZXMgVFNTIEVTTjo3QkYxLUUzRUEtQjgwODElMCMGA1UEAxMcTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIaAxUAoKKvc/E/pEILJUwlIBWg
# xXrXI16ggYMwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkq
# hkiG9w0BAQUFAAIFAOT+ojUwIhgPMjAyMTA5MjkxNjI4MzdaGA8yMDIxMDkzMDE2
# MjgzN1owdzA9BgorBgEEAYRZCgQBMS8wLTAKAgUA5P6iNQIBADAKAgEAAgIHDgIB
# /zAHAgEAAgIRNTAKAgUA5P/ztQIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEE
# AYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBBQUAA4GB
# AERhGILydp9XfRHxp6IpBaaj+ofluAMLLVZokSjMVosTGNNt/8Eu1Smw5qcHrcaj
# xR1yommaOHblILPCItaBmqnj2LpDL0ahiS0W/vv14EN92vVt2ncR3zMrslaIL1Uk
# 3TLGzbXYxtVrmIjLlBo5jfFzAtcGJgETFEme8sKFcgIbMYIDDTCCAwkCAQEwgZMw
# fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMd
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAFRw1DnWWyqxqcAAAAA
# AVEwDQYJYIZIAWUDBAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRAB
# BDAvBgkqhkiG9w0BCQQxIgQgNLQhY7vEPxjgfHKtWOiwwGi8Ux7uzwx1kZweiIHX
# 4CQwgfoGCyqGSIb3DQEJEAIvMYHqMIHnMIHkMIG9BCAuzVyZiPjWwVkHAKYW+/1J
# w/m265SHGy/+3QH1cXrlQTCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
# QSAyMDEwAhMzAAABUcNQ51lsqsanAAAAAAFRMCIEIEZOrADZqGfudv+HfY68tq9O
# +GN3FJJ/N6x7KJKzwkH+MA0GCSqGSIb3DQEBCwUABIIBAHJ7dJLZOkmLiKA7HZ5B
# Rf61f60tffoV7giWQgvpRmbrB9bJ/P4KTiObNXHWwte/a2r9J0LxwylIGPAAfo/v
# K8F7o8qsMd9uDyCBDW6rPSgb3k35/hX4txhd49gTsOaZIVgLGU1U5aFlxwfHyNsS
# AA8escgK0xSQ+2JtzjnELys/JNHvCl5eaEU+k+J+QSQT49+lBM0bf0d96Dzucjgw
# RlUrJwJaoRQ2KHVSvUPl/D6LbXedYeFonjhywkuCHPhbUa7z17vNCsCgwqGp6F7W
# 6U0qDokhyTQG92v+qzVdm0qYpX03zPy2uuqCbVUgZWpyUIk6g0/nFfVsriwUM3vg
# +4M=
# SIG # End signature block
