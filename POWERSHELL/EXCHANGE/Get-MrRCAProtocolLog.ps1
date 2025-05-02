#Requires -Version 3.0
function Get-MrRCAProtocolLog {
<#
.SYNOPSIS
    Identifies and reports which Outlook client versions are being used to access Exchange.
.DESCRIPTION
    Get-MrRCAProtocolLog is an advanced PowerShell function that parses Exchange Server RPC
    logs to determine what Outlook client versions are being used to access the Exchange Server.
.PARAMETER LogFile
    The path to the Exchange RPC log files.
.EXAMPLE
     Get-MrRCAProtocolLog -LogFile 'C:\Program Files\Microsoft\Exchange Server\V15\Logging\RPC Client Access\RCA_20140831-1.LOG'
.EXAMPLE
     Get-ChildItem -Path '\\servername\c$\Program Files\Microsoft\Exchange Server\V15\Logging\RPC Client Access\*.log' |
     Get-MrRCAProtocolLog |
     Out-GridView -Title 'Outlook Client Versions'
.INPUTS
    String
.OUTPUTS
    PSCustomObject
.NOTES
    Author:  Mike F Robbins
    Website: http://mikefrobbins.com
    Twitter: @mikefrobbins
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
                   ValueFromPipeline)]
        [ValidateScript({
            Test-Path -Path $_ -PathType Leaf -Include '*.log'
        })]
        [string[]]$LogFile
    )
    PROCESS {
        foreach ($file in $LogFile) {
            $Headers = (Get-Content -Path $file -TotalCount 5 | Where-Object {$_ -like '#Fields*'}) -replace '#Fields: ' -split ','
            Import-Csv -Header $Headers -Path $file |
            Where-Object {$_.operation -eq 'Connect' -and $_.'client-software' -eq 'outlook.exe'} |
            Select-Object -Unique -Property @{label='User';expression={$_.'client-name' -replace '^.*cn='}},
                                            #@{label='DN';expression={$_.'client-name'}},
                                            client-software,
                                            @{label='Version';expression={Get-MrOutlookVersion -OutlookBuild $_.'client-software-version'}},
                                            client-mode,
                                            client-ip,
                                            protocol
        }
    }
}
function Get-MrOutlookVersion {
    param (
        [string]$OutlookBuild
    )
    switch ($OutlookBuild) {
        {$_ -ge '15.0.4569.1506'} {'Outlook 2013 SP1'; break}
        {$_ -ge '15.0.4420.1017'} {'Outlook 2013 RTM'; break}
        {$_ -ge '14.0.7015.1000'} {'Outlook 2010 SP2'; break}
        {$_ -ge '14.0.6029.1000'} {'Outlook 2010 SP1'; break}
        {$_ -ge '14.0.4763.1000'} {'Outlook 2010 RTM'; break}
        {$_ -ge '12.0.6607.1000'} {'Outlook 2007 SP3'; break}
        {$_ -ge '12.0.6423.1000'} {'Outlook 2007 SP2'; break}
        {$_ -ge '12.0.6212.1000'} {'Outlook 2007 SP1'; break}
        {$_ -ge '12.0.4518.1014'} {'Outlook 2007 RTM'; break}
        Default {'Unknown'}
    }
}