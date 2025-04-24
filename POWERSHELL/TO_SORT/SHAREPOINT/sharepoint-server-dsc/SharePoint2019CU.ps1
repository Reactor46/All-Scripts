param (
    [Parameter(Mandatory = $true)] [ValidateNotNullorEmpty()] [PSCredential] $SetupAccount,
    [Parameter(Mandatory = $false)] [ValidateNotNullorEmpty()] [string] $Computer = "localhost"
)

[string]$Scriptpath = $MyInvocation.MyCommand.Path
[string]$Dir = Split-Path $Scriptpath
Set-Location $Dir

Configuration SharePoint2019CU
{
    Import-DscResource -ModuleName PSDesiredStateConfiguration
    Import-DscResource -ModuleName SharePointDsc

    node $Computer
    {
        SPProductUpdate ProductUpdate {
            SetupFile            = "D:\Sources\SharePoint2019\Updates\wssloc2019-kb5002310-fullfile-x64-glb.exe"
            ShutdownServices     = $true
            PsDscRunAsCredential = $SetupAccount
        }

        SPConfigWizard PSConfig {
            IsSingleInstance     = "Yes"
            Ensure               = "Present"
            PsDscRunAscredential = $SetupAccount
            DependsOn            = "[SPProductUpdate]ProductUpdate"
        }
    }
}

$config = @{
    AllNodes = @(
        @{
            NodeName                    = $Computer
            PSDscAllowPlainTextPassword = $true
            PSDscAllowDomainUser        = $true
        }
    )
}

SharePoint2019CU -ConfigurationData $config -OutputPath $Dir
Start-DscConfiguration -Path $Dir -Verbose -Wait -Force -ComputerName $Computer -ErrorAction Stop
