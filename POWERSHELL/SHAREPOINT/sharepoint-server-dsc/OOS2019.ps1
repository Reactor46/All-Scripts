param (
    [Parameter(Mandatory = $true)] [ValidateNotNullorEmpty()] [string] $InternalUrl,
    [Parameter(Mandatory = $false)] [ValidateNotNullorEmpty()] [string] $Computer = "localhost"
)

[string]$Scriptpath = $MyInvocation.MyCommand.Path
[string]$Dir = Split-Path $Scriptpath
Set-Location $Dir

Configuration OOS2019
{
    Import-DscResource -ModuleName PSDesiredStateConfiguration
    Import-DscResource -ModuleName OfficeOnlineServerDSC

    node $Computer
    {
        WindowsFeature WebServer
        {
            Name = "Web-Server"
            Ensure = "Present"
            IncludeAllSubFeature = $true
        }

        OfficeOnlineServerInstall OOSInstall
        {
            Ensure = "Present"
            Path = "D:\Sources\OfficeOnlineServer\2019\setup.exe"
            DependsOn = "[WindowsFeature]WebServer"
        }

        Script ExtractBinariesFR {
            GetScript  = {
                @{
                    GetScript  = $GetScript
                    SetScript  = $SetScript
                    TestScript = $TestScript
                }
            }
            SetScript  = {
                D:\Sources\OfficeOnlineServer\fr_office_online_server_language_pack_last_updated_november_2018_x64_b65c95d5.exe /extract:D:\Sources\OfficeOnlineServer\2019\lang\fr
                Start-Sleep 5
            }
            TestScript = {
                Test-Path "D:\Sources\OfficeOnlineServer\2019\lang\fr"
            }
            DependsOn  = "[OfficeOnlineServerInstall]OOSInstall"
        }

        OfficeOnlineServerInstallLanguagePack InstallLanguagePackFR
        {
            Ensure = "Present"
            BinaryDir = "D:\Sources\OfficeOnlineServer\2019\lang\fr"
            Language = "fr-fr"
            DependsOn = "[Script]ExtractBinariesFR"
        }

        OfficeOnlineServerFarm FarmConfig
        {
            InternalUrl = $InternalUrl
            AllowHttp = $true
            EditingEnabled = $true
            AllowHttpSecureStoreConnections = $true
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

OOS2019 -ConfigurationData $config -OutputPath $Dir
Start-DscConfiguration -Path $Dir -Verbose -Wait -Force -ComputerName $Computer -ErrorAction Stop