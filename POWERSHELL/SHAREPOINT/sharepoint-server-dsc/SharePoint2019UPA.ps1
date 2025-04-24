param (
    [Parameter(Mandatory = $true)] [ValidateNotNullorEmpty()] [PSCredential] $SetupAccount,
    [Parameter(Mandatory = $true)] [ValidateNotNullorEmpty()] [PSCredential] $SyncAccount,
    [Parameter(Mandatory = $false)] [ValidateNotNullorEmpty()] [string] $Computer = "localhost",
    [Parameter(Mandatory = $false)] [ValidateNotNullorEmpty()] [string] $SqlInstance = "DEV-SQL\SHAREPOINT"
)

[string]$Scriptpath = $MyInvocation.MyCommand.Path
[string]$Dir = Split-Path $Scriptpath
Set-Location $Dir

Configuration SharePoint2019UPA
{
    Import-DscResource -ModuleName PSDesiredStateConfiguration
    Import-DscResource -ModuleName SharePointDsc

    node $Computer
    {
        $serviceAppPoolName = "SharePoint Service Applications"

        SPUserProfileServiceApp UserProfileServiceApp {
            Name                 = "User Profile Service Application"
            ProxyName            = "User Profile Service Application Proxy"
            ApplicationPool      = $serviceAppPoolName
            MySiteHostLocation   = "http://dev-sharepoint.local:8080"
            MySiteManagedPath    = "personal"
            ProfileDBName        = "SP_Profile"
            ProfileDBServer      = $SqlInstance
            SocialDBName         = "SP_Social"
            SocialDBServer       = $SqlInstance
            SyncDBName           = "SP_Sync"
            SyncDBServer         = $SqlInstance
            EnableNetBIOS        = $false
            PsDscRunAsCredential = $SetupAccount
        }

        SPUserProfileSyncConnection UserProfileSyncConnection {
            UserProfileService    = "User Profile Service Application"
            Forest                = "ad.local"
            Name                  = "AD"
            ConnectionCredentials = $SyncAccount
            Server                = "ad.local"
            UseSSL                = $false
            IncludedOUs           = @("OU=DEVUsers,OU=UserAccounts,DC=ad,DC=local")
            Force                 = $false
            ConnectionType        = "ActiveDirectory"
            PsDscRunAsCredential  = $SetupAccount
            DependsOn             = "[SPUserProfileServiceApp]UserProfileServiceApp"
        }

        SPUserProfileServiceAppPermissions UPAPermissions {
            ProxyName            = "User Profile Service Application Proxy"
            CreatePersonalSite   = @("Everyone")
            FollowAndEditProfile = @("Everyone")
            UseTagsAndNotes      = @("None")
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPUserProfileServiceApp]UserProfileServiceApp"
        }

        LocalConfigurationManager {
            RebootNodeIfNeeded = $true
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

SharePoint2019UPA -ConfigurationData $config -OutputPath $Dir
Start-DscConfiguration -Path $Dir -Verbose -Wait -Force -ComputerName $Computer -ErrorAction Stop
