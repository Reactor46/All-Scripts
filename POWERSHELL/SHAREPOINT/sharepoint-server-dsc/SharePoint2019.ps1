param (
    [Parameter(Mandatory = $true)] [ValidateNotNullorEmpty()] [PSCredential] $FarmAccount,
    [Parameter(Mandatory = $true)] [ValidateNotNullorEmpty()] [PSCredential] $SetupAccount,
    [Parameter(Mandatory = $true)] [ValidateNotNullorEmpty()] [PSCredential] $WebPoolManagedAccount,
    [Parameter(Mandatory = $true)] [ValidateNotNullorEmpty()] [PSCredential] $ServicePoolManagedAccount,
    [Parameter(Mandatory = $true)] [ValidateNotNullorEmpty()] [PSCredential] $SyncAccount,
    [Parameter(Mandatory = $false)] [ValidateNotNullorEmpty()] [PSCredential] $Passphrase,
    [Parameter(Mandatory = $false)] [ValidateNotNullorEmpty()] [string] $Computer = "localhost",
    [Parameter(Mandatory = $false)] [ValidateNotNullorEmpty()] [string] $SqlInstance = "DEV-SQL\SHAREPOINT",
    [Parameter(Mandatory = $false)] [ValidateNotNullorEmpty()] [string] $ProductKey = "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"
)

[string]$Scriptpath = $MyInvocation.MyCommand.Path
[string]$Dir = Split-Path $Scriptpath
Set-Location $Dir

Configuration SharePoint2019
{
    Import-DscResource -ModuleName PSDesiredStateConfiguration
    Import-DscResource -ModuleName SharePointDsc
    Import-DscResource -ModuleName xDownloadFile

    node $Computer
    {

        xDownloadFile AppFabricKBDL {
            SourcePath               = "https://download.microsoft.com/download/F/1/0/F1093AF6-E797-4CA8-A9F6-FC50024B385C/AppFabric-KB3092423-x64-ENU.exe"
            FileName                 = "AppFabric-KB3092423-x64-ENU.exe"
            DestinationDirectoryPath = "D:\Sources\SharePoint2019\Prerequisite"
            PsDscRunAsCredential     = $SetupAccount
        }

        xDownloadFile MicrosoftIdentityExtensionsDL {
            SourcePath               = "http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/rtm/MicrosoftIdentityExtensions-64.msi"
            FileName                 = "MicrosoftIdentityExtensions-64.msi"
            DestinationDirectoryPath = "D:\Sources\SharePoint2019\Prerequisite"
            DependsOn                = "[xDownloadFile]AppFabricKBDL"
            PsDscRunAsCredential     = $SetupAccount
        }

        xDownloadFile MSIPCDL {
            SourcePath               = "https://download.microsoft.com/download/3/C/F/3CF781F5-7D29-4035-9265-C34FF2369FA2/setup_msipc_x64.exe"
            FileName                 = "setup_msipc_x64.exe"
            DestinationDirectoryPath = "D:\Sources\SharePoint2019\Prerequisite"
            DependsOn                = "[xDownloadFile]MicrosoftIdentityExtensionsDL"
            PsDscRunAsCredential     = $SetupAccount
        }

        xDownloadFile SQLNCLIDL {
            SourcePath               = "https://download.microsoft.com/download/B/E/D/BED73AAC-3C8A-43F5-AF4F-EB4FEA6C8F3A/ENU/x64/sqlncli.msi"
            FileName                 = "sqlncli.msi"
            DestinationDirectoryPath = "D:\Sources\SharePoint2019\Prerequisite"
            DependsOn                = "[xDownloadFile]MSIPCDL"
            PsDscRunAsCredential     = $SetupAccount
        }

        xDownloadFile WcfDataServices56DL {
            SourcePath               = "http://download.microsoft.com/download/1/C/A/1CAA41C7-88B9-42D6-9E11-3C655656DAB1/WcfDataServices.exe"
            FileName                 = "WcfDataServices56.exe"
            DestinationDirectoryPath = "D:\Sources\SharePoint2019\Prerequisite"
            DependsOn                = "[xDownloadFile]SQLNCLIDL"
            PsDscRunAsCredential     = $SetupAccount
        }

        xDownloadFile AppFabricDL {
            SourcePath               = "http://download.microsoft.com/download/A/6/7/A678AB47-496B-4907-B3D4-0A2D280A13C0/WindowsServerAppFabricSetup_x64.exe"
            FileName                 = "WindowsServerAppFabricSetup_x64.exe"
            DestinationDirectoryPath = "D:\Sources\SharePoint2019\Prerequisite"
            DependsOn                = "[xDownloadFile]WcfDataServices56DL"
            PsDscRunAsCredential     = $SetupAccount
        }

        xDownloadFile DotNet472 {
            SourcePath               = "http://go.microsoft.com/fwlink/?linkid=863265"
            FileName                 = "NDP472-KB4054530-x86-x64-AllOS-ENU.exe"
            DestinationDirectoryPath = "D:\Sources\SharePoint2019\Prerequisite"
            DependsOn                = "[xDownloadFile]AppFabricDL"
            PsDscRunAsCredential     = $SetupAccount
        }

        xDownloadFile SynchronizationDL {
            SourcePath               = "http://download.microsoft.com/download/E/0/0/E0060D8F-2354-4871-9596-DC78538799CC/Synchronization.msi"
            FileName                 = "Synchronization.msi"
            DestinationDirectoryPath = "D:\Sources\SharePoint2019\Prerequisite"
            DependsOn                = "[xDownloadFile]DotNet472"
            PsDscRunAsCredential     = $SetupAccount
        }

        xDownloadFile MSVCRT141 {
            SourcePath               = "https://aka.ms/vs/15/release/vc_redist.x64.exe"
            FileName                 = "vc_redist.x64.exe"
            DestinationDirectoryPath = "D:\Sources\SharePoint2019\Prerequisite"
            DependsOn                = "[xDownloadFile]SynchronizationDL"
            PsDscRunAsCredential     = $SetupAccount
        }
        xDownloadFile MSVCRT11 {
            SourcePath               = "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe"
            FileName                 = "vcredist_x64.exe"
            DestinationDirectoryPath = "D:\Sources\SharePoint2019\Prerequisite"
            DependsOn                = "[xDownloadFile]MSVCRT141"
            PsDscRunAsCredential     = $SetupAccount
        }

        SPInstallPrereqs InstallPrereqs {
            IsSingleInstance     = "Yes"
            Ensure               = "Present"
            InstallerPath        = "D:\Sources\SharePoint2019\prerequisiteinstaller.exe"
            OnlineMode           = $false
            SQLNCli              = "D:\Sources\SharePoint2019\Prerequisite\sqlncli.msi"
            Sync                 = "D:\Sources\SharePoint2019\Prerequisite\Synchronization.msi"
            AppFabric            = "D:\Sources\SharePoint2019\Prerequisite\WindowsServerAppFabricSetup_x64.exe"
            IDFX11               = "D:\Sources\SharePoint2019\Prerequisite\MicrosoftIdentityExtensions-64.msi"
            MSIPCClient          = "D:\Sources\SharePoint2019\Prerequisite\setup_msipc_x64.exe"
            WCFDataServices56    = "D:\Sources\SharePoint2019\Prerequisite\WcfDataServices56.exe"
            MSVCRT11             = "D:\Sources\SharePoint2019\Prerequisite\vcredist_x64.exe"
            MSVCRT141            = "D:\Sources\SharePoint2019\Prerequisite\vc_redist.x64.exe"
            KB3092423            = "D:\Sources\SharePoint2019\Prerequisite\AppFabric-KB3092423-x64-ENU.exe"
            DotNet472            = "D:\Sources\SharePoint2019\Prerequisite\NDP472-KB4054530-x86-x64-AllOS-ENU.exe"
            DependsOn            = "[xDownloadFile]MSVCRT141"
            SXSPath              = "D:\Sources\sxs"
            PsDscRunAsCredential = $SetupAccount
        }
		
        SPInstall InstallSharePoint {
            IsSingleInstance = "Yes"
            Ensure           = "Present"
            BinaryDir        = "D:\Sources\SharePoint2019"
            ProductKey       = $ProductKey
            DependsOn        = "[SPInstallPrereqs]InstallPrereqs"
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
                D:\Sources\fr_sharepoint_server_2019_language_pack_x64_88abf125.exe /extract:D:\Sources\SharePoint2019\lang\fr
                Start-Sleep 5
            }
            TestScript = {
                Test-Path "D:\Sources\SharePoint2019\lang\fr"
            }
            DependsOn  = "[SPInstall]InstallSharePoint"
        }
		
        SPInstallLanguagePack InstallLanguagePackFR {
            BinaryDir            = "D:\Sources\SharePoint2019\lang\fr"
            Ensure               = "Present"
            DependsOn            = "[Script]ExtractBinariesFR"
            PsDscRunAsCredential = $SetupAccount
        }
		
        SPFarm CreateSPFarm {
            IsSingleInstance          = "Yes"
            Ensure                    = "Present"
            DatabaseServer            = $SqlInstance
            FarmConfigDatabaseName    = "SP_Config"
            Passphrase                = $Passphrase
            FarmAccount               = $FarmAccount
            PsDscRunAsCredential      = $SetupAccount
            AdminContentDatabaseName  = "SP_AdminContent"
            CentralAdministrationPort = 4321
            RunCentralAdmin           = $true
            ServerRole                = "SingleServerFarm"
            DeveloperDashboard        = "On"
            DependsOn                 = "[SPInstallLanguagePack]InstallLanguagePackFR"
        }

        SPManagedAccount ServicePoolManagedAccount {
            AccountName          = $ServicePoolManagedAccount.UserName
            Account              = $ServicePoolManagedAccount
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPFarm]CreateSPFarm"
        }

        SPManagedAccount WebPoolManagedAccount {
            AccountName          = $WebPoolManagedAccount.UserName
            Account              = $WebPoolManagedAccount
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPFarm]CreateSPFarm"
        }

        SPDiagnosticLoggingSettings ApplyDiagnosticLogSettings {
            IsSingleInstance                            = "Yes"
            PsDscRunAsCredential                        = $SetupAccount
            LogPath                                     = "L:\ULS"
            LogSpaceInGB                                = 8
            AppAnalyticsAutomaticUploadEnabled          = $false
            CustomerExperienceImprovementProgramEnabled = $true
            DaysToKeepLogs                              = 5
            DownloadErrorReportingUpdatesEnabled        = $false
            ErrorReportingAutomaticUploadEnabled        = $false
            ErrorReportingEnabled                       = $false
            EventLogFloodProtectionEnabled              = $true
            EventLogFloodProtectionNotifyInterval       = 5
            EventLogFloodProtectionQuietPeriod          = 2
            EventLogFloodProtectionThreshold            = 5
            EventLogFloodProtectionTriggerPeriod        = 2
            LogCutInterval                              = 15
            LogMaxDiskSpaceUsageEnabled                 = $true
            ScriptErrorReportingDelay                   = 30
            ScriptErrorReportingEnabled                 = $true
            ScriptErrorReportingRequireAuth             = $true
            DependsOn                                   = "[SPFarm]CreateSPFarm"
        }

        SPUsageApplication UsageApplication {
            Name                  = "Usage Service Application"
            DatabaseName          = "SP_Usage"
            UsageLogCutTime       = 5
            UsageLogLocation      = "L:\UsageLogs"
            UsageLogMaxFileSizeKB = 1024
            PsDscRunAsCredential  = $SetupAccount
            DependsOn             = "[SPFarm]CreateSPFarm"
        }

        SPStateServiceApp StateServiceApp {
            Name                 = "State Service Application"
            DatabaseName         = "SP_State"
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPFarm]CreateSPFarm"
        }

        SPDistributedCacheService EnableDistributedCache {
            Name                 = "AppFabricCachingService"
            Ensure               = "Present"
            CacheSizeInMB        = 1024
            ServiceAccount       = $ServicePoolManagedAccount.UserName
            PsDscRunAsCredential = $SetupAccount
            CreateFirewallRules  = $true
            DependsOn            = @('[SPFarm]CreateSPFarm', '[SPManagedAccount]ServicePoolManagedAccount')
        }

        SPWebApplication SharePointSites {
            Name                   = "Web App 2019"
            ApplicationPool        = "Web App 2019"
            ApplicationPoolAccount = $WebPoolManagedAccount.UserName
            AllowAnonymous         = $false
            DatabaseName           = "SP_Content"
            WebAppUrl              = "http://dev-sharepoint"
            HostHeader             = "dev-sharepoint"
            Port                   = 80
            PsDscRunAsCredential   = $SetupAccount
            DependsOn              = "[SPManagedAccount]WebPoolManagedAccount"
        }

        SPCacheAccounts WebAppCacheAccounts {
            WebAppUrl            = "http://dev-sharepoint"
            SuperUserAlias       = "AD\dev_shp_SU"
            SuperReaderAlias     = "AD\dev_shp_SR"
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPWebApplication]SharePointSites"
        }

        SPSite TeamSite {
            Url                  = "http://dev-sharepoint"
            OwnerAlias           = "AD\dev_shp_Pool"
            Name                 = "Team Site 2019"
            Template             = "STS#0"
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPWebApplication]SharePointSites"
        }

        SPWebApplication SharePointMySite {
            Name                   = "MySite"
            ApplicationPool        = "MySite"
            ApplicationPoolAccount = $WebPoolManagedAccount.UserName
            AllowAnonymous         = $false
            DatabaseName           = "SP_MySiteContent"
            WebAppUrl              = "http://dev-sharepoint"
            HostHeader             = "dev-sharepoint"
            Port                   = 8080
            PsDscRunAsCredential   = $SetupAccount
            DependsOn              = "[SPManagedAccount]WebPoolManagedAccount"
        }

        SPCacheAccounts WebAppCacheMySiteAccounts {
            WebAppUrl            = "http://dev-sharepoint:8080"
            SuperUserAlias       = "AD\dev_shp_SU"
            SuperReaderAlias     = "AD\dev_shp_SR"
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPWebApplication]SharePointMySite"
        }

        SPSite MySite {
            Url                  = "http://dev-sharepoint:8080"
            OwnerAlias           = "AD\dev_shp_Portal"
            Name                 = "MySite"
            Template             = "SPSMSITEHOST#0"
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPWebApplication]SharePointMySite"
        }

        SPServiceInstance ClaimsToWindowsTokenServiceInstance {
            Name                 = "Claims to Windows Token Service"
            Ensure               = "Present"
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPFarm]CreateSPFarm"
        }

        SPServiceInstance SecureStoreServiceInstance {
            Name                 = "Secure Store Service"
            Ensure               = "Present"
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPFarm]CreateSPFarm"
        }

        SPServiceInstance ManagedMetadataServiceInstance {
            Name                 = "Managed Metadata Web Service"
            Ensure               = "Present"
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPFarm]CreateSPFarm"
        }

        SPServiceInstance SearchServiceInstance {
            Name                 = "SharePoint Server Search"
            Ensure               = "Present"
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPFarm]CreateSPFarm"
        }

        $serviceAppPoolName = "SharePoint Service Applications"
        SPServiceAppPool MainServiceAppPool {
            Name                 = $serviceAppPoolName
            ServiceAccount       = $ServicePoolManagedAccount.UserName
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPFarm]CreateSPFarm"
        }

        SPSecureStoreServiceApp SecureStoreServiceApp {
            Name                 = "Secure Store Service Application"
            ApplicationPool      = $serviceAppPoolName
            AuditingEnabled      = $true
            AuditlogMaxSize      = 30
            DatabaseName         = "SP_SecureStore"
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPServiceAppPool]MainServiceAppPool"
        }

        SPManagedMetaDataServiceApp ManagedMetadataServiceApp {
            Name                 = "Managed Metadata Service Application"
            PsDscRunAsCredential = $SetupAccount
            ApplicationPool      = $serviceAppPoolName
            DatabaseName         = "SP_MMS"
            DependsOn            = "[SPServiceAppPool]MainServiceAppPool"
        }

        SPSearchServiceApp SearchServiceApp {
            Name                 = "Search Service Application"
            DatabaseName         = "SP_Search"
            ApplicationPool      = $serviceAppPoolName
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPServiceAppPool]MainServiceAppPool"
        }

        SPUserProfileServiceApp UserProfileServiceApp {
            Name                 = "User Profile Service Application"
            ProxyName            = "User Profile Service Application Proxy"
            ApplicationPool      = $serviceAppPoolName
            MySiteHostLocation   = "http://dev-sharepoint:8080"
            MySiteManagedPath    = "personal"
            ProfileDBName        = "SP_Profile"
            ProfileDBServer      = $SqlInstance
            SocialDBName         = "SP_Social"
            SocialDBServer       = $SqlInstance
            SyncDBName           = "SP_Sync"
            SyncDBServer         = $SqlInstance
            EnableNetBIOS        = $false
            PsDscRunAsCredential = $SetupAccount
            DependsOn            = "[SPServiceAppPool]MainServiceAppPool"
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

SharePoint2019 -ConfigurationData $config -OutputPath $Dir
Start-DscConfiguration -Path $Dir -Verbose -Wait -Force -ComputerName $Computer -ErrorAction Stop