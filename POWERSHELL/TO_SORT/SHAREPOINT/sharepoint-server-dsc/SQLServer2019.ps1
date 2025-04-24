param (
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][System.Management.Automation.PSCredential]	$SqlInstallCredential,
    [Parameter()][ValidateNotNullOrEmpty()][System.Management.Automation.PSCredential] $SqlAdministratorCredential = $SqlInstallCredential,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][System.Management.Automation.PSCredential]	$SqlServiceCredential,
    [Parameter()][ValidateNotNullOrEmpty()][System.Management.Automation.PSCredential] $SqlAgentServiceCredential = $SqlServiceCredential,
    [Parameter(Mandatory = $false)] [ValidateNotNullorEmpty()] [string] $Computer = "localhost",
    [Parameter(Mandatory = $false)] [ValidateNotNullorEmpty()] [string] $Instance = "SHAREPOINT"
)

[string]$Scriptpath = $MyInvocation.MyCommand.Path
[string]$Dir = Split-Path $Scriptpath
Set-Location $Dir

Configuration SQLServer2019
{
    Import-DscResource -ModuleName PSDesiredStateConfiguration
    Import-DscResource -ModuleName SqlServerDsc

    node $Computer
    {
        WindowsFeature 'NetFramework' {
            Name   = "NET-Framework-45-Core"
            Ensure = "Present"
        }

        SqlSetup $Instance {
            InstanceName          = $Instance
            Features              = "SQLENGINE,AS,FULLTEXT"
            SQLCollation          = "Latin1_General_CI_AS_KS_WS"
            SQLSvcAccount         = $SqlServiceCredential
            AgtSvcAccount         = $SqlAgentServiceCredential
            ASSvcAccount          = $SqlServiceCredential
            SQLSysAdminAccounts   = $SqlAdministratorCredential.UserName
            ASSysAdminAccounts    = $SqlAdministratorCredential.UserName
            InstallSharedDir      = "D:\Program Files\Microsoft SQL Server"
            InstallSharedWOWDir   = "D:\Program Files (x86)\Microsoft SQL Server"
            InstanceDir           = "D:\Program Files\Microsoft SQL Server"
            InstallSQLDataDir     = "D:\Program Files\Microsoft SQL Server\MSSQL15.$Instance\MSSQL\Data"
            SQLUserDBDir          = "D:\Program Files\Microsoft SQL Server\MSSQL15.$Instance\MSSQL\Data"
            SQLUserDBLogDir       = "D:\Program Files\Microsoft SQL Server\MSSQL15.$Instance\MSSQL\Data"
            SQLTempDBDir          = "D:\Program Files\Microsoft SQL Server\MSSQL15.$Instance\MSSQL\Data"
            SQLTempDBLogDir       = "D:\Program Files\Microsoft SQL Server\MSSQL15.$Instance\MSSQL\Data"
            SQLBackupDir          = "D:\Program Files\Microsoft SQL Server\MSSQL15.$Instance\MSSQL\Backup"
            ASConfigDir           = "D:\MSSQL15.$Instance\Config"
            ASDataDir             = "D:\MSSQL15.$Instance\Data"
            ASLogDir              = "D:\MSSQL15.$Instance\Log"
            ASBackupDir           = "D:\MSSQL15.$Instance\Backup"
            ASTempDir             = "D:\MSSQL15.$Instance\Temp"
            SourcePath            = "D:\Sources\SQLServer2019"
            UpdateEnabled         = "False"
            ForceReboot           = $false
            BrowserSvcStartupType = "Automatic"

            PsDscRunAsCredential  = $SqlInstallCredential

            DependsOn             = "[WindowsFeature]NetFramework"
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

SQLServer2019 -ConfigurationData $config -OutputPath $Dir
Start-DscConfiguration -Path $Dir -Verbose -Wait -Force -ComputerName $Computer -ErrorAction Stop


Import-Module -Name dbatools

$SqlInstance = "$Computer\$Instance"

# Set TCP Port 1433 for ALL IP Addresses
Set-DbaTcpPort -SqlInstance $SqlInstance -Port 1433 -Confirm:$false

# Disable logging successful backups
Set-DbaStartupParameter -SqlInstance $SqlInstance -TraceFlag 3226 -TraceFlagOverride -Confirm:$false -Force

# Restart the service for the TCP & TraceFlag settings to take effect
Restart-DbaService -ComputerName $Computer -InstanceName $Instance -Type Engine -Force

# Disable the SA account, create a new sql login with SA role if needed
Set-DbaLogin -SqlInstance $SqlInstance -Login 'sa' -Disable

# Update Model DB settings
Set-DbaDbRecoveryModel -SqlInstance $SqlInstance -RecoveryModel Simple -Database model -Confirm:$False -Verbose
Invoke-Sqlcmd -ServerInstance $SqlInstance -Database 'master' -Query "ALTER DATABASE [model] MODIFY FILE ( NAME = N'modeldev', FILEGROWTH = 262144KB )"
Invoke-Sqlcmd -ServerInstance $SqlInstance -Database 'master' -Query "ALTER DATABASE [model] MODIFY FILE ( NAME = N'modellog', FILEGROWTH = 262144KB )"
Invoke-Sqlcmd -ServerInstance $SqlInstance -Database 'master' -Query "ALTER DATABASE [model] SET PAGE_VERIFY CHECKSUM  WITH NO_WAIT"
Invoke-Sqlcmd -ServerInstance $SqlInstance -Database 'master' -Query "ALTER DATABASE [model] SET AUTO_CLOSE OFF WITH NO_WAIT"
Invoke-Sqlcmd -ServerInstance $SqlInstance -Database 'master' -Query "ALTER DATABASE [model] SET AUTO_SHRINK OFF WITH NO_WAIT"

# Set the maxdop based on recommend value
Set-DbaMaxDop -SqlInstance $SqlInstance -MaxDop 1

# Set the maxmemory based on recommend value
Set-DbaMaxMemory -SqlInstance $SqlInstance

# Set cost threshold for parallelism
Set-DbaSpConfigure -SqlInstance $SqlInstance -ConfigName costthresholdforparallelism -Value 50

# Set the DBA error log to 12 files
Set-DbaErrorLogConfig -SqlInstance $SqlInstance -LogCount 12

# Restart the service for new tempdb settings to take effect
Restart-DbaService -ComputerName $Computer -InstanceName $Instance -Type Engine -Force
