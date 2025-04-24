# Begin Script 2
Import-Module WebAdministration

$Sites = Get-Website
$TimeStamp = Get-date -Format "MM-dd-yy-HHmmss"
$BackupPath = "C:\Temp"

# Backup site web.config
ForEach($Site In $Sites){
    $ConfigPath = (Get-WebConfigFile -PSPath "IIS:\Sites\$($Site.Name)").FullName

    Write-Host "Backing up $ConfigPath"
    Copy-Item $ConfigPath "$BackupPath\$TimeStamp-$($Site.Name)-web.config"
}

# Backup .NET Global configs
$DotnetPaths = ("Framework\v2.0.50727\Config",
                "Framework\v4.0.30319\Config",
                "Framework64\v2.0.50727\Config",
                "Framework64\v4.0.30319\Config")

$ConfigTypes = ("machine.config", "web.config")

# Backup Global configs
ForEach($DotnetPath In $DotnetPaths){
    ForEach($ConfigType In $ConfigTypes){
        #$SitePath = [System.Environment]::ExpandEnvironmentVariables("$($Site.physicalPath)\web.config")
        $ConfigPath = "C:\Windows\Microsoft.NET\$DotnetPath\$ConfigType"

        If(Test-Path -Path $ConfigPath){
            Write-Host "Backing up $ConfigPath"
            Copy-Item $ConfigPath "$BackupPath\$TimeStamp-$($DotnetPath -replace "\\", "-")-$ConfigType"
        }
    }
}

# Backup Administration Config
$AdministrationConfig = "C:\Windows\System32\inetsrv\Config\administration.config"
If(Test-Path -Path $AdministrationConfig){
    Write-Host "Backing up $AdministrationConfig"
    Copy-Item $AdministrationConfig "$BackupPath\$TimeStamp-administration.config"
}

# Backup Application Host Config
$ApplicationHostConfig = "C:\Windows\System32\inetsrv\Config\applicationHost.config"
If(Test-Path -Path $ApplicationHostConfig){
    Write-Host "Backing up $ApplicationHostConfig"
    Copy-Item $ApplicationHostConfig "$BackupPath\$TimeStamp-applicationHost.config"
}
# End Script 2
