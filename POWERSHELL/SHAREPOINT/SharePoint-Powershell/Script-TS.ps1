Asnp *sh* # same as Add-PSSnapin microsoft.sharepoint.powershell -ea SilentlyContinue
Function Perform-ConfigRefresh
{
    param
    (
        [parameter(Mandatory=$false)][String[]]$ServerList
    )
    
    [String[]]$serverlistInternal = $null
    if ( $ServerList -ne $null )
    {
        $serverlist | % { [String[]]$serverlistInternal = $serverlistInternal + $_.ToUpper() }
    }
 
    try
    {
        $farm = Get-SPFarm
 
    }
    catch
    {
        $_.exception.message
        break
    }
    if ( -not $farm.CurrentUserIsAdministrator() )
    {
        Write-host "Please run with account that is farmAdmin and localAdminRights on server" -ForegroundColor Yellow
    }
 
    $configGUID = (Get-SPDatabase | where { $_.Type -eq "Configuration Database"}).id.tostring()
    $programData = $env:ProgramData
    $drive = $programData.Split(':')[0]
 
 
    foreach ( $instance in $farm.TimerService.Instances )
    {
        if ( $serverlistInternal -eq $null -or $serverlistInternal.Contains($instance.Server.Name.ToUpper()))
        {
            try
            {
                Write-host "Performing Config refresh on Server : $($instance.server.Name.ToUpper())" `n
                $configpath = "\\$($instance.Server.Name)\$drive`$\ProgramData\Microsoft\SharePoint\Config\$configGUID"
                Write-host `t "Stopping Timer Service on Server..." -NoNewline 
                Set-Service -ComputerName $instance.Server.Name -Name SPTimerV4 -Status Stopped -ErrorAction Stop -WarningAction SilentlyContinue
                Write-Host "Done."
                Write-host `t "Backing up Config Cache..." -NoNewline
                Rename-Item -Path $configpath  -NewName ($configpath + "."+"Bak-"+"$([datetime]::now.ToFileTime().ToString())") -ErrorAction Stop
                Write-host "Done."
                Write-host `t "Clearing Cache and Starting Service..." -NoNewline
                mkdir $Configpath -ErrorAction Stop | Out-Null
                Set-Content -Value 1 -Path "$Configpath\cache.ini" -Encoding UTF8 -ErrorAction Stop
                Set-Service -ComputerName $instance.Server.Name -Name SPTimerV4 -Status Running -ErrorAction Stop  -WarningAction SilentlyContinue
                Write-host "Done." `n
            }
            catch 
            {
                Write-host `t "Failed" -ForegroundColor Red
                Write-host `t "Error :" $_.exception.Message -ForegroundColor Red
                Write-host "Config Refresh Failed on Server $($instance.server.name.ToUpper())" -ForegroundColor Red
            }
        }
     }
     Write-host "Please give Config refresh couple of minutes to populate the cache fully." -ForegroundColor Yellow 
} 
Perform-ConfigRefresh 
