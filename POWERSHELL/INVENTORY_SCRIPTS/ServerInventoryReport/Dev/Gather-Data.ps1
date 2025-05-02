# Begin Region Data Gathering

$InventoryDate = Get-Date
$ServerGroup = 'MemberServer'

    Get-Server -MemberServer | Start-RSJob -Name {$_} -FunctionsToLoad `
    Get-ScheduledTask,
    Invoke-SQLCmd,
    Get-LocalGroup,
    Get-LocalUser,
    Get-SecurityUpdate,
    Get-Software,
    Get-AdminShare,
    Get-UserShareDACL,
    Convert-ChassisType,
    Write-DataTable,
    Out-DataTable
    -ScriptBlock {
    
    Write-Verbose "[$($_)] - Initializing" -Verbose
    # Begin Region Variables
    $Computername = $_
    $SQLParams = @{
        Computername = $Using:SQLServer
        Database = 'SITEST'
        CommandType = 'NonQuery'
        ErrorAction = 'Stop'
        SQLParameter = @{
            '@Computername' = $Computername
        }
        Verbose = $True
    }
    $Date = $Using:InventoryDate
    # End Region Variables

    # Begin Region Lookups
    $DomainRole = @{
        0x0 = 'Standalone Workstation'
        0x1 = 'Member Workstation'
        0x2 = 'Standalone Server'
        0x3 = 'Member Server'
        0x4 = 'Backup Domain Controller'
        0x5 = 'Primary Domain Controller'
    }
    $DriveType = @{
        0x0 = 'N/A'
        0x1 = 'No Root Directory'
        0x2 = 'Removable Disk'
        0x3 = 'Local Disk'
        0x4 = 'Network Drive'
        0x5 = 'Compact Disk'
        0x6 = 'RAM Disk'
    }
    $GroupType = @{
        0x2 = 'Global Group'
        0x4 = 'Local Group'
        0x8 = 'Universal Group'
        2147483648 = 'Security Enabled'
    }
    $ShareType = @{
        0x0 = 'Disk Drive'
        0x1 = 'Print Queue'
        0x2 = 'Device'
        0x3 = 'IPC'
        2147483648 = 'Disk Drive Admin'
        2147483647 = 'Print Queue Admin'
        2147483646 = 'Device Admin'
        2147483651 = 'IPC Admin'
    }
    $AceType = @{
        0x0 = 'Access Allowed'
        0x1 = 'Access Denied'
        0x2 = 'Audit'
    }
    $AceFlags = @{
        0x1 = 'OBJECT_INHERIT_ACE'
        0X2 = 'CONTAINER_INHERIT_ACE'
        0X4 = 'NO_PROPAGATE_ACE'
        0X8 = 'INHERIT_ONLY_ACE'
        0X10 = 'INHERITED_ACE'
    }
    $AccessMask = @{
        0x1F01FF = "FullControl"
        0x120089 = "Read"
        0x12019F = "Read, Write"
        0x1200A9 = "ReadAndExecute"
        1610612736 = "ReadAndExecuteExtended"
        0x1301BF = "ReadAndExecute, Modify, Write"
        0x1201BF = "ReadAndExecute, Write"
        0x10000000 = "FullControl (Sub Only)"
    }
    $ProcessorType = @{
        0x1 = 'Other'
        0x2 = 'N/A'
        0x3 = 'Central Processor'
        0x4 = 'Math Processor'
        0x5 = 'DSP Processor'
        0x6 = 'Video Processor'
    }
    $TypeDetail = @{
        0x1 = 'Reserved'
        0x2 = 'Other'
        0x4 = 'N/A'
        0x8 = 'Fast-paged'
        0x16 = 'Static column'
        0x32 = 'Pseudo-static'
        0x64 = 'RAMBUS'
        0x128 = 'Synchronous'
        0x256 = 'CMOS'
        0x512 = 'EDO'
        0x1024 = 'Window DRAM'
        0x2048 = 'Cache DRAM'
        0x4096 = 'Nonvolatile'
    }
    $MemoryType = @{
        0x0 = 'N/A'
        0x1 = 'Other'
        0x2 = 'DRAM'
        0x3 = 'Synchronous DRAM'
        0x4 = 'Cache DRAM'
        0x5 = 'EDO'
        0x6 = 'EDRAM'
        0x7 = 'VRAM'
        0x8 = 'SRAM'
        0x9 = 'RAM'
        0x10 = 'ROM'
        0x11 = 'Flash'
        0x12 = 'EEPROM'
        0x13 = 'FEPROM'
        0x14 = 'EPROM'
        0x15 = 'CDRAM'
        0x16 = '3DRAM'
        0x17 = 'SDRAM'
        0x18 = 'SGRAM'
        0x19 = 'RDRAM'
        0x20 = 'DDR'
        0x21 = 'DDR-2'
        0x22 = 'DDR-2 FB-DIMM'
        0x24 = 'DDR-3'
        0x25 = 'FBD2'
    }
    $DebugType = @{
        0x0 = 'None'
        0x1 = 'Complete Memory Dump'
        0x2 = 'Kernel Memory Dump'
        0x3 = 'Small Memory Dump'
    }
    $AUOptions = @{
        0x2 = 'Notify before download'
        0x3 = 'Automatically download and notify of installation.'
        0x4 = 'Automatic download and scheduled installation.'
        0x5 = 'Automatic Updates is required, but end users can configure it.'
    }
    $DefaultNetworkRole = @{
        0x0 = 'ClusterNetworkRoleNone'
        0x1 = 'ClusterNetworkRoleInternalUse'
        0x2 = 'ClusterNetworkRoleClientAccess'
        0x3 = 'ClusterNetworkRoleInternalAndClient'
    }
    $ClusState = @{
        -1 = 'StateUnknown'
        0x0 = 'Inherited'
        0x1 = 'Initializing'
        0x2 = 'Online'
        0x3 = 'Offline'
        0x4 = 'Failed'
        0x80 = 'Pending'
        0x81 = 'Online Pending'
        0x82 = 'Offline Pending'
    }
    $NetState = @{
        -1 = 'StateUnknown'
        0x0 = 'Unavailable'
        0x1 = 'Failed'
        0x2 = 'Unreachable'
        0x3 = 'Up'
    }
    $FailbackType = @{
        0x0 = 'ClusterGroupPreventFailback'
        0x1 = 'ClusterGroupAllowFailback'
    }
    # End Region Lookups
    
    # Begin Region General
    Try {
        $CS = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $Computername -ErrorAction Continue
        $SrvApp = Get-ADComputer -Identity $Computername -Properties Description -ErrorAction Continue
        $Enclosure = Get-CimInstance -ClassName Win32_SystemEnclosure -ComputerName $Computername
        $BIOS = Get-CimInstance -ClassName Win32_Bios -ComputerName $Computername 
        $General = [pscustomobject]@{
            Computername = $Computername
            ServerApplication = $SrvApp.Description
            Domain = $CS.Domain
            Manufacturer = $CS.Manufacturer
            Model = $CS.Model
            SystemType = $CS.SystemType
            SerialNumber = $Enclosure.SerialNumber
            ChassisType = (Convert-ChassisType $Enclosure.ChassisTypes)
            Description = $Enclosure.Description
            BIOSManufacturer = $BIOS.Manufacturer
            BIOSName = $BIOS.Name
            BIOSSerialNumber = $BIOS.SerialNumber
            BIOSVersion = $BIOS.SMBIOSBIOSVersion
            InventoryDate = $Date
        }
        If ($General) {
            Write-Verbose "[$Computername - $(Get-Date)] Removing old data" -Verbose
            $SQLParams.CommandType = 'NonQuery'
            $SQLParams.TSQL = "DELETE FROM tbGeneral WHERE Computername = @Computername"            
            Invoke-SQLCmd @SQLParams
            If ($Return.DefaultView) {
                Write-Verbose "[$Computername - $(Get-Date)] Throwing error" -Verbose
                Throw 'FAIL'
            }
            Else {
                $SQLParams.CommandType = 'NonQuery'
                $DataTable = $General | Out-DataTable
                Write-Verbose "[$Computername - $(Get-Date)] Updating data" -Verbose
                Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbGeneral -Data $DataTable -ErrorAction Continue
            }
        }
    }
    Catch {
        Write-Verbose "WARNING - [$Computername - $(Get-Date)]" -Verbose
        Write-Warning $_
        BREAK
    }
    # End Region General

    # Begin Region OperatingSystem
    ## TODO: Include Product Key?
    $TimeZone = Get-CimInstance -ClassName Win32_TimeZone -ComputerName $Computername -ErrorAction Continue
    $OS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $Computername 
    $PageFile = Get-CimInstance -ClassName win32_PageFile -ComputerName $Computername 
    $LastReboot = Try {
        $OS.ConvertToDateTime($OS.LastBootUpTime)
    } 
    Catch {
    }
    $OperatingSystem = [pscustomobject]@{
        Computername = $Computername
        Caption = $OS.caption
        Version = $OS.version
        ServicePack = ("{0}.{1}" -f $OS.ServicePackMajorVersion, $OS.ServicePackMinorVersion)
        LastReboot = $LastReboot
        OSArchitecture = $OS.OSArchitecture
        TimeZone = $TimeZone.Caption
        PageFile = $PageFile.Name
        PageFileSizeGB = ("{0:N2}" -f ($PageFile.FileSize /1GB))
        InventoryDate = $Date
    }
    If ($OperatingSystem) {
        $SQLParams.TSQL = "DELETE FROM tbOperatingSystem WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $OperatingSystem | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbOperatingSystem -Data $DataTable
    }
    # End Region OperatingSystem

    # Begin Region Memory
    $Memory = @(Get-CimInstance -ClassName Win32_PhysicalMemory -ComputerName $Computername -ErrorAction Continue| ForEach-Object {
        [pscustomobject]@{
            Computername = $Computername
            DeviceID = $_.tag
            MemoryType = $MemoryType[[int]$_.MemoryType]
            "Capacity(GB)" = "{0}" -f ($_.capacity/1GB)
            TypeDetail = $TypeDetail[[int]$_.TypeDetail]
            Locator = $_.DeviceLocator
            InventoryDate = $Date
        }
    })

    If ($Memory) {
        $SQLParams.TSQL = "DELETE FROM tbMemory WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $Memory | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbMemory -Data $DataTable
    }
    # End Region Memory

    # Begin Region Network
    $Network = @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $Computername -ErrorAction Continue) | ForEach-Object {
        [pscustomobject]@{
            Computername = $Computername
            DeviceName = $_.Caption
            DHCPEnabled = $_.DHCPEnabled
            MACAddress = $_.MACAddress    
            IPAddress = ($_.IpAddress  -join '; ')
            SubnetMask = ($_.IPSubnet  -join '; ')
            DefaultGateway = ($_.DefaultIPGateway -join '; ')
            DNSServers = ($_.DNSServerSearchOrder  -join '; ')
            InventoryDate = $Date
        }
    }

    If ($Network) {
        $SQLParams.TSQL = "DELETE FROM tbNetwork WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $Network | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbNetwork -Data $DataTable
    }
    # End Region Network

    # Begin Region CPU
    $Processor = @(Get-CimInstance -ClassName Win32_Processor -ComputerName $Computername -ErrorAction Continue | ForEach-Object {
        [pscustomobject]@{
            Computername = $Computername
            DeviceID = $_.DeviceID
            Description = $_.Description
            ProcessorType = $ProcessorType[$_.processortype]
            CoreCount = $_.NumberofCores
            NumLogicalProcessors = $_.NumberOfLogicalProcessors
            MaxSpeed = ("{0:N2} GHz" -f ($_.MaxClockSpeed/1000))
            InventoryDate = $Date
        }
    })

    If ($Processor) {
        $SQLParams.TSQL = "DELETE FROM tbProcessor WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $Processor | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbProcessor -Data $DataTable
    }
    # End Region CPU

    # Begin Region Drives
    $Disk = @(Get-CimInstance -ClassName Win32_Volume -Filter "(Not Name LIKE '\\\\?\\%')" -ComputerName $Computername -ErrorAction Continue | ForEach-Object {
        [pscustomobject]@{
            Computername = $Computername
            Drive = $_.Name
            DriveType = $DriveType[[int]$_.DriveType]
            Label = $_.label
            FileSystem = $_.FileSystem
            FreeSpaceGB = "{0:N2}" -f ($_.FreeSpace /1GB)
            CapacityGB = "{0:N2}" -f ($_.Capacity/1GB)
            PercentFree = ($_.FreeSpace/$_.Capacity)
            InventoryDate = $Date
        }
    })

    If ($Disk) {
        $SQLParams.TSQL = "DELETE FROM tbDrives WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $Disk | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbDrives -Data $DataTable
    }
    # End Region Drives

    # Begin Region AdminShares
    $AdminShare = @(Get-AdminShare -Computername $Computername -ErrorAction Continue | ForEach-Object {
        [pscustomobject]@{
            Computername = $Computername
            Name = $_.Name
            Path = $_.Path
            Type = $_.Type  
            InventoryDate = $Date  
        }
    })

    If ($AdminShare) {
        $SQLParams.TSQL = "DELETE FROM tbAdminShare WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $AdminShare | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbAdminShare -Data $DataTable
    }
    # End Region AdminShares

    # Begin Region UserShares
    $UserShare = Try {
    Get-UserShareDACL -Computername $Computername -ErrorAction Continue | Select-Object *,@{L='InventoryDate';E={$Date}}
    } 
    Catch {}
    If ($UserShare) {
        $SQLParams.TSQL = "DELETE FROM tbUserShare WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $UserShare | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbUserShare -Data $DataTable
    }
    # End Region UserShares

    # Begin Region Local Users
    If ($Using:ServerGroup -eq 'MemberServer') {
        $Users = @(Get-LocalUser -ComputerName $Computername -ErrorAction Continue) | Select-Object *,@{L='InventoryDate';E={$Date}}

        If ($Users) {
            $SQLParams.TSQL = "DELETE FROM tbUsers WHERE Computername = @Computername"
            Invoke-SQLCmd @SQLParams
            $DataTable = $Users | Out-DataTable
            Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbUsers -Data $DataTable
        }
    }
    # End Region Local Users

    # Begin Region Local Groups
    If ($Using:ServerGroup -eq 'MemberServer') {
        $Groups = @(Get-LocalGroup -ComputerName $Computername -ErrorAction Continue) | Select-Object *,@{L='InventoryDate';E={$Date}}

        If ($Groups) {
            $SQLParams.TSQL = "DELETE FROM tbGroups WHERE Computername = @Computername"
            Invoke-SQLCmd @SQLParams
            $DataTable = $Groups | Out-DataTable
            Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbGroups -Data $DataTable
        }
    }
    # End Region Local Groups

    # Begin Region Server Roles
    $ServerRoles = Try {
        Get-CimInstance -ClassName Win32_ServerFeature -ComputerName $Computername -ErrorAction Continue | ForEach-Object {
            [pscustomobject]@{
                Computername = $Computername
                ID = $_.Id
                Name = $_.Name
                InventoryDate = $Date
            }
        }
    } 
    Catch {}

    If ($ServerRoles) {
        $SQLParams.TSQL = "DELETE FROM tbServerRoles WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $ServerRoles | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbServerRoles -Data $DataTable
    }
    # End Region Server Roles

    # Begin Region Scheduled Tasks
    $ScheduledTasks = Get-ScheduledTask -Computername $Computername -ErrorAction Continue | Select-Object *,@{L='InventoryDate';E={$Date}}

    If ($ScheduledTasks) {
        $SQLParams.TSQL = "DELETE FROM tbScheduledTasks WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $ScheduledTasks | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbScheduledTasks -Data $DataTable
    }
    # End Region Scheduled Tasks

    # Begin Region Software
    $Software = Get-Software -Computername $Computername -ErrorAction Continue | Sort-Object DisplayName | Select-Object *,@{L='InventoryDate';E={$Date}}

    If ($Software) {
        $SQLParams.TSQL = "DELETE FROM tbSoftware WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $Software | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbSoftware -Data $DataTable
    }
    # End Region Software

    # Begin Region Updates
    $Updates = Get-SecurityUpdate -Computername $Computername -ErrorAction Continue | 
        Select-Object @{L='Computername';E={$Computername}},Description, HotFixID, InstalledOn, Type,@{L='InventoryDate';E={$Date}} | 
        Group-Object HotFixID | ForEach-Object {$_.Group | Sort-Object -Unique DisplayName}
    $Hotfixes = Get-HotFix -ComputerName $Computername | ForEach-Object {
        Switch -Wildcard ($_.Description) {
            "Service Pack*" {$Type = 'Service Pack'}
            "Hotfix*" {$Type = 'Hotfix'}
            "Update*" {$Type = 'Update'}
            "Security Update*" {$Type = 'Security Update'}
            Default {$Type = 'N/A'}
        }
        [pscustomobject]@{
            Computername = $Computername
            Description = $_.Description
            HotFixID = $_.HotFixID
            InstalledOn = $_.InstalledOn
            Type = $Type
            InventoryDate = $Date
        }
    } 
    $TotalUpdates = $Hotfixes + $Updates

    If ($TotalUpdates) {
        $SQLParams.TSQL = "DELETE FROM tbUpdates WHERE Computername = @Computername"
        Invoke-SQLCmd @SQLParams
        $DataTable = $TotalUpdates | Out-DataTable
        Write-DataTable -Computername $Using:SQLServer -Database SITEST -TableName tbUpdates -Data $DataTable
    }
    # End Region Updates
} | Wait-RSJob -ShowProgress | Remove-RSJob
# End Region Data Gathering