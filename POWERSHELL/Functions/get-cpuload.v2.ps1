#requires -RunAsAdministrator

# fake reading in a list of computer names
#    in real life, use Get-Content or (Get-ADComputer).Name
$ComputerList = @(
'vdcmaddc01.corp.vegas.com',
'stg-theoapp02.corp.vegas.com',
'VVCNETAP02.corp.vegas.com',
'VDCNETAP01.corp.vegas.com',
'vvcexcwa01.corp.vegas.com',
'VVCEXCMB01.corp.vegas.com',
'VVCEXCWA02.corp.vegas.com',
'VVCEXCMB02.corp.vegas.com',
'VVCPRTAP01.corp.vegas.com',
'VVDBHRAP01.corp.vegas.com',
'vvsbhrap02.corp.vegas.com',
'VVSVDCDB01.corp.vegas.com'
    )
$IC_ScriptBlock = {
    $CIM_ComputerSystem = Get-CimInstance -ClassName CIM_ComputerSystem
    $CIM_BIOSElement = Get-CimInstance -ClassName CIM_BIOSElement
    $CIM_OperatingSystem = Get-CimInstance -ClassName CIM_OperatingSystem
    $CIM_Processor = Get-CimInstance -ClassName CIM_Processor
    $CIM_LogicalDisk = Get-CimInstance -ClassName CIM_LogicalDisk |
        Where-Object {$_.Name -eq $CIM_OperatingSystem.SystemDrive}

    [PSCustomObject]@{
        LocalComputerName = $env:COMPUTERNAME
        Manufacturer = $CIM_ComputerSystem.Manufacturer
        Model = $CIM_ComputerSystem.Model
        SerialNumber = $CIM_BIOSElement.SerialNumber
        CPU = $CIM_Processor.Name
        RAM_GB = '{0:N2}' -f ($CIM_ComputerSystem.TotalPhysicalMemory / 1GB)
        SysDrive_Capacity_GB = '{0:N2}' -f ($CIM_LogicalDisk.Size / 1GB)
        SysDrive_FreeSpace_GB ='{0:N2}' -f ($CIM_LogicalDisk.FreeSpace / 1GB)
        SysDrive_FreeSpace_Pct = '{0:N0}' -f ($CIM_LogicalDisk.FreeSpace / $CIM_LogicalDisk.Size * 100)
        OperatingSystem_Name = $CIM_OperatingSystem.Caption
        OperatingSystem_Version = $CIM_OperatingSystem.Version
        OperatingSystem_BuildNumber = $CIM_OperatingSystem.BuildNumber
        OperatingSystem_ServicePack = $CIM_OperatingSystem.ServicePackMajorVersion
        CurrentUser = $CIM_ComputerSystem.UserName
        LastBootUpTime = $CIM_OperatingSystem.LastBootUpTime
        UpTime_Days = '{0:N2}' -f ([datetime]::Now - $CIM_OperatingSystem.LastBootUpTime).Days
        }

    }

$IC_Params = @{
    ComputerName = $ComputerList
    ScriptBlock = $IC_ScriptBlock
    ErrorAction = 'SilentlyContinue'
    }
$RespondingSystems = Invoke-Command @IC_Params
$NOT_RespondingSystems = $ComputerList.Where({
    # these two variants are needed to deal with an ipv6 localhost address
    "[$_]" -notin $RespondingSystems.PSComputerName -and
    $_ -notin $RespondingSystems.PSComputerName
    })

$RespondingSystems
$NOT_RespondingSystems
