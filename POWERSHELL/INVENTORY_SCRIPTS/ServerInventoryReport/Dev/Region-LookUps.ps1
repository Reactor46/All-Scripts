﻿#region Lookups
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
    #endregion Lookups