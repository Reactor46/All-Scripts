function Get-WACRDDesktopAvailability {
<#

.SYNOPSIS
Gets the availability of the installed desktop GUI features to enable Remote Applications

.DESCRIPTION
Gets the availability of the installed desktop GUI features across all Operating System types. This information is
is used as a basis for enabling connections to Remote Apps.

.ROLE
Readers

#>

Set-Variable StateAvailable -Option Constant -Value 'Available' -ErrorAction SilentlyContinue
Set-Variable MessageAvailable -Option Constant -Value 'Remote App tool is enabled and ready' -ErrorAction SilentlyContinue

Set-Variable StateNotSupported -Option Constant -Value 'NotSupported' -ErrorAction SilentlyContinue
Set-Variable MessageNotSupported -Option Constant -Value 'Remote App tool is not supported for this system configuration' -ErrorAction SilentlyContinue

Set-Variable RegistryKey -Option Constant -Value 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
Set-Variable PropertyName -Option Constant -Value 'InstallationType' -ErrorAction SilentlyContinue

# Gets the value of the installed operating system type. Returns a string, eg: 'Server Core'
function GetOperatingSystemType () {
    return Get-ItemProperty -Path $RegistryKey -Name $PropertyName | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty $PropertyName
}

# Checks to see if the Desktop UI is installed on the OS. Builds an 'Available' response if it is.
function CheckIfFeatureInstalled ([string] $oSTypeString) {
    switch ($oSTypeString) {
        'Client' {
            return BuildResponseObject -State $StateAvailable -Message $MessageAvailable;
        }
        'Server' {
            if (Get-Command Get-WindowsFeature -errorAction SilentlyContinue) {
                $DesktopFeature = Get-WindowsFeature -Name Desktop-Experience
                if ($DesktopFeature.Installed) {
                    return BuildResponseObject -State $StateAvailable -Message $MessageAvailable;
                }
            }
        }
    }
}

# The object that is returned to the inventory loader when a tool should be available or not. 
function BuildResponseObject ([string] $state, [string] $message) {
    $ResponseObject = @{
        State      = $state;
        Message    = $message;
        Properties = @{};
    }

    return $ResponseObject
}

# Checks the OS type, and builds a response object based on whether the desktop UI is available or not.
function GetAvailability() {
    $OperatingSystemType = GetOperatingSystemType
    $Response = CheckIfFeatureInstalled -oSTypeString $OperatingSystemType
    if (!$Response) {
        $Response = BuildResponseObject -state $StateNotSupported -message $MessageNotSupported
    }

    return $Response
}

return GetAvailability

}
## [END] Get-WACRDDesktopAvailability ##
function Get-WACRDEnhancedSessionState {
<#

.SYNOPSIS
Gets the enumerated Enhanced Session state from a VM

.DESCRIPTION
Gets the enumerated Enhanced Session state between a VM and it's host system by querying the Msvm_Computersystem class in the cim instance.
It's enumeration values can be found at https://msdn.microsoft.com/en-us/library/windows/desktop/hh850116(v=vs.85).aspx

.ROLE
Readers

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]
    $vmGuid
)

function GetEnhancedSessionState([string] $vmGuid) {
    $state = Get-CimInstance  -Namespace Root\Virtualization\V2 -ClassName Msvm_ComputerSystem | 
        Where-Object {$_.Name -eq $vmGuid} | 
        Microsoft.PowerShell.Utility\Select-Object -ExpandProperty EnhancedSessionModeState
    
    return $state
}

GetEnhancedSessionState $vmGuid
}
## [END] Get-WACRDEnhancedSessionState ##
function Get-WACRDIfFileExists {
<#

.SYNOPSIS
Checks the target server node to see if the selected app exists before launching it.

.DESCRIPTION
Checks the target server node to see if the selected app exists before launching it. This is used
as verification that the original file location hasn't changed and that the file is ready for use.

.ROLE
Readers

#>

Param(
    [string]$selectedFilePath
)

Test-Path -Path $selectedFilePath
}
## [END] Get-WACRDIfFileExists ##
function Get-WACRDRemoteAppIcon {
<#

.SYNOPSIS
Gets an embedded Icon from a selected executable file.

.DESCRIPTION
Gets an embedded Icon from a selected executable file by extracting it to a bitmap
and then converting it to a base64 string. This is useful for transferring the graphical
information as direct HTML data.

.ROLE
Readers

#>

Param(
    # [string]$fileName,
    [string]$selectedFilePath
)

Add-Type -AssemblyName System.Drawing
$Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($selectedFilePath)
$MemoryStream = New-Object System.IO.MemoryStream
$Icon.ToBitmap().Save($MemoryStream, [System.Drawing.Imaging.ImageFormat]::Png)
$Bytes = $MemoryStream.ToArray()   
$MemoryStream.Flush() 
$MemoryStream.Dispose()
[convert]::ToBase64String($Bytes)
}
## [END] Get-WACRDRemoteAppIcon ##
function Get-WACRDRemoteApplicationSettings {
<#

.SYNOPSIS
Gets the Remote App tool allow setting on the target node.

.DESCRIPTION
Gets the Remote App tool allow setting on the target node. If enabled, returns an object with the 'Available' state. Used in manifest
to dynamically enable/disable the Remote App tool from the inventory.

.ROLE
Readers

#>

$response = @{
    state = 'NotConfigured';
    message = 'Remote App tool is not enabled in settings';
}

$key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server\TSAppAllowList'

$exists = Get-ItemProperty -Path $key -Name fDisabledAllowList -ErrorAction SilentlyContinue

if($exists.fDisabledAllowList) {
    $response.state = 'Available';
    $response.message = 'Remote App is enabled and ready';
}

$response

}
## [END] Get-WACRDRemoteApplicationSettings ##
function Get-WACRDRemoteDesktopAccountLockout {
<#

.SYNOPSIS
Gets the Remote Access lockout settings for the target node.

.DESCRIPTION
Gets the Remote Access lockout settings for the target node. Retrieves information on reset time
and max lockouts for the node. If a failure occurs, error message is returned.

.ROLE
Readers

#>

try {
    $accessSettings = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\RemoteAccess\Parameters\AccountLockout" -ErrorAction Stop
} catch {
    return @{
        success = $false
        error = $_.Exception.Message
    }
}

return @{
    success = $true
    maxDenials = $accessSettings.MaxDenials
    resetTime = $accessSettings.'ResetTime (mins)'
}
}
## [END] Get-WACRDRemoteDesktopAccountLockout ##
function Get-WACRDRemoteDesktopCertificate {
<#

.SYNOPSIS
Gets the Remote Desktop certificate on the target node.

.DESCRIPTION
Gets the Remote Desktop certificate on the target node.

.ROLE
Readers

#>

$thumbprint = Get-CimInstance -Class " Win32_TSGeneralSetting " -Namespace root\cimv2\terminalservices | Where-Object {$_.TerminalName -EQ "RDP-Tcp"} | `
    Microsoft.PowerShell.Utility\Select-Object -ExpandProperty SSLCertificateSHA1Hash
$cert = Get-ChildItem -Path "Cert:\LocalMachine\*$thumbprint" -Recurse | Microsoft.PowerShell.Utility\Select-Object -First 1

$result = @{
    "FingerprintSha1" = $cert.Thumbprint;
    "Subject" = $cert.Subject;
    "Issuer" = $cert.Issuer;
    "ExtendedUsage" = $cert.EnhancedKeyUsageList.FriendlyName; # Not sure if this is everu an array
    "ValidFrom" = $cert.NotBefore;
    "ValidTo" = $cert.NotAfter;
}
 
return $result

}
## [END] Get-WACRDRemoteDesktopCertificate ##
function Get-WACRDRemoteDesktopSettings {
<#

.SYNOPSIS
Gets the Remote Desktop setting on the target node.

.DESCRIPTION
Gets the Remote Desktop setting on the target node.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

Set-Variable -Option Constant -Name RdpSystemRegistryKey -Value "HKLM:\\SYSTEM\CurrentControlSet\Control\Terminal Server"
Set-Variable -Option Constant -Name RdpGroupPolicyProperty -Value "fDenyTSConnections" -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name RdpNlaGroupPolicyProperty -Value "UserAuthentication" -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name RdpGroupPolicyRegistryKey -Value "HKLM:\\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"
Set-Variable -Option Constant -Name RdpListenerRegistryKey -Value "$RdpSystemRegistryKey\WinStations"
Set-Variable -Option Constant -Name RdpProtocolTypeUM -Value "{5828227c-20cf-4408-b73f-73ab70b8849f}"
Set-Variable -Option Constant -Name RdpProtocolTypeKM -Value "{18b726bb-6fe6-4fb9-9276-ed57ce7c7cb2}"
Set-Variable -Option Constant -Name RdpWdfSubDesktop -Value 0x00008000
Set-Variable -Option Constant -Name RdpFirewallGroup -Value "@FirewallAPI.dll,-28752"

<#

.SYNOPSIS
Gets the Remote Desktop Network Level Authentication settings of the current machine.

.DESCRIPTION
Gets the Remote Desktop Network Level Authentication settings of the system.

.ROLE
Readers

#>
function Get-RdpNlaGroupPolicySettings {
    $nlaGroupPolicySettings = @{}
    $nlaGroupPolicySettings.GroupPolicyIsSet = $false
    $nlaGroupPolicySettings.GroupPolicyIsEnabled = $false
    $registryKey = Get-ItemProperty -Path $RdpGroupPolicyRegistryKey -ErrorAction SilentlyContinue
    if (!!$registryKey) {
        if ((Get-Member -InputObject $registryKey -name $RdpNlaGroupPolicyProperty -MemberType Properties) -and ($null -ne $registryKey.$RdpNlaGroupPolicyProperty)) {
            $nlaGroupPolicySettings.GroupPolicyIsSet = $true
            $nlaGroupPolicySettings.GroupPolicyIsEnabled = $registryKey.$RdpNlaGroupPolicyProperty -eq 1
        }
    }

    return $nlaGroupPolicySettings
}

<#

.SYNOPSIS
Gets the Remote Desktop settings of the system related to Group Policy.

.DESCRIPTION
Gets the Remote Desktop settings of the system related to Group Policy.

.ROLE
Readers

#>
function Get-RdpGroupPolicySettings {
    $rdpGroupPolicySettings = @{}
    $rdpGroupPolicySettings.GroupPolicyIsSet = $false
    $rdpGroupPolicySettings.GroupPolicyIsEnabled = $false
    $registryKey = Get-ItemProperty -Path $RdpGroupPolicyRegistryKey -ErrorAction SilentlyContinue
    if (!!$registryKey) {
        if ((Get-Member -InputObject $registryKey -name $RdpGroupPolicyProperty -MemberType Properties) -and ($null -ne $registryKey.$RdpGroupPolicyProperty)) {
            $rdpGroupPolicySettings.groupPolicyIsSet = $true
            $rdpGroupPolicySettings.groupPolicyIsEnabled = $registryKey.$RdpGroupPolicyProperty -eq 0
        }
    }

    return $rdpGroupPolicySettings
}

<#

.SYNOPSIS
Gets all of the valid Remote Desktop Protocol listeners.

.DESCRIPTION
Gets all of the valid Remote Desktop Protocol listeners.

.ROLE
Readers

#>
function Get-RdpListener {
    $listeners = @()
    Get-ChildItem -Name $RdpListenerRegistryKey | Where-Object { $_.PSChildName.ToLower() -ne "console" } | ForEach-Object {
        $registryKeyValues = Get-ItemProperty -Path "$RdpListenerRegistryKey\$_" -ErrorAction SilentlyContinue
        if ($registryKeyValues -ne $null) {
            $protocol = $registryKeyValues.LoadableProtocol_Object
            $isProtocolRDP = ($protocol -ne $null) -and ($protocol -eq $RdpProtocolTypeUM -or $protocol -eq $RdpProtocolTypeKM)

            $wdFlag = $registryKeyValues.WdFlag
            $isSubDesktop = ($wdFlag -ne $null) -and ($wdFlag -band $RdpWdfSubDesktop)

            $isRDPListener = $isProtocolRDP -and !$isSubDesktop
            if ($isRDPListener) {
                $listeners += $registryKeyValues
            }
        }
    }

    return ,$listeners
}

<#

.SYNOPSIS
Gets the number of the ports that the Remote Desktop Protocol is operating over.

.DESCRIPTION
Gets the number of the ports that the Remote Desktop Protocol is operating over.


.ROLE
Readers

#>
function Get-RdpPortNumber {
    $portNumbers = @()
    Get-RdpListener | Where-Object { $_.PortNumber -ne $null } | ForEach-Object { $portNumbers += $_.PortNumber }
    return ,$portNumbers
}

<#

.SYNOPSIS
Gets the Remote Desktop settings of the system.

.DESCRIPTION
Gets the Remote Desktop settings of the system.

.ROLE
Readers

#>
function Get-RdpSettings {
    $remoteDesktopSettings = New-Object -TypeName PSObject
    $rdpEnabledSource = $null
    $rdpIsEnabled = Test-RdpEnabled
    $rdpRequiresNla = Test-RdpUserAuthentication
    $rdpPortNumbers = Get-RdpPortNumber
    if ($rdpIsEnabled) {
        $rdpGroupPolicySettings = Get-RdpGroupPolicySettings
        if ($rdpGroupPolicySettings.groupPolicyIsEnabled) {
            $rdpEnabledSource = "GroupPolicy"
        } else {
            $rdpEnabledSource = "System"
        }
    }

    $remoteDesktopSettings | Add-Member -MemberType NoteProperty -Name "IsEnabled" -Value $rdpIsEnabled
    $remoteDesktopSettings | Add-Member -MemberType NoteProperty -Name "RequiresNLA" -Value $rdpRequiresNla
    $remoteDesktopSettings | Add-Member -MemberType NoteProperty -Name "Ports" -Value $rdpPortNumbers
    $remoteDesktopSettings | Add-Member -MemberType NoteProperty -Name "EnabledSource" -Value $rdpEnabledSource

    return $remoteDesktopSettings
}

<#

.SYNOPSIS
Tests whether Remote Desktop Protocol is enabled.

.DESCRIPTION
Tests whether Remote Desktop Protocol is enabled.


.ROLE
Readers

#>
function Test-RdpEnabled {
    $rdpEnabledWithGP = $false
    $rdpEnabledLocally = $false
    $rdpGroupPolicySettings = Get-RdpGroupPolicySettings
    $rdpEnabledWithGP = $rdpGroupPolicySettings.GroupPolicyIsSet -and $rdpGroupPolicySettings.GroupPolicyIsEnabled
    $rdpEnabledLocally = !($rdpGroupPolicySettings.GroupPolicyIsSet) -and (Test-RdpSystem)

    return (Test-RdpListener) -and ($rdpEnabledWithGP -or $rdpEnabledLocally)
}

<#

.SYNOPSIS
Tests whether or not a Remote Desktop Protocol listener exists.

.DESCRIPTION
Tests whether or not a Remote Desktop Protocol listener exists.

.ROLE
Readers

#>
function Test-RdpListener {
    $listeners = Get-RdpListener
    return $listeners.Count -gt 0
}

<#

.SYNOPSIS
Tests whether Remote Desktop Protocol is enabled via local system settings.

.DESCRIPTION
Tests whether Remote Desktop Protocol is enabled via local system settings.

.ROLE
Readers

#>
function Test-RdpSystem {
    $registryKey = Get-ItemProperty -Path $RdpSystemRegistryKey -ErrorAction SilentlyContinue
    return $registryKey.fDenyTSConnections -eq 0
}

<#

.SYNOPSIS
Tests whether Remote Desktop connections require Network Level Authentication while enabled via local system settings.

.DESCRIPTION
Tests whether Remote Desktop connections require Network Level Authentication while enabled via local system settings.

.ROLE
Readers

#>
function Test-RdpSystemUserAuthentication {
    $listener = Get-RdpListener | Where-Object { $_.UserAuthentication -ne $null } | Microsoft.PowerShell.Utility\Select-Object -First 1
    return $listener.UserAuthentication -eq 1
}

<#

.SYNOPSIS
Tests whether Remote Desktop connections require Network Level Authentication.

.DESCRIPTION
Tests whether Remote Desktop connections require Network Level Authentication.

.ROLE
Readers

#>
function Test-RdpUserAuthentication {
    $nlaEnabledWithGP = $false
    $nlaEnabledLocally = $false
    $nlaGroupPolicySettings = Get-RdpNlaGroupPolicySettings
    $nlaEnabledWithGP = $nlaGroupPolicySettings.GroupPolicyIsSet -and $nlaGroupPolicySettings.GroupPolicyIsEnabled
    $nlaEnabledLocally = !($nlaGroupPolicySettings.GroupPolicyIsSet) -and (Test-RdpSystemUserAuthentication)

    return $nlaEnabledWithGP -or $nlaEnabledLocally
}

#########
# Main
#########

Get-RdpSettings

}
## [END] Get-WACRDRemoteDesktopSettings ##
function Set-WACRDRemoteApplicationSettings {
<#

.SYNOPSIS
Sets the remote application allow list setting on target node.

.DESCRIPTION
Sets the remote application allow list setting on target node. This value is used
to determine whether the Remote App tool is available for usage.

.ROLE
Administrators

#>

param(
     [Parameter(Mandatory=$true)]
    [int]
    $allowRemoteApplication)

$key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server\TSAppAllowList'
Set-ItemProperty -Path $key -Name 'fDisabledAllowList' -Value $allowRemoteApplication


}
## [END] Set-WACRDRemoteApplicationSettings ##
function Set-WACRDRemoteDesktopSettings {
<#

.SYNOPSIS
Sets the remote desktop settings on target node.

.DESCRIPTION
Sets the remote desktop settings on target node. The value stored is used to
determine whether Remote Desktop and Remote App capabilities will be avaiable.

.ROLE
Administrators

#>

param(
    [Parameter(Mandatory=$False)]
    [boolean]
    $AllowRemoteDesktop,

    [Parameter(Mandatory=$False)]
    [boolean]
    $AllowRemoteDesktopWithNLA,

    [Parameter(Mandatory=$False)]
    [int]
    $PortNumber)

    Set-Variable -Option Constant -Name RdpFirewallGroup -Value "@FirewallAPI.dll,-28752" -ErrorAction SilentlyContinue

    $regKey1 = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'
    $regKey2 = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'

    $keyProperty1 = "fDenyTSConnections"
    $keyProperty2 = "UserAuthentication"
    $keyProperty3 = "PortNumber"

    $keyPropertyValue1 = $(if ($AllowRemoteDesktop -eq $True) { 0 } else { 1 })
    $keyPropertyValue2 = $(if ($AllowRemoteDesktopWithNLA -eq $True) { 1 } else { 0 })

    if (!(Test-Path $regKey1))
    {
        New-Item -Path $regKey1 -Force | Out-Null
    }
    New-ItemProperty -Path $regKey1 -Name $keyProperty1 -Value $keyPropertyValue1 -PropertyType DWORD -Force | Out-Null

    if (!(Test-Path $regKey2))
    {
        New-Item -Path $regKey2 -Force | Out-Null
    }
    New-ItemProperty -Path $regKey2 -Name $keyProperty2 -Value $keyPropertyValue2 -PropertyType DWORD -Force | Out-Null

    if ($keyPropertyValue3) {
        New-ItemProperty -Path $regKey2 -Name $keyProperty3 -Value $PortNumber -PropertyType DWORD -Force | Out-Null
    }

    Enable-NetFirewallRule -Group $RdpFirewallGroup

}
## [END] Set-WACRDRemoteDesktopSettings ##
function Add-WACRDFolderShare {
<#

.SYNOPSIS
Gets a new share name for the folder.

.DESCRIPTION
Gets a new share name for the folder. It starts with the folder name. Then it keeps appending "2" to the name
until the name is free. Finally return the name.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to the folder to be shared.

.PARAMETER Name
    String -- The suggested name to be shared (the folder name).

.PARAMETER Force
    boolean -- override any confirmations

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,    

    [Parameter(Mandatory = $true)]
    [String]
    $Name
)

Set-StrictMode -Version 5.0

while([bool](Get-SMBShare -Name $Name -ea 0)){
    $Name = $Name + '2';
}

New-SmbShare -Name "$Name" -Path "$Path"
@{ shareName = $Name }

}
## [END] Add-WACRDFolderShare ##
function Add-WACRDFolderShareNameUser {
<#

.SYNOPSIS
Adds a user to the folder share.

.DESCRIPTION
Adds a user to the folder share.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Name
    String -- Name of the share.

.PARAMETER AccountName
    String -- The user identification (AD / Local user).

.PARAMETER AccessRight
    String -- Access rights of the user.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Name,

    [Parameter(Mandatory = $true)]
    [String]
    $AccountName,

    [Parameter(Mandatory = $true)]
    [String]
    $AccessRight
)

Set-StrictMode -Version 5.0

Grant-SmbShareAccess -Name "$Name" -AccountName "$AccountName" -AccessRight "$AccessRight" -Force


}
## [END] Add-WACRDFolderShareNameUser ##
function Add-WACRDFolderShareUser {
<#

.SYNOPSIS
Adds a user access to the folder.

.DESCRIPTION
Adds a user access to the folder.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to the folder.

.PARAMETER Identity
    String -- The user identification (AD / Local user).

.PARAMETER FileSystemRights
    String -- File system rights of the user.

.PARAMETER AccessControlType
    String -- Access control type of the user.    

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $Identity,

    [Parameter(Mandatory = $true)]
    [String]
    $FileSystemRights,

    [ValidateSet('Deny','Allow')]
    [Parameter(Mandatory = $true)]
    [String]
    $AccessControlType
)

Set-StrictMode -Version 5.0

function Remove-UserPermission
{
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Path,
    
        [Parameter(Mandatory = $true)]
        [String]
        $Identity,
        
        [ValidateSet('Deny','Allow')]
        [Parameter(Mandatory = $true)]
        [String]
        $ACT
    )

    $Acl = Get-Acl $Path
    $AccessRule = New-Object system.security.accesscontrol.filesystemaccessrule($Identity, 'ReadAndExecute','ContainerInherit, ObjectInherit', 'None', $ACT)
    $Acl.RemoveAccessRuleAll($AccessRule)
    Set-Acl $Path $Acl
}

If ($AccessControlType -eq 'Deny') {
    $FileSystemRights = 'FullControl'
    Remove-UserPermission $Path $Identity 'Allow'
} else {
    Remove-UserPermission $Path $Identity 'Deny'
}

$Acl = Get-Acl $Path
$AccessRule = New-Object system.security.accesscontrol.filesystemaccessrule($Identity, $FileSystemRights,'ContainerInherit, ObjectInherit', 'None', $AccessControlType)
$Acl.AddAccessRule($AccessRule)
Set-Acl $Path $Acl

}
## [END] Add-WACRDFolderShareUser ##
function Compress-WACRDArchiveFileSystemEntity {
<#

.SYNOPSIS
Compresses the specified file system entity (files, folders) of the system.

.DESCRIPTION
Compresses the specified file system entity (files, folders) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER pathSource
    String -- The path to compress.

.PARAMETER PathDestination
    String -- The destination path to compress into.

.PARAMETER Force
    boolean -- override any confirmations

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $PathSource,    

    [Parameter(Mandatory = $true)]
    [String]
    $PathDestination,

    [Parameter(Mandatory = $false)]
    [boolean]
    $Force
)

Set-StrictMode -Version 5.0

if ($Force) {
    Compress-Archive -Path $PathSource -Force -DestinationPath $PathDestination
} else {
    Compress-Archive -Path $PathSource -DestinationPath $PathDestination
}
if ($error) {
    $code = $error[0].Exception.HResult
    @{ status = "error"; code = $code; message = $error }
} else {
    @{ status = "ok"; }
}

}
## [END] Compress-WACRDArchiveFileSystemEntity ##
function Disable-WACRDKdcProxy {
<#
.SYNOPSIS
Disables kdc proxy on the server

.DESCRIPTION
Disables kdc proxy on the server

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [string]
    $KdcPort
)

$urlLeft = "https://+:"
$urlRight = "/KdcProxy/"
$url = $urlLeft + $KdcPort + $urlRight
$deleteOutput = netsh http delete urlacl url=$url
if ($LASTEXITCODE -ne 0) {
    throw $deleteOutput
}
Remove-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "HttpsClientAuth"
Remove-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "DisallowUnprotectedPasswordAuth"
Stop-Service -Name kpssvc
Set-Service -Name kpssvc -StartupType Disabled
$firewallString = "KDC Proxy Server service (KPS) for SMB over QUIC"
Remove-NetFirewallRule -DisplayName $firewallString -ErrorAction SilentlyContinue
}
## [END] Disable-WACRDKdcProxy ##
function Disable-WACRDSmbOverQuic {
<#

.SYNOPSIS
Disables smb over QUIC on the server.

.DESCRIPTION
Disables smb over QUIC on the server.

.ROLE
Administrators

#>

Set-SmbServerConfiguration -EnableSMBQUIC $false -Force
}
## [END] Disable-WACRDSmbOverQuic ##
function Edit-WACRDFolderShareInheritanceFlag {
<#

.SYNOPSIS
Modifies all users' IsInherited flag to false

.DESCRIPTION
Modifies all users' IsInherited flag to false
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to the folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

$Acl = Get-Acl $Path
$Acl.SetAccessRuleProtection($True, $True)
Set-Acl -Path $Path -AclObject $Acl

}
## [END] Edit-WACRDFolderShareInheritanceFlag ##
function Edit-WACRDFolderShareUser {
<#

.SYNOPSIS
Edits a user access to the folder.

.DESCRIPTION
Edits a user access to the folder.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to the folder.

.PARAMETER Identity
    String -- The user identification (AD / Local user).

.PARAMETER FileSystemRights
    String -- File system rights of the user.

.PARAMETER AccessControlType
    String -- Access control type of the user.    

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $Identity,

    [Parameter(Mandatory = $true)]
    [String]
    $FileSystemRights,

    [ValidateSet('Deny','Allow')]
    [Parameter(Mandatory = $true)]
    [String]
    $AccessControlType
)

Set-StrictMode -Version 5.0

function Remove-UserPermission
{
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Path,
    
        [Parameter(Mandatory = $true)]
        [String]
        $Identity,
        
        [ValidateSet('Deny','Allow')]
        [Parameter(Mandatory = $true)]
        [String]
        $ACT
    )

    $Acl = Get-Acl $Path
    $AccessRule = New-Object system.security.accesscontrol.filesystemaccessrule($Identity, 'ReadAndExecute','ContainerInherit, ObjectInherit', 'None', $ACT)
    $Acl.RemoveAccessRuleAll($AccessRule)
    Set-Acl $Path $Acl
}

If ($AccessControlType -eq 'Deny') {
    $FileSystemRights = 'FullControl'
    Remove-UserPermission $Path $Identity 'Allow'
} else {
    Remove-UserPermission $Path $Identity 'Deny'
}

$Acl = Get-Acl $Path
$AccessRule = New-Object system.security.accesscontrol.filesystemaccessrule($Identity, $FileSystemRights,'ContainerInherit, ObjectInherit', 'None', $AccessControlType)
$Acl.SetAccessRule($AccessRule)
Set-Acl $Path $Acl




}
## [END] Edit-WACRDFolderShareUser ##
function Edit-WACRDSmbFileShare {
<#

.SYNOPSIS
Edits the smb file share details on the server.

.DESCRIPTION
Edits the smb file share details on the server.

.ROLE
Administrators

#>
param (
    [Parameter(Mandatory = $true)]
    [String]
    $name,

    [Parameter(Mandatory = $false)]
    [String[]]
    $noAccess,

    [Parameter(Mandatory = $false)]
    [String[]]
    $fullAccess,

    [Parameter(Mandatory = $false)]
    [String[]]
    $changeAccess,

    [Parameter(Mandatory = $false)]
    [String[]]
    $readAccess,
    
    [Parameter(Mandatory = $false)]
    [String[]]
    $unblockAccess,

    [Parameter(Mandatory = $false)]
    [Int]
    $cachingMode,

    [Parameter(Mandatory = $false)]
    [boolean]
    $encryptData,

    # TODO: 
    # [Parameter(Mandatory = $false)]
    # [Int]
    # $folderEnumerationMode

    [Parameter(Mandatory = $false)]
    [boolean]
    $compressData,

    [Parameter(Mandatory = $false)]
    [boolean]
    $isCompressDataEnabled
)

if($fullAccess.count -gt 0){
    Grant-SmbShareAccess -Name "$name" -AccountName $fullAccess -AccessRight Full -SmbInstance Default -Force
}
if($changeAccess.count -gt 0){
    Grant-SmbShareAccess -Name "$name" -AccountName $changeAccess -AccessRight Change -SmbInstance Default -Force
}
if($readAccess.count -gt 0){
    Grant-SmbShareAccess -Name "$name" -AccountName $readAccess -AccessRight Read -SmbInstance Default -Force
}
if($noAccess.count -gt 0){
    Revoke-SmbShareAccess -Name "$name" -AccountName $noAccess -SmbInstance Default  -Force
    Block-SmbShareAccess -Name "$name" -AccountName $noAccess -SmbInstance Default -Force
}
if($unblockAccess.count -gt 0){
    Unblock-SmbShareAccess -Name "$name" -AccountName $unblockAccess -SmbInstance Default  -Force
}
if($isCompressDataEnabled){
    Set-SmbShare -Name "$name" -CompressData $compressData -Force
}

Set-SmbShare -Name "$name" -CachingMode "$cachingMode" -EncryptData  $encryptData -Force




}
## [END] Edit-WACRDSmbFileShare ##
function Edit-WACRDSmbServerCertificateMapping {
<#
.SYNOPSIS
Edit SMB Server Certificate Mapping

.DESCRIPTION
Edits smb over QUIC certificate.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER thumbprint
    String -- The thumbprint of the certifiacte selected.

.PARAMETER newSelectedDnsNames
    String[] -- The addresses newly added to the certificate mapping.

.PARAMETER unSelectedDnsNames
    String[] -- To addresses to be removed from the certificate mapping.

#>
param (
    [Parameter(Mandatory = $true)]
    [String]
    $Thumbprint,

    [Parameter(Mandatory = $false)]
    [String[]]
    $NewSelectedDnsNames,

    [Parameter(Mandatory = $false)]
    [String[]]
    $UnSelectedDnsNames,

    [Parameter(Mandatory = $true)]
    [boolean]
    $IsKdcProxyEnabled,

    [Parameter(Mandatory = $true)]
    [String]
    $KdcProxyOptionSelected,

    [Parameter(Mandatory = $true)]
    [boolean]
    $IsKdcProxyMappedForSmbOverQuic,

    [Parameter(Mandatory = $false)]
    [String]
    $KdcPort,

    [Parameter(Mandatory = $false)]
    [String]
    $CurrentkdcPort,

    [Parameter(Mandatory = $false)]
    [boolean]
    $IsSameCertificate
)

Import-Module -Name Microsoft.PowerShell.Management -ErrorAction SilentlyContinue
Set-Variable -Name LogName -Option Constant -Value "Microsoft-ServerManagementExperience" -ErrorAction SilentlyContinue
Set-Variable -Name LogSource -Option Constant -Value "SmeScripts-ConfigureKdcProxy" -ErrorAction SilentlyContinue

function writeInfoLog($logMessage) {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
        -Message $logMessage -ErrorAction SilentlyContinue
}

Set-Location Cert:\LocalMachine\My

$port = "0.0.0.0:"
$urlLeft = "https://+:"
$urlRight = "/KdcProxy"

if ($UnSelectedDnsNames.count -gt 0) {
    foreach ($unSelectedDnsName in $UnSelectedDnsNames) {
        Remove-SmbServerCertificateMapping -Name $unSelectedDnsName -Force
    }
}

if ($NewSelectedDnsNames.count -gt 0) {
    foreach ($newSelectedDnsName in $NewSelectedDnsNames) {
        New-SmbServerCertificateMapping -Name $newSelectedDnsName -Thumbprint $Thumbprint -StoreName My -Force
    }
}

function Delete-KdcSSLCert([string]$deletePort) {
    $ipport = $port+$deletePort
    $deleteCertKdc = netsh http delete sslcert ipport=$ipport
    if ($LASTEXITCODE -ne 0) {
        throw $deleteCertKdc
    }
    $message = 'Completed deleting  ssl certificate port'
    writeInfoLog $message
    return;
}

function Enable-KdcProxy {

    $ipport = $port+$KdcPort
    $ComputerName = (Get-CimInstance Win32_ComputerSystem).DNSHostName + "." + (Get-CimInstance Win32_ComputerSystem).Domain
    try {
        if(!$IsSameCertificate -or ($KdcPort -ne $CurrentkdcPort) -or (!$IsKdcProxyMappedForSmbOverQuic) ) {
            $guid = [Guid]::NewGuid()
            $netshAddCertBinding = netsh http add sslcert ipport=$ipport certhash=$Thumbprint certstorename="my" appid="{$guid}"
            if ($LASTEXITCODE -ne 0) {
              throw $netshAddCertBinding
            }
            $message = 'Completed adding ssl certificate port'
            writeInfoLog $message
        }
        if ($NewSelectedDnsNames.count -gt 0) {
            foreach ($newSelectedDnsName in $NewSelectedDnsNames) {
                if($ComputerName.trim() -ne $newSelectedDnsName.trim()){
                    $output = Echo 'Y' | netdom computername $ComputerName /add $newSelectedDnsName
                    if ($LASTEXITCODE -ne 0) {
                        throw $output
                    }
                }
                $message = 'Completed adding alternate names for the computer'
                writeInfoLog $message
            }
        }
        if (!$IsKdcProxyEnabled) {
            $url = $urlLeft + $KdcPort + $urlRight
            $netshOutput = netsh http add urlacl url=$url user="NT authority\Network Service"
            if ($LASTEXITCODE -ne 0) {
                throw $netshOutput
            }
            $message = 'Completed adding urlacl'
            writeInfoLog $message
            New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings"  -force
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "HttpsClientAuth" -Value 0x0 -type DWORD
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "DisallowUnprotectedPasswordAuth" -Value 0x0 -type DWORD
            Set-Service -Name kpssvc -StartupType Automatic
            Start-Service -Name kpssvc
        }
        $message = 'Returning from call Enable-KdcProxy'
        writeInfoLog $message
        return $true;
    }
    catch {
        throw $_
    }
}

if($IsKdcProxyEnabled -and $KdcProxyOptionSelected -eq "enabled" -and ($KdcPort -ne $CurrentkdcPort)) {
    $url = $urlLeft + $CurrentkdcPort + $urlRight
    $deleteOutput = netsh http delete urlacl url=$url
    if ($LASTEXITCODE -ne 0) {
        throw $deleteOutput
    }
    $message = 'Completed deleting urlacl'
    writeInfoLog $message
    $newUrl = $urlLeft + $KdcPort + $urlRight
    $netshOutput = netsh http add urlacl url=$newUrl user="NT authority\Network Service"
    if ($LASTEXITCODE -ne 0) {
        throw $netshOutput
    }
    $message = 'Completed adding urlacl'
    writeInfoLog $message
}

if ($KdcProxyOptionSelected -eq "enabled" -and $KdcPort -ne $null) {
    if($IsKdcProxyMappedForSmbOverQuic -and (!$IsSameCertificate -or ($KdcPort -ne $CurrentkdcPort))) {
        Delete-KdcSSLCert $CurrentkdcPort
    }
    $result = Enable-KdcProxy
    if ($result) {
        $firewallString = "KDC Proxy Server service (KPS) for SMB over QUIC"
        $firewallDesc = "The KDC Proxy Server service runs on edge servers to proxy Kerberos protocol messages to domain controllers on the corporate network. Default port is TCP/443."
        New-NetFirewallRule -DisplayName $firewallString -Description $firewallDesc -Protocol TCP -LocalPort $KdcPort -Direction Inbound -Action Allow
    }
}

if ($IsKdcProxyMappedForSmbOverQuic -and $KdcProxyOptionSelected -ne "enabled" ) {
    Delete-KdcSSLCert $CurrentKdcPort
    $firewallString = "KDC Proxy Server service (KPS) for SMB over QUIC"
    Remove-NetFirewallRule -DisplayName $firewallString
}



}
## [END] Edit-WACRDSmbServerCertificateMapping ##
function Enable-WACRDSmbOverQuic {
<#

.SYNOPSIS
Disables smb over QUIC on the server.

.DESCRIPTION
Disables smb over QUIC on the server.

.ROLE
Administrators

#>

Set-SmbServerConfiguration -EnableSMBQUIC $true -Force
}
## [END] Enable-WACRDSmbOverQuic ##
function Expand-WACRDArchiveFileSystemEntity {
<#

.SYNOPSIS
Expands the specified file system entity (files, folders) of the system.

.DESCRIPTION
Expands the specified file system entity (files, folders) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER pathSource
    String -- The path to expand.

.PARAMETER PathDestination
    String -- The destination path to expand into.

.PARAMETER Force
    boolean -- override any confirmations

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $PathSource,    

    [Parameter(Mandatory = $true)]
    [String]
    $PathDestination,

    [Parameter(Mandatory = $false)]
    [boolean]
    $Force
)

Set-StrictMode -Version 5.0

if ($Force) {
    Expand-Archive -Path $PathSource -Force -DestinationPath $PathDestination
} else {
    Expand-Archive -Path $PathSource -DestinationPath $PathDestination
}

if ($error) {
    $code = $error[0].Exception.HResult
    @{ status = "error"; code = $code; message = $error }
} else {
    @{ status = "ok"; }
}

}
## [END] Expand-WACRDArchiveFileSystemEntity ##
function Get-WACRDBestHostNode {
<#

.SYNOPSIS
Returns the list of available cluster node names, and the best node name to host a new virtual machine.

.DESCRIPTION
Use the cluster CIM provider (MSCluster) to ask the cluster which node is the best to host a new virtual machine.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

Import-Module CimCmdlets -ErrorAction SilentlyContinue
Import-Module FailoverClusters -ErrorAction SilentlyContinue

Import-LocalizedData -BindingVariable strings -FileName strings.psd1 -ErrorAction SilentlyContinue


<#

.SYNOPSIS
Setup the script environment.

.DESCRIPTION
Setup the script environment.  Create read only (constant) variables
that add context to the said constants.

#>

function setupScriptEnv() {
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScripts" -Scope Script
    Set-Variable -Name clusterCimNameSpace -Option ReadOnly -Value "root/MSCluster" -Scope Script
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Get-BestHostNode.ps1" -Scope Script
    Set-Variable -Name BestNodePropertyName -Option ReadOnly -Value "BestNode" -Scope Script
    Set-Variable -Name StateUp -Option ReadOnly -Value "0" -Scope Script
}

<#

.SYNOPSIS
Cleanup the script environment.

.DESCRIPTION
Cleanup the script environment.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name clusterCimNameSpace -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name BestNodePropertyName -Scope Script -Force
    Remove-Variable -Name StateUp -Scope Script -Force
}
    
<#

.SYNOPSIS
Get the fully qualified domain name for the passed in server name from DNS.

.DESCRIPTION
Get the fully qualified domain name for the passed in server name from DNS.

#>

function GetServerFqdn([string]$netBIOSName) {
    try {
        $fqdn = [System.Net.DNS]::GetHostByName($netBIOSName).HostName

        return $fqdn.ToLower()
    } catch {
        $errMessage = $_.Exception.Message

        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]: There was an error looking up the FQDN for server $netBIOSName.  Error: $errMessage"  -ErrorAction SilentlyContinue

        return $netBIOSName
    }    
}

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell cmdlets installed on this server?

#>

function getIsClusterCmdletsAvailable() {
    $cmdlet = Get-Command "Get-Cluster" -ErrorAction SilentlyContinue

    return !!$cmdlet
}

<#

.SYNOPSIS
is the cluster CIM (WMI) provider installed on this server?

.DESCRIPTION
Returns true when the cluster CIM provider is installed on this server.

#>

function isClusterCimProviderAvailable() {
    $namespace = Get-CimInstance -Namespace $clusterCimNamespace -ClassName __NAMESPACE -ErrorAction SilentlyContinue

    return !!$namespace
}

<#

.SYNOPSIS
Get the MSCluster Cluster Service CIM instance from this server.

.DESCRIPTION
Get the MSCluster Cluster Service CIM instance from this server.

#>

function getClusterServiceCimInstance() {
    return Get-CimInstance -Namespace $clusterCimNamespace MSCluster_ClusterService -ErrorAction SilentlyContinue
}

<#

.SYNOPSIS
Get the list of the cluster nodes that are running.

.DESCRIPTION
Returns a list of cluster node names that are running using PowerShell.

#>

function getAllUpClusterNodeNames() {
    # Constants
    Set-Variable -Name stateUp -Option Readonly -Value "up" -Scope Local

    try {
        return Get-ClusterNode | Where-Object { $_.State -eq $stateUp } | ForEach-Object { (GetServerFqdn $_.Name) }
    } finally {
        Remove-Variable -Name stateUp -Scope Local -Force
    }
}

<#

.SYNOPSIS
Get the list of the cluster nodes that are running.

.DESCRIPTION
Returns a list of cluster node names that are running using CIM.

#>

function getAllUpClusterCimNodeNames() {
##SkipCheck=true##
    $query = "select name, state from MSCluster_Node Where state = '{0}'" -f $StateUp
##SkipCheck=false##
    return Get-CimInstance -Namespace $clusterCimNamespace -Query $query | ForEach-Object { (GetServerFqdn $_.Name) }
}

<#

.SYNOPSIS
Create a new instance of the "results" PS object.

.DESCRIPTION
Create a new PS object and set the passed in nodeNames to the appropriate property.

#>

function newResult([string []] $nodeNames) {
    $result = new-object PSObject
    $result | Add-Member -Type NoteProperty -Name Nodes -Value $nodeNames

    return $result;
}

<#

.SYNOPSIS
Remove any old lingering reservation for our typical VM.

.DESCRIPTION
Remove the reservation from the passed in id.

#>

function removeReservation($clusterService, [string] $rsvId) {
    Set-Variable removeReservationMethodName -Option Constant -Value "RemoveVmReservation"

    Invoke-CimMethod -CimInstance $clusterService -MethodName $removeReservationMethodName -Arguments @{ReservationId = $rsvId} -ErrorVariable +err | Out-Null    
}

<#

.SYNOPSIS
Create a reservation for our typical VM.

.DESCRIPTION
Create a reservation for the passed in id.

#>

function createReservation($clusterService, [string] $rsvId) {
    Set-Variable -Name createReservationMethodName -Option ReadOnly -Value "CreateVmReservation" -Scope Local
    Set-Variable -Name reserveSettings -Option ReadOnly -Value @{VmMemory = 2048; VmVirtualCoreCount = 2; VmCpuReservation = 0; VmFlags = 0; TimeSpan = 2000; ReservationId = $rsvId; LocalDiskSize = 0; Version = 0} -Scope Local

    try {
        $vmReserve = Invoke-CimMethod -CimInstance $clusterService -MethodName $createReservationMethodName -ErrorAction SilentlyContinue -ErrorVariable va -Arguments $reserveSettings

        if (!!$vmReserve -and $vmReserve.ReturnValue -eq 0 -and !!$vmReserve.NodeId) {
            return $vmReserve.NodeId
        }

        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]: Could not create a reservation for a virtual machine. Output from $createReservationMethodName is $vmReserve"  -ErrorAction SilentlyContinue

        return $null
    } finally {
        Remove-Variable -Name createReservationMethodName -Scope Local -Force
        Remove-Variable -Name reserveSettings -Scope Local -Force
    }
}

<#

.SYNOPSIS
Use the Cluster CIM provider to find the best host name for a typical VM.

.DESCRIPTION
Returns the best host node name, or null when none are found.

#>

function askClusterServiceForBestHostNode() {
    # API parameters
    Set-Variable -Name rsvId -Option ReadOnly -Value "TempVmId1" -Scope Local
    
    try {
        # If the class exist, using api to get optimal host
        $clusterService = getClusterServiceCimInstance
        if (!!$clusterService) {
            $nodeNames = @(getAllUpClusterCimNodeNames)
            $result = newResult $nodeNames
        
            # remove old reserveration if there is any
            removeReservation $clusterService $rsvId

            $id = createReservation $clusterService $rsvId

            if (!!$id) {
    ##SkipCheck=true##            
                $query = "select name, id from MSCluster_Node where id = '{0}'" -f $id
    ##SkipCheck=false##
                $bestNode = Get-CimInstance -Namespace $clusterCimNamespace -Query $query -ErrorAction SilentlyContinue

                if ($bestNode) {
                    $result | Add-Member -Type NoteProperty -Name $BestNodePropertyName -Value (GetServerFqdn $bestNode.Name)

                    return $result
                }
            } 
        }

        return $null
    } finally {
        Remove-Variable -Name rsvId -Scope Local -Force
    }
}

<#

.SYNOPSIS
Get the name of the cluster node that has the least number of VMs running on it.

.DESCRIPTION
Return the name of the cluster node that has the least number of VMs running on it.

#>

function getLeastLoadedNode() {
    # Constants
    Set-Variable -Name vmResourceTypeName -Option ReadOnly -Value "Virtual Machine" -Scope Local
    Set-Variable -Name OwnerNodePropertyName -Option ReadOnly -Value "OwnerNode" -Scope Local

    try {
        $nodeNames = @(getAllUpClusterNodeNames)
        $bestNodeName = $null;

        $result = newResult $nodeNames

        $virtualMachinesPerNode = @{}

        # initial counts as 0
        $nodeNames | ForEach-Object { $virtualMachinesPerNode[$_] = 0 }

        $ownerNodes = Get-ClusterResource | Where-Object { $_.ResourceType -eq $vmResourceTypeName } | Microsoft.PowerShell.Utility\Select-Object $OwnerNodePropertyName
        $ownerNodes | ForEach-Object { $virtualMachinesPerNode[$_.OwnerNode.Name]++ }

        # find node with minimum count
        $bestNodeName = $nodeNames[0]
        $min = $virtualMachinesPerNode[$bestNodeName]

        $nodeNames | ForEach-Object { 
            if ($virtualMachinesPerNode[$_] -lt $min) {
                $bestNodeName = $_
                $min = $virtualMachinesPerNode[$_]
            }
        }

        $result | Add-Member -Type NoteProperty -Name $BestNodePropertyName -Value (GetServerFqdn $bestNodeName)

        return $result
    } finally {
        Remove-Variable -Name vmResourceTypeName -Scope Local -Force
        Remove-Variable -Name OwnerNodePropertyName -Scope Local -Force
    }
}

<#

.SYNOPSIS
Main

.DESCRIPTION
Use the various mechanism available to determine the best host node.

#>

function main() {
    if (isClusterCimProviderAvailable) {
        $bestNode = askClusterServiceForBestHostNode
        if (!!$bestNode) {
            return $bestNode
        }
    }

    if (getIsClusterCmdletsAvailable) {
        return getLeastLoadedNode
    } else {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Warning `
            -Message "[$ScriptName]: The required PowerShell module (FailoverClusters) was not found."  -ErrorAction SilentlyContinue

        Write-Warning $strings.FailoverClustersModuleRequired
    }

    return $null    
}

###############################################################################
# Script execution begins here.
###############################################################################

if (-not ($env:pester)) {
    setupScriptEnv

    try {
        Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

        $result = main
        if (!!$result) {
            return $result
        }

        # If neither cluster CIM provider or PowerShell cmdlets are available then simply
        # return this computer's name as the best host node...
        $nodeName = GetServerFqdn $env:COMPUTERNAME

        $result = newResult @($nodeName)
        $result | Add-Member -Type NoteProperty -Name $BestNodePropertyName -Value $nodeName

        return $result
    } finally {
        cleanupScriptEnv
    }
}

}
## [END] Get-WACRDBestHostNode ##
function Get-WACRDCertificates {
<#

.SYNOPSIS
Get the certificates stored in my\store

.DESCRIPTION
Get the certificates stored in my\store

.ROLE
Readers

#>

$nearlyExpiredThresholdInDays = 60

$dnsNameList = @{}

<#
.Synopsis
    Name: Compute-ExpirationStatus
    Description: Computes expiration status based on notAfter date.
.Parameters
    $notAfter: A date object refering to certificate expiry date.

.Returns
    Enum values "Expired", "NearlyExpired" and "Healthy"
#>
function Compute-ExpirationStatus {
    param (
        [Parameter(Mandatory = $true)]
        [DateTime]$notAfter
    )

    if ([DateTime]::Now -gt $notAfter) {
        $expirationStatus = "Expired"
    }
    else {
        $nearlyExpired = [DateTime]::Now.AddDays($nearlyExpiredThresholdInDays);

        if ($nearlyExpired -ge $notAfter) {
            $expirationStatus = "NearlyExpired"
        }
        else {
            $expirationStatus = "Healthy"
        }
    }

    $expirationStatus
}

<# main - script starts here #>

Set-Location Cert:\LocalMachine\My

$certificates = Get-ChildItem -Recurse | Microsoft.PowerShell.Utility\Select-Object Subject, FriendlyName, NotBefore, NotAfter,
 Thumbprint, Issuer, @{n="DnsNameList";e={$_.DnsNameList}}, @{n="SignatureAlgorithm";e={$_.SignatureAlgorithm.FriendlyName}} |
ForEach-Object {
    return @{
        CertificateName = $_.Subject;
        FriendlyName = $_.FriendlyName;
        NotBefore = $_.NotBefore;
        NotAfter = $_.NotAfter;
        Thumbprint = $_.Thumbprint;
        Issuer = $_.Issuer;
        DnsNameList = $_.DnsNameList;
        Status = $(Compute-ExpirationStatus $_.NotAfter);
        SignatureAlgorithm  = $_.SignatureAlgorithm;
    }
}

return $certificates;

}
## [END] Get-WACRDCertificates ##
function Get-WACRDCimWin32LogicalDisk {
<#

.SYNOPSIS
Gets Win32_LogicalDisk object.

.DESCRIPTION
Gets Win32_LogicalDisk object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_LogicalDisk

}
## [END] Get-WACRDCimWin32LogicalDisk ##
function Get-WACRDCimWin32NetworkAdapter {
<#

.SYNOPSIS
Gets Win32_NetworkAdapter object.

.DESCRIPTION
Gets Win32_NetworkAdapter object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_NetworkAdapter

}
## [END] Get-WACRDCimWin32NetworkAdapter ##
function Get-WACRDCimWin32PhysicalMemory {
<#

.SYNOPSIS
Gets Win32_PhysicalMemory object.

.DESCRIPTION
Gets Win32_PhysicalMemory object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_PhysicalMemory

}
## [END] Get-WACRDCimWin32PhysicalMemory ##
function Get-WACRDCimWin32Processor {
<#

.SYNOPSIS
Gets Win32_Processor object.

.DESCRIPTION
Gets Win32_Processor object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_Processor

}
## [END] Get-WACRDCimWin32Processor ##
function Get-WACRDClusterInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a cluster.

.DESCRIPTION
Retrieves the inventory data for a cluster.

.ROLE
Readers

#>

Import-Module CimCmdlets -ErrorAction SilentlyContinue

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
Import-Module FailoverClusters -ErrorAction SilentlyContinue

Import-Module Storage -ErrorAction SilentlyContinue
<#

.SYNOPSIS
Get the name of this computer.

.DESCRIPTION
Get the best available name for this computer.  The FQDN is preferred, but when not avaialble
the NetBIOS name will be used instead.

#>

function getComputerName() {
    $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object Name, DNSHostName

    if ($computerSystem) {
        $computerName = $computerSystem.DNSHostName

        if ($null -eq $computerName) {
            $computerName = $computerSystem.Name
        }

        return $computerName
    }

    return $null
}

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell cmdlets installed on this server?

#>

function getIsClusterCmdletAvailable() {
    $cmdlet = Get-Command "Get-Cluster" -ErrorAction SilentlyContinue

    return !!$cmdlet
}

<#

.SYNOPSIS
Get the MSCluster Cluster CIM instance from this server.

.DESCRIPTION
Get the MSCluster Cluster CIM instance from this server.

#>
function getClusterCimInstance() {
    $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue

    if ($namespace) {
        return Get-CimInstance -Namespace root/mscluster MSCluster_Cluster -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object fqdn, S2DEnabled
    }

    return $null
}


<#

.SYNOPSIS
Determines if the current cluster supports Failover Clusters Time Series Database.

.DESCRIPTION
Use the existance of the path value of cmdlet Get-StorageHealthSetting to determine if TSDB
is supported or not.

#>
function getClusterPerformanceHistoryPath() {
    $storageSubsystem = Get-StorageSubSystem clus* -ErrorAction SilentlyContinue
    $storageHealthSettings = Get-StorageHealthSetting -InputObject $storageSubsystem -Name "System.PerformanceHistory.Path" -ErrorAction SilentlyContinue

    return $null -ne $storageHealthSettings
}

<#

.SYNOPSIS
Get some basic information about the cluster from the cluster.

.DESCRIPTION
Get the needed cluster properties from the cluster.

#>
function getClusterInfo() {
    $returnValues = @{}

    $returnValues.Fqdn = $null
    $returnValues.isS2DEnabled = $false
    $returnValues.isTsdbEnabled = $false

    $cluster = getClusterCimInstance
    if ($cluster) {
        $returnValues.Fqdn = $cluster.fqdn
        $isS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -eq 1)
        $returnValues.isS2DEnabled = $isS2dEnabled

        if ($isS2DEnabled) {
            $returnValues.isTsdbEnabled = getClusterPerformanceHistoryPath
        } else {
            $returnValues.isTsdbEnabled = $false
        }
    }

    return $returnValues
}

<#

.SYNOPSIS
Are the cluster PowerShell Health cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell Health cmdlets installed on this server?

s#>
function getisClusterHealthCmdletAvailable() {
    $cmdlet = Get-Command -Name "Get-HealthFault" -ErrorAction SilentlyContinue

    return !!$cmdlet
}
<#

.SYNOPSIS
Are the Britannica (sddc management resources) available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) available on the cluster?

#>
function getIsBritannicaEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual machine available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual machine available on the cluster?

#>
function getIsBritannicaVirtualMachineEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualMachine -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual switch available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual switch available on the cluster?

#>
function getIsBritannicaVirtualSwitchEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualSwitch -ErrorAction SilentlyContinue)
}

###########################################################################
# main()
###########################################################################

$clusterInfo = getClusterInfo

$result = New-Object PSObject

$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $clusterInfo.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2DEnabled' -Value $clusterInfo.isS2DEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsTsdbEnabled' -Value $clusterInfo.isTsdbEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterHealthCmdletAvailable' -Value (getIsClusterHealthCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value (getIsBritannicaEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualMachineEnabled' -Value (getIsBritannicaVirtualMachineEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualSwitchEnabled' -Value (getIsBritannicaVirtualSwitchEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterCmdletAvailable' -Value (getIsClusterCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'CurrentClusterNode' -Value (getComputerName)

$result

}
## [END] Get-WACRDClusterInventory ##
function Get-WACRDClusterNodes {
<#

.SYNOPSIS
Retrieves the inventory data for cluster nodes in a particular cluster.

.DESCRIPTION
Retrieves the inventory data for cluster nodes in a particular cluster.

.ROLE
Readers

#>

import-module CimCmdlets

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
import-module FailoverClusters -ErrorAction SilentlyContinue

###############################################################################
# Constants
###############################################################################

Set-Variable -Name LogName -Option Constant -Value "Microsoft-ServerManagementExperience" -ErrorAction SilentlyContinue
Set-Variable -Name LogSource -Option Constant -Value "SMEScripts" -ErrorAction SilentlyContinue
Set-Variable -Name ScriptName -Option Constant -Value $MyInvocation.ScriptName -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed?

.DESCRIPTION
Use the Get-Command cmdlet to quickly test if the cluster PowerShell cmdlets
are installed on this server.

#>

function getClusterPowerShellSupport() {
    $cmdletInfo = Get-Command 'Get-ClusterNode' -ErrorAction SilentlyContinue

    return $cmdletInfo -and $cmdletInfo.Name -eq "Get-ClusterNode"
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster CIM provider.

.DESCRIPTION
When the cluster PowerShell cmdlets are not available fallback to using
the cluster CIM provider to get the needed information.

#>

function getClusterNodeCimInstances() {
    # Change the WMI property NodeDrainStatus to DrainStatus to match the PS cmdlet output.
    return Get-CimInstance -Namespace root/mscluster MSCluster_Node -ErrorAction SilentlyContinue | `
        Microsoft.PowerShell.Utility\Select-Object @{Name="DrainStatus"; Expression={$_.NodeDrainStatus}}, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster PowerShell cmdlets.

.DESCRIPTION
When the cluster PowerShell cmdlets are available use this preferred function.

#>

function getClusterNodePsInstances() {
    return Get-ClusterNode -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object DrainStatus, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Use DNS services to get the FQDN of the cluster NetBIOS name.

.DESCRIPTION
Use DNS services to get the FQDN of the cluster NetBIOS name.

.Notes
It is encouraged that the caller add their approprate -ErrorAction when
calling this function.

#>

function getClusterNodeFqdn([string]$clusterNodeName) {
    return ([System.Net.Dns]::GetHostEntry($clusterNodeName)).HostName
}

<#

.SYNOPSIS
Writes message to event log as warning.

.DESCRIPTION
Writes message to event log as warning.

#>

function writeToEventLog([string]$message) {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Warning `
        -Message $message  -ErrorAction SilentlyContinue
}

<#

.SYNOPSIS
Get the cluster nodes.

.DESCRIPTION
When the cluster PowerShell cmdlets are available get the information about the cluster nodes
using PowerShell.  When the cmdlets are not available use the Cluster CIM provider.

#>

function getClusterNodes() {
    $isClusterCmdletAvailable = getClusterPowerShellSupport

    if ($isClusterCmdletAvailable) {
        $clusterNodes = getClusterNodePsInstances
    } else {
        $clusterNodes = getClusterNodeCimInstances
    }

    $clusterNodeMap = @{}

    foreach ($clusterNode in $clusterNodes) {
        $clusterNodeName = $clusterNode.Name.ToLower()
        try 
        {
            $clusterNodeFqdn = getClusterNodeFqdn $clusterNodeName -ErrorAction SilentlyContinue
        }
        catch 
        {
            $clusterNodeFqdn = $clusterNodeName
            writeToEventLog "[$ScriptName]: The fqdn for node '$clusterNodeName' could not be obtained. Defaulting to machine name '$clusterNodeName'"
        }

        $clusterNodeResult = New-Object PSObject

        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FullyQualifiedDomainName' -Value $clusterNodeFqdn
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'Name' -Value $clusterNodeName
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DynamicWeight' -Value $clusterNode.DynamicWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'NodeWeight' -Value $clusterNode.NodeWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FaultDomain' -Value $clusterNode.FaultDomain
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'State' -Value $clusterNode.State
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DrainStatus' -Value $clusterNode.DrainStatus

        $clusterNodeMap.Add($clusterNodeName, $clusterNodeResult)
    }

    return $clusterNodeMap
}

###########################################################################
# main()
###########################################################################

getClusterNodes

}
## [END] Get-WACRDClusterNodes ##
function Get-WACRDComputerName {
<#

.SYNOPSIS
Gets the computer name.

.DESCRIPTION
Gets the compuiter name.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

$ComputerName = $env:COMPUTERNAME
@{ computerName = $ComputerName }

}
## [END] Get-WACRDComputerName ##
function Get-WACRDDecryptedDataFromNode {
<#

.SYNOPSIS
Gets data after decrypting it on a node.

.DESCRIPTION
Decrypts data on node using a cached RSAProvider used during encryption within 3 minutes of encryption and returns the decrypted data.
This script should be imported or copied directly to other scripts, do not send the returned data as an argument to other scripts.

.PARAMETER encryptedData
Encrypted data to be decrypted (String).

.ROLE
Readers

#>
param (
  [Parameter(Mandatory = $true)]
  [String]
  $encryptedData
)

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function DecryptDataWithJWKOnNode {
  if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue) {
    $rsaProvider = (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedData), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
  }
  # If you copy this script directly to another, you can get rid of the throw statement and add custom error handling logic such as "Write-Error"
  throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
}

}
## [END] Get-WACRDDecryptedDataFromNode ##
function Get-WACRDEncryptionJWKOnNode {
<#

.SYNOPSIS
Gets encrytion JSON web key from node.

.DESCRIPTION
Gets encrytion JSON web key from node.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function Get-RSAProvider
{
    if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue)
    {
        return (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    }

    $Global:RSA = New-Object System.Security.Cryptography.RSACryptoServiceProvider -ArgumentList 4096
    return $RSA
}

function Get-JsonWebKey
{
    $rsaProvider = Get-RSAProvider
    $parameters = $rsaProvider.ExportParameters($false)
    return [PSCustomObject]@{
        kty = 'RSA'
        alg = 'RSA-OAEP'
        e = [Convert]::ToBase64String($parameters.Exponent)
        n = [Convert]::ToBase64String($parameters.Modulus).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }
}

$jwk = Get-JsonWebKey
ConvertTo-Json $jwk -Compress

}
## [END] Get-WACRDEncryptionJWKOnNode ##
function Get-WACRDFileNamesInPath {
<#

.SYNOPSIS
Enumerates all of the file system entities (files, folders, volumes) of the system.

.DESCRIPTION
Enumerates all of the file system entities (files, folders, volumes) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- The path to enumerate.

.PARAMETER OnlyFolders
    switch -- 

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $false)]
    [switch]
    $OnlyFolders
)

Set-StrictMode -Version 5.0

function isFolder($item) {
    return $item.Attributes -match "Directory"
}

function getName($item) {
    $slash = '';

    if (isFolder $item) {
        $slash = '\';
    }

    return "$($_.Name)$slash"
}

if ($onlyFolders) {
    return (Get-ChildItem -Path $Path | Where-Object {isFolder $_}) | ForEach-Object { return "$($_.Name)\"} | Sort-Object
}

return (Get-ChildItem -Path $Path) | ForEach-Object { return getName($_)} | Sort-Object

}
## [END] Get-WACRDFileNamesInPath ##
function Get-WACRDFileSystemEntities {
<#

.SYNOPSIS
Enumerates all of the file system entities (files, folders, volumes) of the system.

.DESCRIPTION
Enumerates all of the file system entities (files, folders, volumes) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- The path to enumerate.

.PARAMETER OnlyFiles
    switch --

.PARAMETER OnlyFolders
    switch --

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $false)]
    [Switch]
    $OnlyFiles,

    [Parameter(Mandatory = $false)]
    [Switch]
    $OnlyFolders
)

Set-StrictMode -Version 5.0

<#
.Synopsis
    Name: Get-FileSystemEntities
    Description: Gets all the local file system entities of the machine.

.Parameter Path
    String -- The path to enumerate.

.Returns
    The local file system entities.
#>
function Get-FileSystemEntities {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Path
    )

    $folderShares = Get-CimInstance -Class Win32_Share;

    if ($Path -match '\[' -or $Path -match '\]') {
        return Get-ChildItem -LiteralPath $Path -Force |
        Microsoft.PowerShell.Utility\Select-Object @{Name = "Caption"; Expression = { $_.FullName } },
        @{Name = "CreationDate"; Expression = { $_.CreationTimeUtc } },
        Extension,
        @{Name = "IsHidden"; Expression = { $_.Attributes -match "Hidden" } },
        @{Name = "IsShared"; Expression = { [bool]($folderShares | Where-Object Path -eq $_.FullName) } },
        Name,
        @{Name = "Type"; Expression = { Get-FileSystemEntityType -Attributes $_.Attributes } },
        @{Name = "LastModifiedDate"; Expression = { $_.LastWriteTimeUtc } },
        @{Name = "Size"; Expression = { if ($_.PSIsContainer) { $null } else { $_.Length } } };
    }


    return Get-ChildItem -Path $Path -Force |
    Microsoft.PowerShell.Utility\Select-Object @{Name = "Caption"; Expression = { $_.FullName } },
    @{Name = "CreationDate"; Expression = { $_.CreationTimeUtc } },
    Extension,
    @{Name = "IsHidden"; Expression = { $_.Attributes -match "Hidden" } },
    @{Name = "IsShared"; Expression = { [bool]($folderShares | Where-Object Path -eq $_.FullName) } },
    Name,
    @{Name = "Type"; Expression = { Get-FileSystemEntityType -Attributes $_.Attributes } },
    @{Name = "LastModifiedDate"; Expression = { $_.LastWriteTimeUtc } },
    @{Name = "Size"; Expression = { if ($_.PSIsContainer) { $null } else { $_.Length } } };
}

<#
.Synopsis
    Name: Get-FileSystemEntityType
    Description: Gets the type of a local file system entity.

.Parameter Attributes
    The System.IO.FileAttributes of the FileSystemEntity.

.Returns
    The type of the local file system entity.
#>
function Get-FileSystemEntityType {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileAttributes]
        $Attributes
    )

    if ($Attributes -match "Directory") {
        return "Folder";
    }
    else {
        return "File";
    }
}

$entities = Get-FileSystemEntities -Path $Path;
if ($OnlyFiles -and $OnlyFolders) {
    return $entities;
}

if ($OnlyFiles) {
    return $entities | Where-Object { $_.Type -eq "File" };
}

if ($OnlyFolders) {
    return $entities | Where-Object { $_.Type -eq "Folder" };
}

return $entities;

}
## [END] Get-WACRDFileSystemEntities ##
function Get-WACRDFileSystemRoot {
<#

.SYNOPSIS
Enumerates the root of the file system (volumes and related entities) of the system.

.DESCRIPTION
Enumerates the root of the file system (volumes and related entities) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
import-module CimCmdlets

<#
.Synopsis
    Name: Get-FileSystemRoot
    Description: Gets the local file system root entities of the machine.

.Returns
    The local file system root entities.
#>
function Get-FileSystemRoot
{
    $volumes = Enumerate-Volumes;

    return $volumes |
        Microsoft.PowerShell.Utility\Select-Object @{Name="Caption"; Expression={$_.DriveLetter +":\"}},
                      @{Name="CreationDate"; Expression={$null}},
                      @{Name="Extension"; Expression={$null}},
                      @{Name="IsHidden"; Expression={$false}},
                      @{Name="Name"; Expression={if ($_.FileSystemLabel) { $_.FileSystemLabel + " (" + $_.DriveLetter + ":)"} else { "(" + $_.DriveLetter + ":)" }}},
                      @{Name="Type"; Expression={"Volume"}},
                      @{Name="LastModifiedDate"; Expression={$null}},
                      @{Name="Size"; Expression={$_.Size}},
                      @{Name="SizeRemaining"; Expression={$_.SizeRemaining}}
}

<#
.Synopsis
    Name: Get-Volumes
    Description: Gets the local volumes of the machine.

.Returns
    The local volumes.
#>
function Enumerate-Volumes
{
    Remove-Module Storage -ErrorAction Ignore; # Remove the Storage module to prevent it from automatically localizing

    $isDownlevel = [Environment]::OSVersion.Version.Major -lt 10;
    if ($isDownlevel)
    {
        $disks = Get-CimInstance -ClassName MSFT_Disk -Namespace root/Microsoft/Windows/Storage | Where-Object { !$_.IsClustered };
        $partitions = @($disks | Get-CimAssociatedInstance -ResultClassName MSFT_Partition)
        if ($partitions.Length -eq 0) {
            $volumes = Get-CimInstance -ClassName MSFT_Volume -Namespace root/Microsoft/Windows/Storage;
        } else {
            $volumes = $partitions | Get-CimAssociatedInstance -ResultClassName MSFT_Volume;
        }
    }
    else
    {
        $subsystem = Get-CimInstance -ClassName MSFT_StorageSubSystem -Namespace root/Microsoft/Windows/Storage| Where-Object { $_.FriendlyName -like "*Win*" };
        $volumes = $subsystem | Get-CimAssociatedInstance -ResultClassName MSFT_Volume;
    }

    return $volumes | Where-Object {
        try {
            [byte]$_.DriveLetter -ne 0 -and $_.DriveLetter -ne $null -and $_.Size -gt 0
        } catch {
            $false
        }
    };
}

Get-FileSystemRoot;

}
## [END] Get-WACRDFileSystemRoot ##
function Get-WACRDFolderItemCount {
<#

.SYNOPSIS
Gets the count of elements in the folder

.DESCRIPTION
Gets the count of elements in the folder
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- The path to the folder

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

$directoryInfo = Get-ChildItem $Path | Microsoft.PowerShell.Utility\Measure-Object
$directoryInfo.count

}
## [END] Get-WACRDFolderItemCount ##
function Get-WACRDFolderOwner {
<#

.SYNOPSIS
Gets the owner of a folder.

.DESCRIPTION
Gets the owner of a folder.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

$Owner = (Get-Acl $Path).Owner
@{ owner = $Owner; }

}
## [END] Get-WACRDFolderOwner ##
function Get-WACRDFolderShareNames {
<#

.SYNOPSIS
Gets the existing share names of a shared folder

.DESCRIPTION
Gets the existing share names of a shared folder
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- The path to the folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

Get-CimInstance -Class Win32_Share -Filter Path="'$Path'" | Microsoft.PowerShell.Utility\Select-Object Name

}
## [END] Get-WACRDFolderShareNames ##
function Get-WACRDFolderSharePath {
<#

.SYNOPSIS
Gets the existing share names of a shared folder

.DESCRIPTION
Gets the existing share names of a shared folder
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Name
    String -- The share name to the shared folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Name
)

Set-StrictMode -Version 5.0

Get-SmbShare -Includehidden | Where-Object { $_.Name -eq $Name } | Microsoft.PowerShell.Utility\Select-Object Path

}
## [END] Get-WACRDFolderSharePath ##
function Get-WACRDFolderShareStatus {
<#

.SYNOPSIS
Checks if a folder is shared

.DESCRIPTION
Checks if a folder is shared
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- the path to the folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

$Shared = [bool](Get-CimInstance -Class Win32_Share -Filter Path="'$Path'")
@{ isShared = $Shared }

}
## [END] Get-WACRDFolderShareStatus ##
function Get-WACRDFolderShareUsers {
<#

.SYNOPSIS
Gets the user access rights of a folder

.DESCRIPTION
Gets the user access rights of a folder
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- The path to the folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

Get-Acl $Path |  Microsoft.PowerShell.Utility\Select-Object -ExpandProperty Access | Microsoft.PowerShell.Utility\Select-Object IdentityReference, FileSystemRights, AccessControlType

}
## [END] Get-WACRDFolderShareUsers ##
function Get-WACRDIsAzureTurbineServer {
<#
.SYNOPSIS
Checks if the current server is Azure Turbine edition.

.DESCRIPTION
Returns true if the current server is Azure Turbine which supports smb over QUIC.

.ROLE
Readers

#>

Set-Variable -Name Server21H2SkuNumber -Value 407 -Scope Script
Set-Variable -Name Server21H2VersionNumber -Value 10.0.20348 -Scope Script
Set-Variable -Name Server2012R2VersionNumber -Value 6.3 -Scope Script

$result = Get-WmiObject -Class Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object OperatingSystemSKU, Version
$isAzureTurbineServer = $result.OperatingSystemSKU -eq $Server21H2SkuNumber -and [version]$result.version -ge [version]$Server21H2VersionNumber

$version = [System.Environment]::OSVersion.Version
$ver = $version.major.toString() + "." + $version.minor.toString()
$isWS2012R2orGreater = [version]$ver -ge [version]$Server2012R2VersionNumber

return @{ isAzureTurbineServer = $isAzureTurbineServer;
  isWS2012R2orGreater =  $isWS2012R2orGreater }

}
## [END] Get-WACRDIsAzureTurbineServer ##
function Get-WACRDItemProperties {
<#

.SYNOPSIS
Get item's properties.

.DESCRIPTION
Get item's properties on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- the path to the item whose properites are requested.

.PARAMETER ItemType
    String -- What kind of item?

#>

param (
    [Parameter(Mandatory = $true)]
    [String[]]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $ItemType
)

Set-StrictMode -Version 5.0

switch ($ItemType) {
    0 {
        Get-Volume $Path | Microsoft.PowerShell.Utility\Select-Object -Property *
    }
    default {
        Get-ItemProperty -Path $Path | Microsoft.PowerShell.Utility\Select-Object -Property *
    }
}

}
## [END] Get-WACRDItemProperties ##
function Get-WACRDItemType {
<#

.SYNOPSIS
Enumerates all of the file system entities (files, folders, volumes) of the system.

.DESCRIPTION
Enumerates all of the file system entities (files, folders, volumes) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- the path to the folder where enumeration should start.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

<#
.Synopsis
    Name: Get-FileSystemEntityType
    Description: Gets the type of a local file system entity.

.Parameter Attributes
    The System.IO.FileAttributes of the FileSystemEntity.

.Returns
    The type of the local file system entity.
#>
function Get-FileSystemEntityType
{
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileAttributes]
        $Attributes
    )

    if ($Attributes -match "Directory")
    {
        return "Folder";
    }
    else
    {
        return "File";
    }
}

if (Test-Path -LiteralPath $Path) {
    return Get-FileSystemEntityType -Attributes (Get-Item $Path -Force).Attributes
} else {
    return ''
}

}
## [END] Get-WACRDItemType ##
function Get-WACRDLocalGroups {
<#

.SYNOPSIS
Gets the local groups.

.DESCRIPTION
Gets the local groups. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Import-Module Microsoft.PowerShell.LocalAccounts -ErrorAction SilentlyContinue

$isWinServer2016OrNewer = [Environment]::OSVersion.Version.Major -ge 10;
# ADSI does NOT support 2016 Nano, meanwhile New-LocalUser, Get-LocalUser, Set-LocalUser do NOT support downlevel

    if ($isWinServer2016OrNewer)
    {
       return  Get-LocalGroup | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object Name
                                
    }
    else
    {
       return  Get-WmiObject -Class Win32_Group -Filter "LocalAccount='True'" | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object Name
    }


}
## [END] Get-WACRDLocalGroups ##
function Get-WACRDLocalUsers {
<#

.SYNOPSIS
Gets the local users.

.DESCRIPTION
Gets the local users. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>


$isWinServer2016OrNewer = [Environment]::OSVersion.Version.Major -ge 10;

if ($isWinServer2016OrNewer){

	return Get-LocalUser | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object Name

}
else{
    return Get-WmiObject -Class Win32_UserAccount -Filter "LocalAccount='True'" | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object  Name;
}
}
## [END] Get-WACRDLocalUsers ##
function Get-WACRDOSDetails {
<#

.SYNOPSIS
Get OS details

.DESCRIPTION
i) Get OS Version to detrmine if compression is supported on the server.
ii) Returns translated name of SID - S-1-1-0 "Everyone" on the system. This name can differ on OS's with languages that are not English.

.ROLE
Readers

#>

$result = Get-WmiObject -Class Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object Version

$isOSVersionCompatibleForCompression =[version]$result.version -ge [version]"10.0.20348"

# Get translated SID
$SID = New-Object System.Security.Principal.SecurityIdentifier("S-1-1-0")
$name = $SID.Translate([System.Security.Principal.NTAccount])
$translatedSidName = $name.value;

return @{
  isOSVersionCompatibleForCompression = $isOSVersionCompatibleForCompression;
  translatedSidName = $translatedSidName;
}


}
## [END] Get-WACRDOSDetails ##
function Get-WACRDServerInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a server.

.DESCRIPTION
Retrieves the inventory data for a server.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

Import-Module CimCmdlets

Import-Module Storage -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Converts an arbitrary version string into just 'Major.Minor'

.DESCRIPTION
To make OS version comparisons we only want to compare the major and
minor version.  Build number and/os CSD are not interesting.

#>

function convertOsVersion([string]$osVersion) {
  [Ref]$parsedVersion = $null
  if (![Version]::TryParse($osVersion, $parsedVersion)) {
    return $null
  }

  $version = [Version]$parsedVersion.Value
  return New-Object Version -ArgumentList $version.Major, $version.Minor
}

<#

.SYNOPSIS
Determines if CredSSP is enabled for the current server or client.

.DESCRIPTION
Check the registry value for the CredSSP enabled state.

#>

function isCredSSPEnabled() {
  Set-Variable credSSPServicePath -Option Constant -Value "WSMan:\localhost\Service\Auth\CredSSP"
  Set-Variable credSSPClientPath -Option Constant -Value "WSMan:\localhost\Client\Auth\CredSSP"

  $credSSPServerEnabled = $false;
  $credSSPClientEnabled = $false;

  $credSSPServerService = Get-Item $credSSPServicePath -ErrorAction SilentlyContinue
  if ($credSSPServerService) {
    $credSSPServerEnabled = [System.Convert]::ToBoolean($credSSPServerService.Value)
  }

  $credSSPClientService = Get-Item $credSSPClientPath -ErrorAction SilentlyContinue
  if ($credSSPClientService) {
    $credSSPClientEnabled = [System.Convert]::ToBoolean($credSSPClientService.Value)
  }

  return ($credSSPServerEnabled -or $credSSPClientEnabled)
}

<#

.SYNOPSIS
Determines if the Hyper-V role is installed for the current server or client.

.DESCRIPTION
The Hyper-V role is installed when the VMMS service is available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>

function isHyperVRoleInstalled() {
  $vmmsService = Get-Service -Name "VMMS" -ErrorAction SilentlyContinue

  return $vmmsService -and $vmmsService.Name -eq "VMMS"
}

<#

.SYNOPSIS
Determines if the Hyper-V PowerShell support module is installed for the current server or client.

.DESCRIPTION
The Hyper-V PowerShell support module is installed when the modules cmdlets are available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>
function isHyperVPowerShellSupportInstalled() {
  # quicker way to find the module existence. it doesn't load the module.
  return !!(Get-Module -ListAvailable Hyper-V -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Determines if Windows Management Framework (WMF) 5.0, or higher, is installed for the current server or client.

.DESCRIPTION
Windows Admin Center requires WMF 5 so check the registey for WMF version on Windows versions that are less than
Windows Server 2016.

#>
function isWMF5Installed([string] $operatingSystemVersion) {
  Set-Variable Server2016 -Option Constant -Value (New-Object Version '10.0')   # And Windows 10 client SKUs
  Set-Variable Server2012 -Option Constant -Value (New-Object Version '6.2')

  $version = convertOsVersion $operatingSystemVersion
  if (-not $version) {
    # Since the OS version string is not properly formatted we cannot know the true installed state.
    return $false
  }

  if ($version -ge $Server2016) {
    # It's okay to assume that 2016 and up comes with WMF 5 or higher installed
    return $true
  }
  else {
    if ($version -ge $Server2012) {
      # Windows 2012/2012R2 are supported as long as WMF 5 or higher is installed
      $registryKey = 'HKLM:\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine'
      $registryKeyValue = Get-ItemProperty -Path $registryKey -Name PowerShellVersion -ErrorAction SilentlyContinue

      if ($registryKeyValue -and ($registryKeyValue.PowerShellVersion.Length -ne 0)) {
        $installedWmfVersion = [Version]$registryKeyValue.PowerShellVersion

        if ($installedWmfVersion -ge [Version]'5.0') {
          return $true
        }
      }
    }
  }

  return $false
}

<#

.SYNOPSIS
Determines if the current usser is a system administrator of the current server or client.

.DESCRIPTION
Determines if the current usser is a system administrator of the current server or client.

#>
function isUserAnAdministrator() {
  return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

<#

.SYNOPSIS
Get some basic information about the Failover Cluster that is running on this server.

.DESCRIPTION
Create a basic inventory of the Failover Cluster that may be running in this server.

#>
function getClusterInformation() {
  $returnValues = @{ }

  $returnValues.IsS2dEnabled = $false
  $returnValues.IsCluster = $false
  $returnValues.ClusterFqdn = $null
  $returnValues.IsBritannicaEnabled = $false

  $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue
  if ($namespace) {
    $cluster = Get-CimInstance -Namespace root/MSCluster -ClassName MSCluster_Cluster -ErrorAction SilentlyContinue
    if ($cluster) {
      $returnValues.IsCluster = $true
      $returnValues.ClusterFqdn = $cluster.Fqdn
      $returnValues.IsS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -gt 0)
      $returnValues.IsBritannicaEnabled = $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
    }
  }

  return $returnValues
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

#>
function getComputerFqdnAndAddress($computerName) {
  $hostEntry = [System.Net.Dns]::GetHostEntry($computerName)
  $addressList = @()
  foreach ($item in $hostEntry.AddressList) {
    $address = New-Object PSObject
    $address | Add-Member -MemberType NoteProperty -Name 'IpAddress' -Value $item.ToString()
    $address | Add-Member -MemberType NoteProperty -Name 'AddressFamily' -Value $item.AddressFamily.ToString()
    $addressList += $address
  }

  $result = New-Object PSObject
  $result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $hostEntry.HostName
  $result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $addressList
  return $result
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

#>
function getHostFqdnAndAddress($computerSystem) {
  $computerName = $computerSystem.DNSHostName
  if (!$computerName) {
    $computerName = $computerSystem.Name
  }

  return getComputerFqdnAndAddress $computerName
}

<#

.SYNOPSIS
Are the needed management CIM interfaces available on the current server or client.

.DESCRIPTION
Check for the presence of the required server management CIM interfaces.

#>
function getManagementToolsSupportInformation() {
  $returnValues = @{ }

  $returnValues.ManagementToolsAvailable = $false
  $returnValues.ServerManagerAvailable = $false

  $namespaces = Get-CimInstance -Namespace root/microsoft/windows -ClassName __NAMESPACE -ErrorAction SilentlyContinue

  if ($namespaces) {
    $returnValues.ManagementToolsAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ManagementTools" })
    $returnValues.ServerManagerAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ServerManager" })
  }

  return $returnValues
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>
function isRemoteAppEnabled() {
  Set-Variable key -Option Constant -Value "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Terminal Server\\TSAppAllowList"

  $registryKeyValue = Get-ItemProperty -Path $key -Name fDisabledAllowList -ErrorAction SilentlyContinue

  if (-not $registryKeyValue) {
    return $false
  }
  return $registryKeyValue.fDisabledAllowList -eq 1
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>

<#
c
.SYNOPSIS
Get the Win32_OperatingSystem information as well as current version information from the registry

.DESCRIPTION
Get the Win32_OperatingSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller. Included in the results are current version
information from the registry

#>
function getOperatingSystemInfo() {
  $operatingSystemInfo = Get-CimInstance Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object csName, Caption, OperatingSystemSKU, Version, ProductType, OSType, LastBootUpTime, SerialNumber
  $currentVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" | Microsoft.PowerShell.Utility\Select-Object CurrentBuild, UBR, DisplayVersion

  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name CurrentBuild -Value $currentVersion.CurrentBuild
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name UpdateBuildRevision -Value $currentVersion.UBR
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name DisplayVersion -Value $currentVersion.DisplayVersion

  return $operatingSystemInfo
}

<#

.SYNOPSIS
Get the Win32_ComputerSystem information

.DESCRIPTION
Get the Win32_ComputerSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller.

#>
function getComputerSystemInfo() {
  return Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | `
    Microsoft.PowerShell.Utility\Select-Object TotalPhysicalMemory, DomainRole, Manufacturer, Model, NumberOfLogicalProcessors, Domain, Workgroup, DNSHostName, Name, PartOfDomain, SystemFamily, SystemSKUNumber
}

<#

.SYNOPSIS
Get SMBIOS locally from the passed in machineName


.DESCRIPTION
Get SMBIOS locally from the passed in machine name

#>
function getSmbiosData($computerSystem) {
  <#
    Array of chassis types.
    The following list of ChassisTypes is copied from the latest DMTF SMBIOS specification.
    REF: https://www.dmtf.org/sites/default/files/standards/documents/DSP0134_3.1.1.pdf
  #>
  $ChassisTypes =
  @{
    1  = 'Other'
    2  = 'Unknown'
    3  = 'Desktop'
    4  = 'Low Profile Desktop'
    5  = 'Pizza Box'
    6  = 'Mini Tower'
    7  = 'Tower'
    8  = 'Portable'
    9  = 'Laptop'
    10 = 'Notebook'
    11 = 'Hand Held'
    12 = 'Docking Station'
    13 = 'All in One'
    14 = 'Sub Notebook'
    15 = 'Space-Saving'
    16 = 'Lunch Box'
    17 = 'Main System Chassis'
    18 = 'Expansion Chassis'
    19 = 'SubChassis'
    20 = 'Bus Expansion Chassis'
    21 = 'Peripheral Chassis'
    22 = 'Storage Chassis'
    23 = 'Rack Mount Chassis'
    24 = 'Sealed-Case PC'
    25 = 'Multi-system chassis'
    26 = 'Compact PCI'
    27 = 'Advanced TCA'
    28 = 'Blade'
    29 = 'Blade Enclosure'
    30 = 'Tablet'
    31 = 'Convertible'
    32 = 'Detachable'
    33 = 'IoT Gateway'
    34 = 'Embedded PC'
    35 = 'Mini PC'
    36 = 'Stick PC'
  }

  $list = New-Object System.Collections.ArrayList
  $win32_Bios = Get-CimInstance -class Win32_Bios
  $obj = New-Object -Type PSObject | Microsoft.PowerShell.Utility\Select-Object SerialNumber, Manufacturer, UUID, BaseBoardProduct, ChassisTypes, Chassis, SystemFamily, SystemSKUNumber, SMBIOSAssetTag
  $obj.SerialNumber = $win32_Bios.SerialNumber
  $obj.Manufacturer = $win32_Bios.Manufacturer
  $computerSystemProduct = Get-CimInstance Win32_ComputerSystemProduct
  if ($null -ne $computerSystemProduct) {
    $obj.UUID = $computerSystemProduct.UUID
  }
  $baseboard = Get-CimInstance Win32_BaseBoard
  if ($null -ne $baseboard) {
    $obj.BaseBoardProduct = $baseboard.Product
  }
  $systemEnclosure = Get-CimInstance Win32_SystemEnclosure
  if ($null -ne $systemEnclosure) {
    $obj.SMBIOSAssetTag = $systemEnclosure.SMBIOSAssetTag
  }
  $obj.ChassisTypes = Get-CimInstance Win32_SystemEnclosure | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty ChassisTypes
  $obj.Chassis = New-Object -TypeName 'System.Collections.ArrayList'
  $obj.ChassisTypes | ForEach-Object -Process {
    $obj.Chassis.Add($ChassisTypes[[int]$_])
  }
  $obj.SystemFamily = $computerSystem.SystemFamily
  $obj.SystemSKUNumber = $computerSystem.SystemSKUNumber
  $list.Add($obj) | Out-Null

  return $list

}

<#

.SYNOPSIS
Get the azure arc status information

.DESCRIPTION
Get the azure arc status information

#>
function getAzureArcStatus() {

  $LogName = "Microsoft-ServerManagementExperience"
  $LogSource = "SMEScript"
  $ScriptName = "Get-ServerInventory.ps1 - getAzureArcStatus()"

  Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

  Get-Service -Name himds -ErrorVariable Err -ErrorAction SilentlyContinue | Out-Null

  if (!!$Err) {

    $Err = "The Azure arc agent is not installed. Details: $Err"

    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
    -Message "[$ScriptName]: $Err"  -ErrorAction SilentlyContinue

    $status = "NotInstalled"
  }
  else {
    $status = (azcmagent show --json | ConvertFrom-Json -ErrorAction Stop).status
  }

  return $status
}

<#

.SYNOPSIS
Gets an EnforcementMode that describes the system lockdown policy on this computer.

.DESCRIPTION
By checking the system lockdown policy, we can infer if PowerShell is in ConstrainedLanguage mode as a result of an enforced WDAC policy.
Note: $ExecutionContext.SessionState.LanguageMode should not be used within a trusted (by the WDAC policy) script context for this purpose because
the language mode returned would potentially not reflect the system-wide lockdown policy/language mode outside of the execution context.

#>
function getSystemLockdownPolicy() {
  return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy().ToString()
}

<#

.SYNOPSIS
Determines if the operating system is HCI.

.DESCRIPTION
Using the operating system 'Caption' (which corresponds to the 'ProductName' registry key at HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion) to determine if a server OS is HCI.

#>
function isServerOsHCI([string] $operatingSystemCaption) {
  return $operatingSystemCaption -eq "Microsoft Azure Stack HCI"
}

###########################################################################
# main()
###########################################################################

$operatingSystem = getOperatingSystemInfo
$computerSystem = getComputerSystemInfo
$isAdministrator = isUserAnAdministrator
$fqdnAndAddress = getHostFqdnAndAddress $computerSystem
$hostname = [Environment]::MachineName
$netbios = $env:ComputerName
$managementToolsInformation = getManagementToolsSupportInformation
$isWmfInstalled = isWMF5Installed $operatingSystem.Version
$clusterInformation = getClusterInformation -ErrorAction SilentlyContinue
$isHyperVPowershellInstalled = isHyperVPowerShellSupportInstalled
$isHyperVRoleInstalled = isHyperVRoleInstalled
$isCredSSPEnabled = isCredSSPEnabled
$isRemoteAppEnabled = isRemoteAppEnabled
$smbiosData = getSmbiosData $computerSystem
$azureArcStatus = getAzureArcStatus
$systemLockdownPolicy = getSystemLockdownPolicy
$isHciServer = isServerOsHCI $operatingSystem.Caption

$result = New-Object PSObject
$result | Add-Member -MemberType NoteProperty -Name 'IsAdministrator' -Value $isAdministrator
$result | Add-Member -MemberType NoteProperty -Name 'OperatingSystem' -Value $operatingSystem
$result | Add-Member -MemberType NoteProperty -Name 'ComputerSystem' -Value $computerSystem
$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $fqdnAndAddress.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $fqdnAndAddress.AddressList
$result | Add-Member -MemberType NoteProperty -Name 'Hostname' -Value $hostname
$result | Add-Member -MemberType NoteProperty -Name 'NetBios' -Value $netbios
$result | Add-Member -MemberType NoteProperty -Name 'IsManagementToolsAvailable' -Value $managementToolsInformation.ManagementToolsAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsServerManagerAvailable' -Value $managementToolsInformation.ServerManagerAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsWmfInstalled' -Value $isWmfInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCluster' -Value $clusterInformation.IsCluster
$result | Add-Member -MemberType NoteProperty -Name 'ClusterFqdn' -Value $clusterInformation.ClusterFqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2dEnabled' -Value $clusterInformation.IsS2dEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value $clusterInformation.IsBritannicaEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVRoleInstalled' -Value $isHyperVRoleInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVPowershellInstalled' -Value $isHyperVPowershellInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCredSSPEnabled' -Value $isCredSSPEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsRemoteAppEnabled' -Value $isRemoteAppEnabled
$result | Add-Member -MemberType NoteProperty -Name 'SmbiosData' -Value $smbiosData
$result | Add-Member -MemberType NoteProperty -Name 'AzureArcStatus' -Value $azureArcStatus
$result | Add-Member -MemberType NoteProperty -Name 'SystemLockdownPolicy' -Value $systemLockdownPolicy
$result | Add-Member -MemberType NoteProperty -Name 'IsHciServer' -Value $isHciServer

$result

}
## [END] Get-WACRDServerInventory ##
function Get-WACRDShareEntities {
<#

.SYNOPSIS
Enumerates all of the file system entities (files, folders, volumes) of the system.

.DESCRIPTION
Enumerates all of the file system entities (files, folders, volumes) of the system on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

.PARAMETER Path
    String -- The path to enumerate.

.PARAMETER OnlyFiles
    switch --

.PARAMETER OnlyFolders
    switch --

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $ComputerName
)

Set-StrictMode -Version 5.0

<#
.Synopsis
    Name: Get-FileSystemEntities
    Description: Gets all the local file system entities of the machine.

.Parameter Path
    String -- The path to enumerate.

.Returns
    The local file system entities.
#>
function Get-FileSystemEntities {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $ComputerName
    )

    return Invoke-Command -ComputerName $ComputerName -ScriptBlock { get-smbshare | Where-Object { -not ($_.name.EndsWith('$')) } } |
    Microsoft.PowerShell.Utility\Select-Object @{Name = "Caption"; Expression = { "\\" + $_.PSComputerName + "\" + $_.Name } },
    @{Name = "CreationDate"; Expression = { $_.CreationTimeUtc } },
    Extension,
    @{Name = "IsHidden"; Expression = { [bool]$false } },
    @{Name = "IsShared"; Expression = { [bool]$true } },
    Name,
    @{Name = "Type"; Expression = { "FileShare" } },
    @{Name = "LastModifiedDate"; Expression = { $_.LastWriteTimeUtc } },
    @{Name = "Size"; Expression = { if ($_.PSIsContainer) { $null } else { $_.Length } } };
}

$entities = Get-FileSystemEntities -ComputerName $ComputerName;

return $entities;

}
## [END] Get-WACRDShareEntities ##
function Get-WACRDSmb1InstallationStatus {
<#

.SYNOPSIS
Get SMB1 installation status.

.DESCRIPTION
Get SMB1 installation status.

.ROLE
Readers

#>

Import-Module DISM

$Enabled = [bool]( Get-WindowsOptionalFeature -online -featurename SMB1Protocol | Where-Object State -eq "Enabled")
@{ isEnabled = $Enabled }

}
## [END] Get-WACRDSmb1InstallationStatus ##
function Get-WACRDSmbFileShareDetails {
<#

.SYNOPSIS
Enumerates all of the smb local file shares of the system.

.DESCRIPTION
Enumerates all of the smb local file shares of the system.

.ROLE
Readers

#>

<#
.Synopsis
    Name: Get-SmbFileShareDetails
    Description: Retrieves the SMB shares on the computer.
.Returns
    The local smb file share(s).
#>

$shares = Get-SmbShare -includehidden | Where-Object {-not ($_.Name -eq "IPC$")} | Microsoft.PowerShell.Utility\Select-Object Name, Path, CachingMode, EncryptData, CurrentUsers, Special, LeasingMode, FolderEnumerationMode, CompressData

$uncPath = (Get-CimInstance Win32_ComputerSystem).DNSHostName + "." + (Get-CimInstance Win32_ComputerSystem).Domain

return @{
    shares = $shares;
    uncPath = $uncPath
}
   
}
## [END] Get-WACRDSmbFileShareDetails ##
function Get-WACRDSmbOverQuicSettings {
<#

.SYNOPSIS
Retrieves smb over QUIC settings from the server.

.DESCRIPTION
Returns smb over QUIC settings and server dns name

.ROLE
Readers

#>

Import-Module SmbShare

$serverConfigurationSettings = Get-SmbServerConfiguration | Microsoft.PowerShell.Utility\Select-Object DisableSmbEncryptionOnSecureConnection, RestrictNamedpipeAccessViaQuic

return @{
    serverConfigurationSettings = $serverConfigurationSettings
}

}
## [END] Get-WACRDSmbOverQuicSettings ##
function Get-WACRDSmbServerCertificateHealth {
<#

.SYNOPSIS
Retrieves health of the current certificate for smb over QUIC.

.DESCRIPTION
Retrieves health of the current certificate for smb over QUIC based on if the certificate is self signed or not.

.ROLE
Readers

#>

param (
  [Parameter(Mandatory = $true)]
  [String]
  $thumbprint,
  [Parameter(Mandatory = $true)]
  [boolean]
  $isSelfSigned,
  [Parameter(Mandatory = $true)]
  [boolean]
  $fromTaskScheduler,
  [Parameter(Mandatory = $false)]
  [String]
  $ResultFile,
  [Parameter(Mandatory = $false)]
  [String]
  $WarningsFile,
  [Parameter(Mandatory = $false)]
  [String]
  $ErrorFile
)

Set-StrictMode -Version 5.0



function getSmbServerCertificateHealth() {
  param (
    [String]
    $thumbprint,
    [boolean]
    $isSelfSigned,
    [String]
    $ResultFile,
    [String]
    $WarningsFile,
    [String]
    $ErrorFile
  )

  # create local runspace
  $ps = [PowerShell]::Create()
  # define input data but make it completed
  $inData = New-Object -Typename  System.Management.Automation.PSDataCollection[PSObject]
  $inData.Complete()
  # define output data to receive output
  $outData = New-Object -Typename  System.Management.Automation.PSDataCollection[PSObject]
  # register the script
  if ($isSelfSigned) {
    $ps.Commands.AddScript("Get-Item -Path " + $thumbprint + "| Test-Certificate -AllowUntrustedRoot") | Out-Null
  }
  else {
    $ps.Commands.AddScript("Get-Item -Path " + $thumbprint + "| Test-Certificate") | Out-Null
  }
  # execute async way.
  $async = $ps.BeginInvoke($inData, $outData)
  # wait for completion (callback will be called if any)
  $ps.EndInvoke($async)
  Start-Sleep -MilliSeconds 10
  # read output
  if ($outData.Count -gt 0) {
    @{ Output = $outData[0]; } | ConvertTo-Json | Out-File -FilePath $ResultFile
  }
  # read warnings
  if ($ps.Streams.Warning.Count -gt 0) {
    $ps.Streams.Warning | % { $_.ToString() } | Out-File -FilePath $WarningsFile
  }
  # read errors
  if ($ps.HadErrors) {
    $ps.Streams.Error | % { $_.ToString() } | Out-File -FilePath $ErrorFile
  }
}

#---- Script execution starts here ----
function isSystemLockdownPolicyEnforced() {
  return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy() -eq [System.Management.Automation.Security.SystemEnforcementMode]::Enforce
}
$isWdacEnforced = isSystemLockdownPolicyEnforced;

#In WDAC environment script file will already be available on the machine
#In WDAC mode the same script is executed - once normally and once through task Scheduler
if ($isWdacEnforced) {
  if ($fromTaskScheduler) {
    getSmbServerCertificateHealth $thumbprint $isSelfSigned $ResultFile $WarningsFile $ErrorFile;
    return;
  }
}
else {
  #In non-WDAC environment script file will not be available on the machine
  #Hence, a dynamic script is created which is executed through the task Scheduler
  $ScriptFile = $env:temp + "\smbOverQuic-certificateHealth.ps1"
}

$thumbprint = Join-Path "Cert:\LocalMachine\My" $thumbprint

# Pass parameters tpt and generate script file in temp folder
$ResultFile = $env:temp + "\smbOverQuic-certificateHealth_result.txt"
$WarningsFile = $env:temp + "\smbOverQuic-certificateHealth_warnings.txt"
$ErrorFile = $env:temp + "\smbOverQuic-certificateHealth_error.txt"
if (Test-Path $ErrorFile) {
  Remove-Item $ErrorFile
}

if (Test-Path $ResultFile) {
  Remove-Item $ResultFile
}

if (Test-Path $WarningsFile) {
  Remove-Item $WarningsFile
}
$isSelfSignedtemp = if ($isSelfSigned) { "`$true" } else { "`$false" }

if ($isWdacEnforced) {
  $arg = "-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -command ""&{Import-Module Microsoft.SME.FileExplorer; Get-WACFESmbServerCertificateHealth -fromTaskScheduler `$true -thumbprint $thumbprint -isSelfSigned $isSelfSignedtemp -ResultFile $ResultFile -WarningsFile $WarningsFile -ErrorFile $ErrorFile }"""
}
else {
  (Get-Command getSmbServerCertificateHealth).ScriptBlock | Set-Content -path $ScriptFile
  $arg = "-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -command ""&{Set-Location -Path $env:temp; .\smbOverQuic-certificateHealth.ps1 -thumbprint $thumbprint -isSelfSigned $isSelfSignedtemp -ResultFile $ResultFile -WarningsFile $WarningsFile -ErrorFile $ErrorFile }"""
}

# Create a scheduled task
$TaskName = "SMESmbOverQuicCertificate"
$User = [Security.Principal.WindowsIdentity]::GetCurrent()
$Role = (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if (!$Role) {
  Write-Warning "To perform some operations you must run an elevated Windows PowerShell console."
}

$Scheduler = New-Object -ComObject Schedule.Service
#Try to connect to schedule service 3 time since it may fail the first time
for ($i = 1; $i -le 3; $i++) {
  Try {
    $Scheduler.Connect()
    Break
  }
  Catch {
    if ($i -ge 3) {
      writeErrorLog "Can't connect to Schedule service"
      throw "Can't connect to Schedule service"
    }
    else {
      Start-Sleep -s 1
    }
  }
}

$RootFolder = $Scheduler.GetFolder("\")

#Delete existing task
if ($RootFolder.GetTasks(0) | Where-Object { $_.Name -eq $TaskName }) {
  Write-Debug("Deleting existing task" + $TaskName)
  try {
    $RootFolder.DeleteTask($TaskName, 0)
  }
  catch {

  }
}

$Task = $Scheduler.NewTask(0)
$RegistrationInfo = $Task.RegistrationInfo
$RegistrationInfo.Description = $TaskName
$RegistrationInfo.Author = $User.Name

$Triggers = $Task.Triggers
$Trigger = $Triggers.Create(7) #TASK_TRIGGER_REGISTRATION: Starts the task when the task is registered.
$Trigger.Enabled = $true

$Settings = $Task.Settings
$Settings.Enabled = $True
$Settings.StartWhenAvailable = $True
$Settings.Hidden = $False

$Action = $Task.Actions.Create(0)
$Action.Path = "powershell"
$Action.Arguments = $arg

#Tasks will be run with the highest privileges
$Task.Principal.RunLevel = 1

#### example Start the task with user specified invoke username and password
####$Task.Principal.LogonType = 1
####$RootFolder.RegisterTaskDefinition($TaskName, $Task, 6, $invokeUserName, $invokePassword, 1) | Out-Null

#### Start the task with SYSTEM creds
$RootFolder.RegisterTaskDefinition($TaskName, $Task, 6, "SYSTEM", $Null, 1) | Out-Null
#Wait for running task finished
$RootFolder.GetTask($TaskName).Run(0) | Out-Null
while ($RootFolder.GetTask($TaskName).State -ne 4) {
  Start-Sleep -MilliSeconds 10
}
while ($Scheduler.GetRunningTasks(0) | Where-Object { $_.Name -eq $TaskName }) {
  Start-Sleep -Seconds 1
}

#Clean up
try {
  $RootFolder.DeleteTask($TaskName, 0)
}
catch {

}
if (!$isWdacEnforced) {
  Remove-Item $ScriptFile
}

#Return result
if (Test-Path $ResultFile) {
  $result = Get-Content -Raw -Path $ResultFile | ConvertFrom-Json
  Remove-Item $ResultFile
  if ($result.Output) {
    return $result.Output;
  }
  else {
    if (Test-Path $WarningsFile) {
      $result = Get-Content -Path $WarningsFile
      Remove-Item $WarningsFile
    }
    if (Test-Path $ErrorFile) {
      Remove-Item $ErrorFile
    }
  }
  return $result;
}
else {
  if (Test-Path $ErrorFile) {
    $result = Get-Content -Path $ErrorFile
    Remove-Item $ErrorFile
    throw $result
  }
}

}
## [END] Get-WACRDSmbServerCertificateHealth ##
function Get-WACRDSmbServerCertificateMapping {
<#

.SYNOPSIS
Retrieves the current certifcate installed for smb over QUIC and smboverQuic status on the server.

.DESCRIPTION
Retrieves the current certifcate installed for smb over QUIC and smboverQuic status on the server.

.ROLE
Readers

#>

Import-Module SmbShare

$certHash = $null;
$kdcPort = $null;

function Retrieve-WACPort {
  $details = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ServerManagementGateway -ErrorAction SilentlyContinue
  if ($details) {
    $smePort = $details | Microsoft.PowerShell.Utility\Select-Object SmePort;
    return $smePort.smePort -eq 443
  }
  else {
    return $false;
  }
}

# retrieving smbcertificate mappings, if any
$smbCertificateMapping = @(Get-SmbServerCertificateMapping | Microsoft.PowerShell.Utility\Select-Object Thumbprint, Name)

# determining if WAC is installed on port 443
$isWacOnPort443 = Retrieve-WACPort;

# retrieving if smbOverQuic is enable on the server
$isSmbOverQuicEnabled = Get-SmbServerConfiguration | Microsoft.PowerShell.Utility\Select-Object EnableSMBQUIC

try {
  # retrieving kdc Proxy status on the server
  $kdcUrl = netsh http show urlacl | findstr /i "KdcProxy"
  if ($kdcUrl) {
    $pos = $kdcUrl.IndexOf("+")
    $rightPart = $kdcUrl.Substring($pos + 1)
    $pos1 = $rightPart.IndexOf("/")
    $kdcPort = $rightPart.SubString(1, $pos1 - 1)
  }

  [array]$path = Get-Item  -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -ErrorAction SilentlyContinue

  [array]$clientAuthproperty = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "HttpsClientAuth" -ErrorAction SilentlyContinue

  [array]$passwordAuthproperty = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "DisallowUnprotectedPasswordAuth" -ErrorAction SilentlyContinue

  $status = Get-Service -Name kpssvc | Microsoft.PowerShell.Utility\Select-Object Status

  if ($null -ne $kdcPort) {
    $port = "0.0.0.0:"
    $ipport = $port + $kdcPort
    $certBinding = @(netsh http show sslcert ipport=$ipport | findstr "Hash")
    if ($null -ne $certBinding -and $certBinding.count -gt 0) {
      $index = $certBinding[0].IndexOf(":")
      $certHash = $certBinding[0].Substring($index + 1).trim()
    }
  }

  $isKdcProxyMappedForSmbOverQuic = $false;
  $moreThanOneCertMapping = $false;

  if (($null -ne $certHash) -and ($null -ne $smbCertificateMapping) -and ($smbCertificateMapping.count -eq 1)) {
    $isKdcProxyMappedForSmbOverQuic = $smbCertificateMapping.thumbprint -eq $certHash
  }
  elseif ($null -ne $smbCertificateMapping -and $smbCertificateMapping.count -gt 1) {
    $set = New-Object System.Collections.Generic.HashSet[string]
    foreach ($mapping in $smbCertificateMapping) {
      # Adding Out null as set.Add always returns true/false and we do not want that.
      $set.Add($smbCertificateMapping.thumbprint) | Out-Null;
    }
    if ($set.Count -gt 1) {
      $moreThanOneCertMapping = $true;
    }
    if (!$moreThanOneCertMapping -and $null -ne $certHash) {
      $isKdcProxyMappedForSmbOverQuic = $smbCertificateMapping[0].thumbprint -eq $certHash
    }
  }
}
catch {
  throw $_
}

return @{
  smbCertificateMapping          = $smbCertificateMapping
  isSmbOverQuicEnabled           = $isSmbOverQuicEnabled
  isKdcProxyMappedForSmbOverQuic = $isKdcProxyMappedForSmbOverQuic
  kdcPort                        = $kdcPort
  isKdcProxyEnabled              = $kdcPort -and ($null -ne $path -and $path.count -gt 0) -and ($null -ne $clientAuthproperty -and $clientAuthproperty.count -gt 0) -and ($null -ne $passwordAuthproperty -and $passwordAuthproperty.count -gt 0) -and $status.status -eq "Running"
  isWacOnPort443                 = $isWacOnPort443;
}

}
## [END] Get-WACRDSmbServerCertificateMapping ##
function Get-WACRDSmbServerCertificateValues {
<#

.SYNOPSIS
Retrieves other values based on the installed certifcate for smb over QUIC.

.DESCRIPTION
Retrieves other values based on the installed certifcate for smb over QUIC.

.ROLE 
Readers

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $thumbprint
)

Set-Location Cert:\LocalMachine\My

[array] $smbServerDnsNames = Get-SmbServerCertificateMapping | Where-Object { $_.Thumbprint -eq $thumbprint } | Microsoft.PowerShell.Utility\Select-Object Name

$smbCertificateValues = Get-ChildItem -Recurse | Where-Object { $_.Thumbprint -eq $thumbprint } | Microsoft.PowerShell.Utility\Select-Object Subject, Thumbprint, Issuer, NotBefore, NotAfter

return @{ 
    smbServerDnsNames = $smbServerDnsNames
    smbCertificateValues = $smbCertificateValues
}

}
## [END] Get-WACRDSmbServerCertificateValues ##
function Get-WACRDSmbServerSettings {

<#

.SYNOPSIS
Enumerates the SMB server configuration settings on the computer.

.DESCRIPTION
Enumerates  the SMB server configuration settings on the computer.

.ROLE
Readers

#>

<#
.Synopsis
    Name: Get-SmbServerConfiguration
    Description: Retrieves the SMB server configuration settings on the computer.
.Returns
    SMB server configuration settings
#>

Import-Module SmbShare

$settings = Get-SmbServerConfiguration | Microsoft.PowerShell.Utility\Select-Object RequireSecuritySignature,RejectUnencryptedAccess,AuditSmb1Access,EncryptData,RequestCompression
$compressionSettings = Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\LanManServer\parameters" -Name "DisableCompression" -ErrorAction SilentlyContinue

@{ settings = $settings
    compressionSettings = $compressionSettings }

}
## [END] Get-WACRDSmbServerSettings ##
function Get-WACRDSmbShareAccess {
<#

.SYNOPSIS
Enumerates the SMB server access rights and details on the server.

.DESCRIPTION
Enumerates the SMB server access rights and details on the server.

.ROLE
Readers

#>

<#
.Synopsis
    Name: Get-SmbShareAccess
    Description: Retrieves the SMB server access rights and details on the server.
.Returns
    Retrieves the SMB server access rights and details on the server.
#>
param (
    [Parameter(Mandatory = $true)]
    [String]
    $Name
)

[array]$shareAccess = Get-SmbShareAccess -Name "$Name" | Microsoft.PowerShell.Utility\Select-Object AccountName, AccessRight, AccessControlType

$details = Get-SmbShare -Name "$Name" | Microsoft.PowerShell.Utility\Select-Object CachingMode, EncryptData, FolderEnumerationMode, CompressData

return @{ 
    details = $details
    shareAccess = $shareAccess
  }   
}
## [END] Get-WACRDSmbShareAccess ##
function Get-WACRDStorageFileShare {
<#

.SYNOPSIS
Enumerates all of the local file shares of the system.

.DESCRIPTION
Enumerates all of the local file shares of the system.

.ROLE
Readers

.PARAMETER FileShareId
    The file share ID.
#>
param (
    [Parameter(Mandatory = $false)]
    [String]
    $FileShareId
)

Import-Module CimCmdlets

<#
.Synopsis
    Name: Get-FileShares-Internal
    Description: Gets all the local file shares of the machine.

.Parameters
    $FileShareId: The unique identifier of the file share desired (Optional - for cases where only one file share is desired).

.Returns
    The local file share(s).
#>
function Get-FileSharesInternal
{
    param (
        [Parameter(Mandatory = $false)]
        [String]
        $FileShareId
    )

    Remove-Module Storage -ErrorAction Ignore; # Remove the Storage module to prevent it from automatically localizing

    $isDownlevel = [Environment]::OSVersion.Version.Major -lt 10;
    if ($isDownlevel)
    {
        # Map downlevel status to array of [health status, operational status, share state] uplevel equivalent
        $statusMap = @{
            "OK" =         @(0, 2, 1);
            "Error" =      @(2, 6, 2);
            "Degraded" =   @(1, 3, 2);
            "Unknown" =    @(5, 0, 0);
            "Pred Fail" =  @(1, 5, 2);
            "Starting" =   @(1, 8, 0);
            "Stopping" =   @(1, 9, 0);
            "Service" =    @(1, 11, 1);
            "Stressed" =   @(1, 4, 1);
            "NonRecover" = @(2, 7, 2);
            "No Contact" = @(2, 12, 2);
            "Lost Comm" =  @(2, 13, 2);
        };
        
        $shares = Get-CimInstance -ClassName Win32_Share |
            ForEach-Object {
                return @{
                    ContinuouslyAvailable = $false;
                    Description = $_.Description;
                    EncryptData = $false;
                    FileSharingProtocol = 3;
                    HealthStatus = $statusMap[$_.Status][0];
                    IsHidden = $_.Name.EndsWith("`$");
                    Name = $_.Name;
                    OperationalStatus = ,@($statusMap[$_.Status][1]);
                    ShareState = $statusMap[$_.Status][2];
                    UniqueId = "smb|" + (Get-CimInstance Win32_ComputerSystem).DNSHostName + "." + (Get-CimInstance Win32_ComputerSystem).Domain + "\" + $_.Name;
                    VolumePath = $_.Path;
                }
            }
    }
    else
    {        
        $shares = Get-CimInstance -ClassName MSFT_FileShare -Namespace Root\Microsoft\Windows/Storage |
            ForEach-Object {
                return @{
                    IsHidden = $_.Name.EndsWith("`$");
                    VolumePath = $_.VolumeRelativePath;
                    ContinuouslyAvailable = $_.ContinuouslyAvailable;
                    Description = $_.Description;
                    EncryptData = $_.EncryptData;
                    FileSharingProtocol = $_.FileSharingProtocol;
                    HealthStatus = $_.HealthStatus;
                    Name = $_.Name;
                    OperationalStatus = $_.OperationalStatus;
                    UniqueId = $_.UniqueId;
                    ShareState = $_.ShareState;
                }
            }
    }

    if ($FileShareId)
    {
        $shares = $shares | Where-Object { $_.UniqueId -eq $FileShareId };
    }

    return $shares;
}

if ($FileShareId)
{
    Get-FileSharesInternal -FileShareId $FileShareId;
}
else
{
    Get-FileSharesInternal;
}

}
## [END] Get-WACRDStorageFileShare ##
function Get-WACRDTempFolderPath {
<#

.SYNOPSIS
Gets the temporary folder (%temp%) for the user.

.DESCRIPTION
Gets the temporary folder (%temp%) for the user.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

Set-StrictMode -Version 5.0

return $env:TEMP

}
## [END] Get-WACRDTempFolderPath ##
function Move-WACRDFile {
<#

.SYNOPSIS
Moves or Copies a file or folder

.DESCRIPTION
Moves or Copies a file or folder from the source location to the destination location
Folders will be copied recursively

.ROLE
Administrators

.PARAMETER Path
    String -- the path to the source file/folder to copy

.PARAMETER Destination
    String -- the path to the new location

.PARAMETER Copy
    boolean -- Determine action to be performed

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $path,

    [Parameter(Mandatory = $true)]
    [String]
    $destination,

    [Parameter(Mandatory = $true)]
    [boolean]
    $copy,

    [Parameter(Mandatory = $true)]
    [string]
    $entityType,

    [Parameter(Mandatory = $false)]
    [string]
    $existingPath
)

Set-StrictMode -Version 5.0

if($copy){
  $result = Copy-Item -Path $path -Destination $destination -Recurse -Force -PassThru -ErrorAction SilentlyContinue
  if(!$result){
    return $Error[0].Exception.Message
  }
} else {
  if ($entityType -eq "File" -Or  !$existingPath) {
    $result = Move-Item -Path $path -Destination $destination -Force -PassThru -ErrorAction SilentlyContinue
    if(!$result){
      return $Error[0].Exception.Message
    }
  }
  else {
    # Move-Item -Force doesn't work when replacing folders, remove destination folder before replacing
    Remove-Item -Path $existingPath -Confirm:$false -Force -Recurse
    $forceResult = Move-Item -Path $path -Destination $destination -Force -PassThru -ErrorAction SilentlyContinue
    if (!$forceResult) {
      return $Error[0].Exception.Message
    }
  }
}
return $result

}
## [END] Move-WACRDFile ##
function New-WACRDFile {
<#

.SYNOPSIS
Create a new file.

.DESCRIPTION
Create a new file on this server. If the file already exists, it will be overwritten.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- the path to the parent of the new file.

.PARAMETER NewName
    String -- the file name.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $NewName
)

Set-StrictMode -Version 5.0

$newItem = New-Item -ItemType File -Path (Join-Path -Path $Path -ChildPath $NewName) -Force

return $newItem |
    Microsoft.PowerShell.Utility\Select-Object @{Name = "Caption"; Expression = { $_.FullName } },
                  @{Name = "CreationDate"; Expression = { $_.CreationTimeUtc } },
                  Extension,
                  @{Name = "IsHidden"; Expression = { $_.Attributes -match "Hidden" } },
                  @{Name = "IsShared"; Expression = { [bool]($folderShares | Where-Object Path -eq $_.FullName) } },
                  Name,
                  @{Name = "Type"; Expression = { Get-FileSystemEntityType -Attributes $_.Attributes } },
                  @{Name = "LastModifiedDate"; Expression = { $_.LastWriteTimeUtc } },
                  @{Name = "Size"; Expression = { if ($_.PSIsContainer) { $null } else { $_.Length } } };

}
## [END] New-WACRDFile ##
function New-WACRDFolder {
<#

.SYNOPSIS
Create a new folder.

.DESCRIPTION
Create a new folder on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- the path to the parent of the new folder.

.PARAMETER NewName
    String -- the folder name.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $NewName
)

Set-StrictMode -Version 5.0

$pathSeparator = [System.IO.Path]::DirectorySeparatorChar;
$newItem = New-Item -ItemType Directory -Path ($Path.TrimEnd($pathSeparator) + $pathSeparator + $NewName)

return $newItem |
    Microsoft.PowerShell.Utility\Select-Object @{Name="Caption"; Expression={$_.FullName}},
                  @{Name="CreationDate"; Expression={$_.CreationTimeUtc}},
                  Extension,
                  @{Name="IsHidden"; Expression={$_.Attributes -match "Hidden"}},
                  Name,
                  @{Name="Type"; Expression={Get-FileSystemEntityType -Attributes $_.Attributes}},
                  @{Name="LastModifiedDate"; Expression={$_.LastWriteTimeUtc}},
                  @{Name="Size"; Expression={if ($_.PSIsContainer) { $null } else { $_.Length }}};

}
## [END] New-WACRDFolder ##
function New-WACRDSmbFileShare {
<#

.SYNOPSIS
Gets the SMB file share  details on the server.

.DESCRIPTION
Gets the SMB file share  details on the server.

.ROLE
Administrators

#>

<#
.Synopsis
    Name: New-SmbFileShare
    Description: Gets the SMB file share  details on the server.
.Returns
    Retrieves all the SMB file share  details on the server.
#>


param (
    [Parameter(Mandatory = $true)]
    [String]
    $path,    

    [Parameter(Mandatory = $true)]
    [String]
    $name,

    [Parameter(Mandatory = $false)]
    [String[]]
    $fullAccess,

    [Parameter(Mandatory = $false)]
    [String[]]
    $changeAccess,

    [Parameter(Mandatory = $false)]
    [String[]]
    $readAccess,

    [Parameter(Mandatory = $false)]
    [String[]]
    $noAccess,

    [Parameter(Mandatory = $false)]
    [Int]
    $cachingMode,

    [Parameter(Mandatory = $false)]
    [boolean]
    $encryptData,

    # TODO:
    # [Parameter(Mandatory = $false)]
    # [Int]
    # $FolderEnumerationMode

    [Parameter(Mandatory = $false)]
    [boolean]
    $compressData,

    [Parameter(Mandatory = $false)]
    [boolean]
    $isCompressDataEnabled
)

$HashArguments = @{
  Name = "$name"
}

if($fullAccess.count -gt 0){
    $HashArguments.Add("FullAccess", $fullAccess)
}
if($changeAccess.count -gt 0){
    $HashArguments.Add("ChangeAccess", $changeAccess)
}
if($readAccess.count -gt 0){
    $HashArguments.Add("ReadAccess", $readAccess)
}
if($noAccess.count -gt 0){
    $HashArguments.Add("NoAccess", $noAccess)
}
if($cachingMode){
     $HashArguments.Add("CachingMode", "$cachingMode")
}
if($encryptData -ne $null){
    $HashArguments.Add("EncryptData", $encryptData)
}
# TODO: if($FolderEnumerationMode -eq 0){
#     $HashArguments.Add("FolderEnumerationMode", "AccessBased")
# } else {
#     $HashArguments.Add("FolderEnumerationMode", "Unrestricted")
# }
if($isCompressDataEnabled){
    $HashArguments.Add("CompressData", $compressData)
}

New-SmbShare -Path "$path" @HashArguments
@{ shareName = $name } 

}
## [END] New-WACRDSmbFileShare ##
function Remove-WACRDAllShareNames {
<#

.SYNOPSIS
Removes all shares of a folder.

.DESCRIPTION
Removes all shares of a folder.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to the folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path    
)

Set-StrictMode -Version 5.0

$CimInstance = Get-CimInstance -Class Win32_Share -Filter Path="'$Path'"
$RemoveShareCommand = ''
if ($CimInstance.name -And $CimInstance.name.GetType().name -ne 'String') { $RemoveShareCommand = $CimInstance.ForEach{ 'Remove-SmbShare -Name "' + $_.name + '" -Force'} } 
Else { $RemoveShareCommand = 'Remove-SmbShare -Name "' + $CimInstance.Name + '" -Force'}
if($RemoveShareCommand) { $RemoveShareCommand.ForEach{ Invoke-Expression $_ } }


}
## [END] Remove-WACRDAllShareNames ##
function Remove-WACRDFileSystemEntity {
<#

.SYNOPSIS
Remove the passed in file or path.

.DESCRIPTION
Remove the passed in file or path from this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER path
    String -- the file or path to remove.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path
)

Set-StrictMode -Version 5.0

Remove-Item -Path $Path -Confirm:$false -Force -Recurse

}
## [END] Remove-WACRDFileSystemEntity ##
function Remove-WACRDFolderShareUser {
<#

.SYNOPSIS
Removes a user from the folder access.

.DESCRIPTION
Removes a user from the folder access.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to the folder.

.PARAMETER Identity
    String -- The user identification (AD / Local user).

.PARAMETER FileSystemRights
    String -- File system rights of the user.

.PARAMETER AccessControlType
    String -- Access control type of the user.    

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $Identity,

    [Parameter(Mandatory = $true)]
    [String]
    $FileSystemRights,

    [ValidateSet('Deny','Allow')]
    [Parameter(Mandatory = $true)]
    [String]
    $AccessControlType
)

Set-StrictMode -Version 5.0

function Remove-UserPermission
{
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Path,
    
        [Parameter(Mandatory = $true)]
        [String]
        $Identity,
        
        [ValidateSet('Deny','Allow')]
        [Parameter(Mandatory = $true)]
        [String]
        $ACT
    )

    $Acl = Get-Acl $Path
    $AccessRule = New-Object system.security.accesscontrol.filesystemaccessrule($Identity, 'ReadAndExecute','ContainerInherit, ObjectInherit', 'None', $ACT)
    $Acl.RemoveAccessRuleAll($AccessRule)
    Set-Acl $Path $Acl
}

Remove-UserPermission $Path $Identity 'Allow'
Remove-UserPermission $Path $Identity 'Deny'
}
## [END] Remove-WACRDFolderShareUser ##
function Remove-WACRDSmbServerCertificateMapping {
<#
.SYNOPSIS
Removes the currently installed certificate for smb over QUIC on the server.

.DESCRIPTION
Removes the currently installed certificate for smb over QUIC on the server and also sets the status of smbOverQuic to enabled on the server.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $false)]
    [String[]]
    $ServerDNSNames,

    [Parameter(Mandatory = $false)]
    [String]
    $KdcPort,

    [Parameter(Mandatory = $true)]
    [boolean]
    $IsKdcProxyMappedForSmbOverQuic
)

Set-Location Cert:\LocalMachine\My


if($ServerDNSNames.count -gt 0){
    foreach($serverDNSName in $ServerDNSNames){
        Remove-SmbServerCertificateMapping -Name $serverDNSName -Force
    }
}

if($IsKdcProxyMappedForSmbOverQuic -and $KdcPort -ne $null){
    $port = "0.0.0.0:"
    $ipport = $port+$KdcPort
    $deleteCertKdc = netsh http delete sslcert ipport=$ipport
    if ($LASTEXITCODE -ne 0) {
        throw $deleteCertKdc
    }
    $firewallString = "KDC Proxy Server service (KPS) for SMB over QUIC"
    Remove-NetFirewallRule -DisplayName $firewallString
}

Set-SmbServerConfiguration -EnableSMBQUIC $true -Force
}
## [END] Remove-WACRDSmbServerCertificateMapping ##
function Remove-WACRDSmbShare {
<#

.SYNOPSIS
Removes shares of a folder.

.DESCRIPTION
Removes selected shares of a folder.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER name
    String -- The name of the folder.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Name    
)

Remove-SmbShare -Name $Name -Force


}
## [END] Remove-WACRDSmbShare ##
function Rename-WACRDFileSystemEntity {
<#

.SYNOPSIS
Rename a folder.

.DESCRIPTION
Rename a folder on this server.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- the path to the folder.

.PARAMETER NewName
    String -- the new folder name.

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    [Parameter(Mandatory = $true)]
    [String]
    $NewName
)

Set-StrictMode -Version 5.0

<#
.Synopsis
    Name: Get-FileSystemEntityType
    Description: Gets the type of a local file system entity.

.Parameters
    $Attributes: The System.IO.FileAttributes of the FileSystemEntity.

.Returns
    The type of the local file system entity.
#>
function Get-FileSystemEntityType
{
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileAttributes]
        $Attributes
    )

    if ($Attributes -match "Directory")
    {
        return "Folder";
    }
    else
    {
        return "File";
    }
}

Rename-Item -Path $Path -NewName $NewName -PassThru |
    Microsoft.PowerShell.Utility\Select-Object @{Name="Caption"; Expression={$_.FullName}},
                @{Name="CreationDate"; Expression={$_.CreationTimeUtc}},
                Extension,
                @{Name="IsHidden"; Expression={$_.Attributes -match "Hidden"}},
                Name,
                @{Name="Type"; Expression={Get-FileSystemEntityType -Attributes $_.Attributes}},
                @{Name="LastModifiedDate"; Expression={$_.LastWriteTimeUtc}},
                @{Name="Size"; Expression={if ($_.PSIsContainer) { $null } else { $_.Length }}};

}
## [END] Rename-WACRDFileSystemEntity ##
function Restore-WACRDConfigureSmbServerCertificateMapping {
<#
.SYNOPSIS
Rolls back to the previous state of certificate maping if configure/edit action failed.

.DESCRIPTION
Rolls back to the previous state of certificate maping if configure/edit action failed.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [boolean]
    $IsKdcProxyEnabled,

    [Parameter(Mandatory = $false)]
    [string]
    $KdcPort,

    [Parameter(Mandatory = $true)]
    [string]
    $KdcProxyOptionSelected
)

[array]$smbCertificateMappings = Get-SmbServerCertificateMapping | Microsoft.PowerShell.Utility\Select-Object Name
[array]$mappingNames = $smbCertificateMappings.name

if ($mappingNames.count -eq 0) {
    return;
}
if ($mappingNames.count -gt 0) {
    foreach ($mappingName in $mappingNames) {
        Remove-SmbServerCertificateMapping -Name $mappingName -Force
    }
}

if (!$IsKdcProxyEnabled -and $KdcProxyOptionSelected -eq 'enabled' -and $KdcPort -ne $null) {
    $urlLeft = "https://+:"
    $urlRight = "/KdcProxy/"
    $url = $urlLeft + $KdcPort + $urlRight
    $output = @(netsh http show urlacl | findstr /i "KdcProxy")
    if ($LASTEXITCODE -ne 0) {
        throw $output
    }

    if ($null -ne $output -and $output.count -gt 0) {
        $deleteOutput = netsh http delete urlacl url=$url
        if ($LASTEXITCODE -ne 0) {
            throw $deleteOutput
        }
    }
    [array]$clientAuthproperty = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "HttpsClientAuth" -ErrorAction SilentlyContinue
    if ($null -ne $clientAuthproperty -and $clientAuthproperty.count -gt 0) {
        Remove-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "HttpsClientAuth"
    }
    [array]$passwordAuthproperty = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "DisallowUnprotectedPasswordAuth" -ErrorAction SilentlyContinue
    if ($null -ne $passwordAuthproperty -and $passwordAuthproperty.count -gt 0) {
        Remove-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "DisallowUnprotectedPasswordAuth"
    }
    [array]$path = Get-Item  -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -ErrorAction SilentlyContinue
    if ($null -ne $path -and $path.count -gt 0) {
        Remove-Item  -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings"
    }
    Stop-Service -Name kpssvc
    $firewallString = "KDC Proxy Server service (KPS) for SMB over QUIC"
    $rule = @(Get-NetFirewallRule -DisplayName $firewallString -ErrorAction SilentlyContinue)
    if($null -ne $rule -and $rule.count -gt 0) {
        Remove-NetFirewallRule -DisplayName $firewallString
    }

    $port = "0.0.0.0:"
    $ipport = $port+$KdcPort
    $certBinding =  @(netsh http show sslcert ipport=$ipport | findstr "Hash")
    if ($LASTEXITCODE -ne 0) {
        throw $certBinding
    }
    if($null -ne $certBinding -and $certBinding.count -gt 0) {
        $deleteCertKdc = netsh http delete sslcert ipport=$ipport
        if ($LASTEXITCODE -ne 0) {
            throw $deleteCertKdc
        } 
    }
}


}
## [END] Restore-WACRDConfigureSmbServerCertificateMapping ##
function Set-WACRDSmbOverQuicServerSettings {
<#

.SYNOPSIS
Sets smb server settings for QUIC.

.DESCRIPTION
Sets smb server settings for QUIC.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER disableSmbEncryptionOnSecureConnection
To enable or disable smbEncryption on the server.

.PARAMETER restrictNamedPipeAccessViaQuic
To enable or diable namedPipeAccess on the server.

#>

param (
[Parameter(Mandatory = $true)]
[boolean]
$disableSmbEncryptionOnSecureConnection,

[Parameter(Mandatory = $true)]
[boolean]
$restrictNamedPipeAccessViaQuic
)


Set-SmbServerConfiguration -DisableSmbEncryptionOnSecureConnection $disableSmbEncryptionOnSecureConnection -RestrictNamedPipeAccessViaQuic $restrictNamedPipeAccessViaQuic -Force;

}
## [END] Set-WACRDSmbOverQuicServerSettings ##
function Set-WACRDSmbServerCertificateMapping {
<#
.SYNOPSIS
Set Smb Server Certificate Mapping

.DESCRIPTION
Configures smb over QUIC.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Thumbprint
    String -- The thumbprint of the certifiacte selected.

.PARAMETER ServerDNSNames
    String[] -- The addresses of the server for certificate mapping.
#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Thumbprint,

    [Parameter(Mandatory = $true)]
    [String[]]
    $ServerDNSNames,

    [Parameter(Mandatory = $true)]
    [boolean]
    $IsKdcProxyEnabled,

    [Parameter(Mandatory = $true)]
    [String]
    $KdcProxyOptionSelected,

    [Parameter(Mandatory = $false)]
    [String]
    $KdcPort
)

Import-Module -Name Microsoft.PowerShell.Management -ErrorAction SilentlyContinue
Set-Variable -Name LogName -Option Constant -Value "Microsoft-ServerManagementExperience" -ErrorAction SilentlyContinue
Set-Variable -Name LogSource -Option Constant -Value "SmeScripts-ConfigureKdcProxy" -ErrorAction SilentlyContinue

function writeInfoLog($logMessage) {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
        -Message $logMessage -ErrorAction SilentlyContinue
}

Set-Location Cert:\LocalMachine\My

function Enable-KdcProxy {

    $urlLeft = "https://+:"
    $urlRight = "/KdcProxy"
    $url = $urlLeft + $KdcPort + $urlRight

    $port = "0.0.0.0:"
    $ipport = $port+$KdcPort

    $ComputerName = (Get-CimInstance Win32_ComputerSystem).DNSHostName + "." + (Get-CimInstance Win32_ComputerSystem).Domain

    try {
        $certBinding = @(netsh http show sslcert ipport=$ipport | findstr "Hash")
        if($null -ne $certBinding -and $certBinding.count -gt 0) {
            $deleteCertKdc = netsh http delete sslcert ipport=$ipport
            if ($LASTEXITCODE -ne 0) {
                throw $deleteCertKdc
            }
        }
        $guid = [Guid]::NewGuid()
        $netshAddCertBinding = netsh http add sslcert ipport=$ipport certhash=$Thumbprint certstorename="my" appid="{$guid}"
        if ($LASTEXITCODE -ne 0) {
          throw $netshAddCertBinding
        }
        $message = 'Completed adding ssl certificate port'
        writeInfoLog $message
        if ($ServerDNSNames.count -gt 0) {
            foreach ($serverDnsName in $ServerDnsNames) {
                if($ComputerName.trim() -ne $serverDnsName.trim()){
                    $output = Echo 'Y' | netdom computername $ComputerName /add $serverDnsName
                    if ($LASTEXITCODE -ne 0) {
                        throw $output
                    }
                }
            }
            $message = 'Completed adding alternative names'
            writeInfoLog $message
        }
        if(!$IsKdcProxyEnabled) {
            $netshOutput = netsh http add urlacl url=$url user="NT authority\Network Service"
            if ($LASTEXITCODE -ne 0) {
                throw $netshOutput
            }
            $message = 'Completed adding urlacl'
            writeInfoLog $message
            New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings"  -force
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "HttpsClientAuth" -Value 0x0 -type DWORD
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\KPSSVC\Settings" -Name "DisallowUnprotectedPasswordAuth" -Value 0x0 -type DWORD
            Set-Service -Name kpssvc -StartupType Automatic
            Start-Service -Name kpssvc
        }
        $message = 'Returning method Enable-KdcProxy'
        writeInfoLog $message
        return $true;
    }
    catch {
        throw $_
    }
}
function New-SmbCertificateMapping {
    if ($ServerDNSNames.count -gt 0) {
        $message = 'Setting server dns names'
        writeInfoLog $message
        foreach ($serverDNSName in $ServerDNSNames) {
            New-SmbServerCertificateMapping -Name $serverDNSName -Thumbprint $Thumbprint -StoreName My -Force
        }
        if ($KdcProxyOptionSelected -eq "enabled" -and $KdcPort -ne $null){
            $message = 'Enabling Kdc Proxy'
            writeInfoLog $message
            $result = Enable-KdcProxy
            if($result) {
                $firewallString = "KDC Proxy Server service (KPS) for SMB over QUIC"
                $firewallDesc = "The KDC Proxy Server service runs on edge servers to proxy Kerberos protocol messages to domain controllers on the corporate network. Default port is TCP/443."
                New-NetFirewallRule -DisplayName $firewallString -Description $firewallDesc -Protocol TCP -LocalPort $KdcPort -Direction Inbound -Action Allow
            }
        }
        return $true;
    }
    $message = 'Exiting method smb certificate mapping '
    writeInfoLog $message
    return $true;
}

return New-SmbCertificateMapping;

}
## [END] Set-WACRDSmbServerCertificateMapping ##
function Set-WACRDSmbServerSettings {
<#
.SYNOPSIS
Updates the server configuration settings on the server.

.DESCRIPTION
Updates the server configuration settings on the server.

.ROLE
Administrators

#>

<#
.Synopsis
    Name: Set-SmbServerSettings
    Description: Updates the server configuration settings on the server.
#>


param (
    [Parameter(Mandatory = $false)]
    [Nullable[boolean]]
    $AuditSmb1Access,

    [Parameter(Mandatory = $true)]
    [boolean]
    $RequireSecuritySignature,

    [Parameter(Mandatory = $false)]
    [boolean]
    $RejectUnencryptedAccess,

    [Parameter(Mandatory = $true)]
    [boolean]
    $EncryptData,

    [Parameter(Mandatory = $true)]
    [boolean]
    $CompressionSettingsClicked,

    [Parameter(Mandatory = $false)]
    [Nullable[boolean]]
    $RequestCompression
)

$HashArguments = @{
}

if($CompressionSettingsClicked) {
    Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\LanManServer\parameters" -Name "DisableCompression" -Value 0x1 -type DWORD
} else {
    Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\LanManServer\parameters" -Name "DisableCompression" -Value 0x0 -type DWORD
}

if($null -ne $RejectUnencryptedAccess){
    $HashArguments.Add("RejectUnencryptedAccess", $RejectUnencryptedAccess)
}

if($null -ne $AuditSmb1Access){
  $HashArguments.Add("AuditSmb1Access", $AuditSmb1Access)
}

if($null -ne $RequestCompression) {
  $HashArguments.Add("RequestCompression", $RequestCompression)
}

Set-SmbServerConfiguration -RequireSecuritySignature $RequireSecuritySignature -EncryptData $EncryptData @HashArguments -Force

}
## [END] Set-WACRDSmbServerSettings ##
function Test-WACRDFileSystemEntity {
<#

.SYNOPSIS
Checks if a file or folder exists

.DESCRIPTION
Checks if a file or folder exists
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

.PARAMETER Path
    String -- The path to check if it exists

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $Path    
)

Set-StrictMode -Version 5.0

Test-Path -path $Path

}
## [END] Test-WACRDFileSystemEntity ##
function Uninstall-WACRDSmb1 {
<#
.SYNOPSIS
Disables SMB1 on the server.

.DESCRIPTION
Disables SMB1 on the server.

.ROLE
Administrators

#>

<#
.Synopsis
    Name: UninstallSmb1
    Description: Disables SMB1 on the server.
#>

Disable-WindowsOptionalFeature -Online -FeatureName smb1protocol -NoRestart
}
## [END] Uninstall-WACRDSmb1 ##

# SIG # Begin signature block
# MIIoKgYJKoZIhvcNAQcCoIIoGzCCKBcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAVKYQXDW2sX9n2
# VuEJNOmqFp7xAHAKhP44wmajcKCud6CCDXYwggX0MIID3KADAgECAhMzAAADTrU8
# esGEb+srAAAAAANOMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI5WhcNMjQwMzE0MTg0MzI5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDdCKiNI6IBFWuvJUmf6WdOJqZmIwYs5G7AJD5UbcL6tsC+EBPDbr36pFGo1bsU
# p53nRyFYnncoMg8FK0d8jLlw0lgexDDr7gicf2zOBFWqfv/nSLwzJFNP5W03DF/1
# 1oZ12rSFqGlm+O46cRjTDFBpMRCZZGddZlRBjivby0eI1VgTD1TvAdfBYQe82fhm
# WQkYR/lWmAK+vW/1+bO7jHaxXTNCxLIBW07F8PBjUcwFxxyfbe2mHB4h1L4U0Ofa
# +HX/aREQ7SqYZz59sXM2ySOfvYyIjnqSO80NGBaz5DvzIG88J0+BNhOu2jl6Dfcq
# jYQs1H/PMSQIK6E7lXDXSpXzAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUnMc7Zn/ukKBsBiWkwdNfsN5pdwAw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMDUxNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAD21v9pHoLdBSNlFAjmk
# mx4XxOZAPsVxxXbDyQv1+kGDe9XpgBnT1lXnx7JDpFMKBwAyIwdInmvhK9pGBa31
# TyeL3p7R2s0L8SABPPRJHAEk4NHpBXxHjm4TKjezAbSqqbgsy10Y7KApy+9UrKa2
# kGmsuASsk95PVm5vem7OmTs42vm0BJUU+JPQLg8Y/sdj3TtSfLYYZAaJwTAIgi7d
# hzn5hatLo7Dhz+4T+MrFd+6LUa2U3zr97QwzDthx+RP9/RZnur4inzSQsG5DCVIM
# pA1l2NWEA3KAca0tI2l6hQNYsaKL1kefdfHCrPxEry8onJjyGGv9YKoLv6AOO7Oh
# JEmbQlz/xksYG2N/JSOJ+QqYpGTEuYFYVWain7He6jgb41JbpOGKDdE/b+V2q/gX
# UgFe2gdwTpCDsvh8SMRoq1/BNXcr7iTAU38Vgr83iVtPYmFhZOVM0ULp/kKTVoir
# IpP2KCxT4OekOctt8grYnhJ16QMjmMv5o53hjNFXOxigkQWYzUO+6w50g0FAeFa8
# 5ugCCB6lXEk21FFB1FdIHpjSQf+LP/W2OV/HfhC3uTPgKbRtXo83TZYEudooyZ/A
# Vu08sibZ3MkGOJORLERNwKm2G7oqdOv4Qj8Z0JrGgMzj46NFKAxkLSpE5oHQYP1H
# tPx1lPfD7iNSbJsP6LiUHXH1MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAANOtTx6wYRv6ysAAAAAA04wDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIKZOOvJohIAKXqmOnAOe1g/x
# fRG/Ur7y6xqm7btbhVizMEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEAlBEyTB7pP+ISjac6NQqPGXctUZFK0lcmC9LDAxCeEsrFmhxBabTBjXYx
# TRQA2oqFXih1vAzHuhUFSvp+z8iWUvmpmiG8KQuLCA865ZIM3bkR1NcQc6RObrVZ
# k+laJ8KY1Ao6ZymFz9j3XUSYyBTIL2fDTlkFhHcpmanp2KM0ZxlVS1ekpALdI7pT
# egZwdUqtSDAj5wWlK0QVjqaue+oJsCUVLetyFLqDJm6gJdxPsKqVeXfkTaUv/v+6
# DcSkj2o9pY9tDHiWike6xzZ617KIMH6QOFjOchVRcPhaDE64KCJtFzGiTXw8L6D0
# 8WOgAu8MLQbM0a4IDoWdtqZYQb+uQKGCF5QwgheQBgorBgEEAYI3AwMBMYIXgDCC
# F3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsq
# hkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCClFIRdUuVTp6xBkAa1eiTebYm4f13ejIdjkP8vT53o5gIGZVbCyozn
# GBMyMDIzMTIwNzAzNDMyNS4xMzFaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTQwMC0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wg
# ghHqMIIHIDCCBQigAwIBAgITMwAAAdYnaf9yLVbIrgABAAAB1jANBgkqhkiG9w0B
# AQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMzA1MjUxOTEy
# MzRaFw0yNDAyMDExOTEyMzRaMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTQwMC0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDPLM2Om8r5u3fcbDEOXydJtbkW5U34KFaftC+8QyNq
# plMIzSTC1ToE0zcweQCvPIfpYtyPB3jt6fPRprvhwCksUw9p0OfmZzWPDvkt40BU
# Stu813QlrloRdplLz2xpk29jIOZ4+rBbKaZkBPZ4R4LXQhkkHne0Y/Yh85ZqMMRa
# MWkBM6nUwV5aDiwXqdE9Jyl0i1aWYbCqzwBRdN7CTlAJxrJ47ov3uf/lFg9hnVQc
# qQYgRrRFpDNFMOP0gwz5Nj6a24GZncFEGRmKwImL+5KWPnVgvadJSRp6ZqrYV3Fm
# bBmZtsF0hSlVjLQO8nxelGV7TvqIISIsv2bQMgUBVEz8wHFyU3863gHj8BCbEpJz
# m75fLJsL3P66lJUNRN7CRsfNEbHdX/d6jopVOFwF7ommTQjpU37A/7YR0wJDTt6Z
# sXU+j/wYlo9b22t1qUthqjRs32oGf2TRTCoQWLhJe3cAIYRlla/gEKlbuDDsG392
# 6y4EMHFxTjsjrcZEbDWwjB3wrp11Dyg1QKcDyLUs2anBolvQwJTN0mMOuXO8tBz2
# 0ng/+Xw+4w+W9PMkvW1faYi435VjKRZsHfxIPjIzZ0wf4FibmVPJHZ+aTxGsVJPx
# ydChvvGCf4fe8XfYY9P5lbn9ScKc4adTd44GCrBlJ/JOsoA4OvNHY6W+XcKVcIIG
# WwIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFGGaVDY7TQBiMCKg2+j/zRTcYsZOMB8G
# A1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggrBgEFBQcBAQRgMF4w
# XAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2Vy
# dHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQD
# AgeAMA0GCSqGSIb3DQEBCwUAA4ICAQDUv+RjNidwJxSbMk1IvS8LfxNM8VaVhpxR
# 1SkW+FRY6AKkn2s3On29nGEVlatblIv1qVTKkrUxLYMZ0z+RA6mmfXue2Y7/YBbz
# M5kUeUgU2y1Mmbin6xadT9DzECeE7E4+3k2DmZxuV+GLFYQsqkDbe8oy7+3BSiU2
# 9qyZAYT9vRDALPUC5ZwyoPkNfKbqjl3VgFTqIubEQr56M0YdMWlqCqq0yVln9mPb
# hHHzXHOjaQsurohHCf7VT8ct79po34Fd8XcsqmyhdKBy1jdyknrik+F3vEl/90qa
# on5N8KTZoGtOFlaJFPnZ2DqQtb2WWkfuAoGWrGSA43Myl7+PYbUsri/NrMvAd9Z+
# J9FlqsMwXQFxAB7ujJi4hP8BH8j6qkmy4uulU5SSQa6XkElcaKQYSpJcSjkjyTDI
# Opf6LZBTaFx6eeoqDZ0lURhiRqO+1yo8uXO89e6kgBeC8t1WN5ITqXnjocYgDvyF
# pptsUDgnRUiI1M/Ql/O299VktMkIL72i6Qd4BBsrj3Z+iLEnKP9epUwosP1m3N2v
# 9yhXQ1HiusJl63IfXIyfBJaWvQDgU3Jk4eIZSr/2KIj4ptXt496CRiHTi011kcwD
# pdjQLAQiCvoj1puyhfwVf2G5ZwBptIXivNRba34KkD5oqmEoF1yRFQ84iDsf/giy
# n/XIT7YY/zCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUwDQYJKoZI
# hvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# MjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAy
# MDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg4r25Phdg
# M/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1qUoNEt6aORmsHFPPF
# dvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTamDu6
# GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byNpOORj7I5LFGc6XBp
# Dco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL64NF50Zu
# yjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4dPf0gz3N9QZpGdc3E
# XzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0
# lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8QmguEOqEUUbi0b1q
# GFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ
# +QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435UsSFF5PA
# PBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQIDAQABo4IB3TCCAdkw
# EgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTEmr6CkTxG
# NSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMFwGA1UdIARV
# MFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0bTATBgNVHSUEDDAK
# BggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMC
# AYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvX
# zpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYI
# KwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDANBgkqhkiG
# 9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwUtj5OR2R4sQaTlz0x
# M7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th542DYunKmC
# VgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU5HhTdSRXud2f8449
# xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVCs/wM
# nosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGyqVvfSaN0DLzskYDS
# PeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB62FD+CljdQDzHVG2d
# Y3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltEAY5aGZFrDZ+kKNxn
# GSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFpAUR+fKFhbHP+Crvs
# QWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmNcP7ntdAoGokL
# jzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRbatGePu1+oDEzfbzL
# 6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6FwZvKhggNN
# MIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEn
# MCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkE0MDAtMDVFMC1EOTQ3MSUwIwYDVQQD
# ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQD5
# r3DVRpAGQo9sTLUHeBC87NpK+qCBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6RudTTAiGA8yMDIzMTIwNzAxMjQy
# OVoYDzIwMjMxMjA4MDEyNDI5WjB0MDoGCisGAQQBhFkKBAExLDAqMAoCBQDpG51N
# AgEAMAcCAQACAhPiMAcCAQACAhO3MAoCBQDpHO7NAgEAMDYGCisGAQQBhFkKBAIx
# KDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZI
# hvcNAQELBQADggEBAHptbTg1Jtz/WoZ+ydF0/s7FIPWA8PHZGRgQbyaLCjc9BYL8
# UZdEpNkX0N9Rm0Odq8oPZEsF4VBODxvkPYtuYLdZK54CZufIgQSKUfPmsbOujsl9
# 5l8OzZSriSxbf3rf7qHdDDL900SkLN5ggu7OKZSjUUTmB/PfSDs6MTBr3MtjmtIV
# 9cgo63vA4Re3RMvrquVPCULX5CYroBRTMtI+Iifr4vdI8GGM7ub/UxwOWdMyb1D/
# 7ftbG2WWqw75LdD+4gawI9LquDiWajQ9k0XT9EGenF4vpCBBRrnMrAOnDv54tbYP
# 4MidBBi4hH6RpiFlW6QKC2tSIDGoghku+GJVXPAxggQNMIIECQIBATCBkzB8MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdYnaf9yLVbIrgABAAAB1jAN
# BglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8G
# CSqGSIb3DQEJBDEiBCBaXW5la/c4MWoVOjn92wPCjvnSIGG7Lz6ffOwED8cN3TCB
# +gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EINbLTQ1XeNM+EBinOEJMjZd0jMND
# ur+AK+O8P12j5ST8MIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
# MTACEzMAAAHWJ2n/ci1WyK4AAQAAAdYwIgQgfwKv/LY0+Lt3ce5gkywmPdqmdaG5
# Klul+nbii5LMooswDQYJKoZIhvcNAQELBQAEggIATOfPmlxBZ9Ze/kYa/0UC7IUI
# u21XU/GbO683F/PKehdom0fOrMB4gVF45sYL4O6U8gDtYFXnO2wPuR/8W7A1qKBa
# HQmahaxSsaKwJOV3QErIdR/DyBkTxNmmGq9jpZxm6rb28wqMrd/Ly+/fCDMhLUOB
# 1jgLj1mDbk0wMNArIK/08/4Dgm9vNMTwLUxIVIMQ9bZ/i3s7t57KAavpezEtKwPb
# rwu3i002OOQjqYZULEJDYHBbwBUreqVEUeGUr0uZWXr7jxI4r0bG5p/Aa4T6WHqq
# cEO0HfxQMIfmtNvdPZCD7ETt9ZHL/tBEnv5Mq8Xr+7VMX61zRjYfmzvklljryxRc
# QLwEprH7vWcQ8wXmNhg78jhxB5AFBIR9mjNOcgRExmGDAqBZ0kygnqyiHPoKTsMi
# dMTOdmL4lwPAYBvbOXWSzyu7EvfaCHnHp3Acjl5CeZY1qfas0sdHVD4SxZIFLnfG
# yYJhUDG7JVcSeIIBzezMTYTDP7b0rghj+HpqHkGBT4CqWEjbIPe6dmtFMmKAJ7nH
# PXigc9Vv6lm/pMqe82v2iHNDnwADJG8WnThc9CXU8QkS+K+vnuu6BT4P/tmq0YrY
# 8yTuF/7UaBUJQhhcE+RcSBhTIwbyZd9QGUpQrcnefqxl2oCWdejaqtnaVUr01lkg
# XsAIjUnxHVK3CvtzC2A=
# SIG # End signature block
