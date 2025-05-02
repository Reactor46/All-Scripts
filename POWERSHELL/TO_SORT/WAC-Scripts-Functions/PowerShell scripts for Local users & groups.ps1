function Add-WACLGUserToLocalGroups {
<#

.SYNOPSIS
Adds a local or domain user to one or more local groups.

.DESCRIPTION
Adds a local or domain user to one or more local groups. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $UserName,

    [Parameter(Mandatory = $true)]
    [String[]]
    $GroupNames
)
Set-StrictMode -Version 5.0

Import-Module Microsoft.PowerShell.LocalAccounts -ErrorAction SilentlyContinue

$ErrorActionPreference = 'Stop'

# ADSI does NOT support 2016 Nano, meanwhile New-LocalUser, Get-LocalUser, Set-LocalUser do NOT support downlevel
$Error.Clear()
# Get user name or object
$user = $null
$objUser = $null
if (Get-Command 'Get-LocalUser' -errorAction SilentlyContinue) {
    if ($UserName -like '*\*') { # domain user
        $user = $UserName
    } else {
        $user = Get-LocalUser -Name $UserName
    }
} else {
    if ($UserName -like '*\*') { # domain user
        $UserName = $UserName.Replace('\', '/')
    }
    $objUser = "WinNT://$UserName,user"
}
# Add user to groups
Foreach ($name in $GroupNames) {
    if (Get-Command 'Get-LocalGroup' -errorAction SilentlyContinue) {
        $group = Get-LocalGroup $name
        Add-LocalGroupMember -Group $group -Member $user
    }
    else {
        $group = $name
        try {
            $objGroup = [ADSI]("WinNT://localhost/$group,group")
            $objGroup.Add($objUser)
        }
        catch
        {
            # Append user and group name info to error message and then clear it
            $ErrMsg = $_.Exception.Message + " User: " + $UserName + ", Group: " + $group
            Write-Error $ErrMsg
            $Error.Clear()
        }
    }
}

}
## [END] Add-WACLGUserToLocalGroups ##
function Get-WACLGLocalGroupUsers {
<#

.SYNOPSIS
Get users belong to group.

.DESCRIPTION
Get users belong to group. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $group
)

Set-StrictMode -Version 5.0

# ADSI does NOT support 2016 Nano, meanwhile Get-LocalGroupMember does NOT support downlevel and also has bug
$ComputerName = (get-item Env:\COMPUTERNAME).Value
try {
    $groupconnection = [ADSI]("WinNT://localhost/$group,group")
    $contents = $groupconnection.Members() | ForEach-Object {
        $path=$_.GetType().InvokeMember("ADsPath", "GetProperty", $NULL, $_, $NULL)
        # $path will looks like:
        #   WinNT://ComputerName/Administrator
        #   WinNT://DomainName/Domain Admins
        # Find out if this is a local or domain object and trim it accordingly
        if ($path -like "*/$ComputerName/*"){
            $start = 'WinNT://' + $ComputerName + '/'
        }
        else {
            $start = 'WinNT://'
        }
        $name = $path.Substring($start.length)
        $name.Replace('/', '\') #return name here
    }
    return $contents
}
catch { # if above block failed (say in 2016Nano), use another cmdlet
    # clear existing error info from try block
    $Error.Clear()
    #There is a known issue, in some situation Get-LocalGroupMember return: Failed to compare two elements in the array.
    $contents = Get-LocalGroupMember -group $group
    $names = $contents.Name | ForEach-Object {
        $name = $_
        if ($name -like "$ComputerName\*") {
            $name = $name.Substring($ComputerName.length+1)
        }
        $name
    }
    return $names
}

}
## [END] Get-WACLGLocalGroupUsers ##
function Get-WACLGLocalGroups {
<#

.SYNOPSIS
Gets the local groups.

.DESCRIPTION
Gets the local groups. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory = $false)]
    [String]
    $SID
)

Set-StrictMode -Version 5.0

Import-Module Microsoft.PowerShell.LocalAccounts -ErrorAction SilentlyContinue

$isWinServer2016OrNewer = [Environment]::OSVersion.Version.Major -ge 10;
# ADSI does NOT support 2016 Nano, meanwhile New-LocalUser, Get-LocalUser, Set-LocalUser do NOT support downlevel
if ($SID)
{
    if ($isWinServer2016OrNewer)
    {
        Get-LocalGroup -SID $SID | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object Description,
                                          Name,
                                          @{Name="SID"; Expression={$_.SID.Value}},
                                          ObjectClass;
    }
    else
    {
        Get-WmiObject -Class Win32_Group -Filter "LocalAccount='True' AND SID='$SID'" | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object Description, Name, SID, ObjectClass;
    }
}
else
{
    if ($isWinServer2016OrNewer)
    {
        Get-LocalGroup | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object Description,
                                Name,
                                @{Name="SID"; Expression={$_.SID.Value}},
                                ObjectClass;
    }
    else
    {
        Get-WmiObject -Class Win32_Group -Filter "LocalAccount='True'" | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object Description, Name, SID, ObjectClass
    }
}

}
## [END] Get-WACLGLocalGroups ##
function Get-WACLGLocalUserBelongGroups {
<#

.SYNOPSIS
Get a local user belong to group list.

.DESCRIPTION
Get a local user belong to group list. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $UserName
)

Set-StrictMode -Version 5.0

Import-Module CimCmdlets -ErrorAction SilentlyContinue

$operatingSystem = Get-CimInstance Win32_OperatingSystem
$version = [version]$operatingSystem.Version
# product type 3 is server, version number ge 10 is server 2016
$isWinServer2016OrNewer = ($operatingSystem.ProductType -eq 3) -and ($version -ge '10.0')

# ADSI does NOT support 2016 Nano, meanwhile net localgroup do NOT support downlevel "net : System error 1312 has occurred."

# Step 1: get the list of local groups
if ($isWinServer2016OrNewer) {
    $grps = net localgroup | Where-Object {$_ -AND $_ -match "^[*]"}  # group member list as "*%Fws\r\n"
    $groups = $grps.trim('*')
}
else {
    $grps = Get-WmiObject -Class Win32_Group -Filter "LocalAccount='True'" | Microsoft.PowerShell.Utility\Select-Object Name
    $groups = $grps.Name
}

# Step 2: in each group, list members and find match to target $UserName
$groupNames = @()
$regex = '^' + $UserName + '\b'
foreach ($group in $groups) {
    $found = $false
    #find group members
    if ($isWinServer2016OrNewer) {
        $members = net localgroup $group | Where-Object {$_ -AND $_ -notmatch "command completed successfully"} | Microsoft.PowerShell.Utility\Select-Object -skip 4
        if ($members -AND $members.contains($UserName)) {
            $found = $true
        }
    }
    else {
        $groupconnection = [ADSI]("WinNT://localhost/$group,group")
        $members = $groupconnection.Members()
        ForEach ($member in $members) {
            $name = $member.GetType().InvokeMember("Name", "GetProperty", $NULL, $member, $NULL)
            if ($name -AND ($name -match $regex)) {
                $found = $true
                break
            }
        }
    }
    #if members contains $UserName, add group name to list
    if ($found) {
        $groupNames = $groupNames + $group
    }
}
return $groupNames

}
## [END] Get-WACLGLocalUserBelongGroups ##
function Get-WACLGLocalUsers {
<#

.SYNOPSIS
Gets the local users.

.DESCRIPTION
Gets the local users. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory = $false)]
    [String]
    $SID
)

Set-StrictMode -Version 5.0

$isWinServer2016OrNewer = [Environment]::OSVersion.Version.Major -ge 10;
# ADSI does NOT support 2016 Nano, meanwhile New-LocalUser, Get-LocalUser, Set-LocalUser do NOT support downlevel
if ($SID)
{
    if ($isWinServer2016OrNewer)
    {
        Get-LocalUser -SID $SID | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object AccountExpires,
                                         Description,
                                         Enabled,
                                         FullName,
                                         LastLogon,
                                         Name,
                                         ObjectClass,
                                         PasswordChangeableDate,
                                         PasswordExpires,
                                         PasswordLastSet,
                                         PasswordRequired,
                                         @{Name="SID"; Expression={$_.SID.Value}},
                                         UserMayChangePassword;
    }
    else
    {
        Get-WmiObject -Class Win32_UserAccount -Filter "LocalAccount='True' AND SID='$SID'" | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object AccountExpirationDate,
                                                                                      Description,
                                                                                      @{Name="Enabled"; Expression={-not $_.Disabled}},
                                                                                      FullName,
                                                                                      LastLogon,
                                                                                      Name,
                                                                                      ObjectClass,
                                                                                      PasswordChangeableDate,
                                                                                      PasswordExpires,
                                                                                      PasswordLastSet,
                                                                                      PasswordRequired,
                                                                                      SID,
                                                                                      @{Name="UserMayChangePassword"; Expression={$_.PasswordChangeable}}
    }
}
else
{
    if ($isWinServer2016OrNewer)
    {
        Get-LocalUser | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object AccountExpires,
                               Description,
                               Enabled,
                               FullName,
                               LastLogon,
                               Name,
                               ObjectClass,
                               PasswordChangeableDate,
                               PasswordExpires,
                               PasswordLastSet,
                               PasswordRequired,
                               @{Name="SID"; Expression={$_.SID.Value}},
                               UserMayChangePassword;
    }
    else
    {
        Get-WmiObject -Class Win32_UserAccount -Filter "LocalAccount='True'" | Sort-Object -Property Name | Microsoft.PowerShell.Utility\Select-Object AccountExpirationDate,
                                                                                      Description,
                                                                                      @{Name="Enabled"; Expression={-not $_.Disabled}},
                                                                                      FullName,
                                                                                      LastLogon,
                                                                                      Name,
                                                                                      ObjectClass,
                                                                                      PasswordChangeableDate,
                                                                                      PasswordExpires,
                                                                                      PasswordLastSet,
                                                                                      PasswordRequired,
                                                                                      SID,
                                                                                      @{Name="UserMayChangePassword"; Expression={$_.PasswordChangeable}}
    }
}

}
## [END] Get-WACLGLocalUsers ##
function New-WACLGLocalGroup {
<#

.SYNOPSIS
Creates a new local group.

.DESCRIPTION
Creates a new local group. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016 but not Nano.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $GroupName,

    [Parameter(Mandatory = $false)]
    [String]
    $Description
)

Set-StrictMode -Version 5.0

if (-not $Description) {
    $Description = ""
}

# ADSI does NOT support 2016 Nano, meanwhile New-LocalGroup does NOT support downlevel and also with known bug
$Error.Clear()
try {
    $adsiConnection = [ADSI]"WinNT://localhost"
    $group = $adsiConnection.Create("Group", $GroupName)
    $group.InvokeSet("description", $Description)
    $group.SetInfo();
}
catch [System.Management.Automation.RuntimeException]
{ # if above block failed due to no ADSI (say in 2016Nano), use another cmdlet
    if ($_.Exception.Message -ne 'Unable to find type [ADSI].') {
        Write-Error $_.Exception.Message
        return
    }
    # clear existing error info from try block
    $Error.Clear()
    New-LocalGroup -Name $GroupName -Description $Description
}

}
## [END] New-WACLGLocalGroup ##
function New-WACLGLocalUser {
<#

.SYNOPSIS
Creates a new local users.

.DESCRIPTION
Creates a new local users. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016 but not Nano.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $UserName,

    [Parameter(Mandatory = $false)]
    [String]
    $FullName,

    [Parameter(Mandatory = $false)]
    [String]
    $Description,

    [Parameter(Mandatory = $false)]
    [String]
    $Password
)

Set-StrictMode -Version 5.0

if (-not $Description) {
    $Description = ""
}

if (-not $FullName) {
    $FullName = ""
}

if (-not $Password) {
    $Password = ""
}

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function DecryptDataWithJWKOnNode($encryptedData) {
  if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue) {
    $rsaProvider = (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedData), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
  }
  throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
}

# $isWinServer2016OrNewer = [Environment]::OSVersion.Version.Major -ge 10;
# ADSI does NOT support 2016 Nano, meanwhile New-LocalUser does NOT support downlevel and also with known bug
$Error.Clear()
try {
    $adsiConnection = [ADSI]"WinNT://localhost"
    $user = $adsiConnection.Create("User", $UserName)
    if ($Password) {
        $decryptedPassword = DecryptDataWithJWKOnNode $Password
        $user.SetPassword($decryptedPassword)
    }
    $user.InvokeSet("fullName", $FullName)
    $user.InvokeSet("description", $Description)
    $user.SetInfo()
}
catch [System.InvalidOperationException] {
    Write-Error $_.Exception.Message
    return @{
        valid = $false
        reason = "Error"
        data = $_.Exception.Message
    }
}
catch [System.Management.Automation.RuntimeException] {
    # if above block failed due to no ADSI (say in 2016Nano), use another cmdlet
    if ($_.Exception.Message -ne 'Unable to find type [ADSI].') {
        Write-Error $_.Exception.Message
    }
    # clear existing error info from try block
    $Error.Clear()
    if ($Password) {
        #Found a bug where the cmdlet will create a user even if the password is not strong enough
        New-LocalUser -Name $UserName -FullName $FullName -Description $Description -Password $Password;
    }
    else {
        New-LocalUser -Name $UserName -FullName $FullName -Description $Description -NoPassword;
    }
}

}
## [END] New-WACLGLocalUser ##
function Remove-WACLGLocalGroup {
<#

.SYNOPSIS
Delete a local group.

.DESCRIPTION
Delete a local group. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $GroupName
)

Set-StrictMode -Version 5.0

try {
    $adsiConnection = [ADSI]"WinNT://localhost";
    $adsiConnection.Delete("Group", $GroupName);
}
catch {
    # Instead of _.Exception.Message, InnerException.Message is more meaningful to end user
    Write-Error $_.Exception.InnerException.Message
    $Error.Clear()
}

}
## [END] Remove-WACLGLocalGroup ##
function Remove-WACLGLocalUser {
<#

.SYNOPSIS
Delete a local user.

.DESCRIPTION
Delete a local user. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $UserName
)

Set-StrictMode -Version 5.0

try {
    $adsiConnection = [ADSI]"WinNT://localhost";
    $adsiConnection.Delete("User", $UserName);
}
catch {
    # Instead of _.Exception.Message, InnerException.Message is more meaningful to end user
    Write-Error $_.Exception.InnerException.Message
    $Error.Clear()
}

}
## [END] Remove-WACLGLocalUser ##
function Remove-WACLGLocalUserFromLocalGroups {
<#

.SYNOPSIS
Removes a local user from one or more local groups.

.DESCRIPTION
Removes a local user from one or more local groups. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $UserName,

    [Parameter(Mandatory = $true)]
    [String[]]
    $GroupNames
)

Set-StrictMode -Version 5.0

$isWinServer2016OrNewer = [Environment]::OSVersion.Version.Major -ge 10;
# ADSI does NOT support 2016 Nano, meanwhile net localgroup do NOT support downlevel "net : System error 1312 has occurred."
$Error.Clear()
$message = ""
$results = @()
if (!$isWinServer2016OrNewer) {
    $objUser = "WinNT://$UserName,user"
}
Foreach ($group in $GroupNames) {
    if ($isWinServer2016OrNewer) {
        # If execute an external command, the following steps to be done to product correct format errors:
        # -	Use "2>&1" to store the error to the variable.
        # -	Watch $Error.Count to determine the execution result.
        # -	Concatinate the error message to single string and sprit out with Write-Error.
        $Error.Clear()
        $result = & 'net' localgroup $group $UserName /delete 2>&1
        # $LASTEXITCODE here does not return error code, have to use $Error
        if ($Error.Count -ne 0) {
            foreach($item in $result) {
                if ($item.Exception.Message.Length -gt 0) {
                    $message += $item.Exception.Message
                }
            }
            $Error.Clear()
            Write-Error $message
        }
    }
    else {
        $objGroup = [ADSI]("WinNT://localhost/$group,group")
        $objGroup.Remove($objUser)
    }
}

}
## [END] Remove-WACLGLocalUserFromLocalGroups ##
function Remove-WACLGUsersFromLocalGroup {
<#

.SYNOPSIS
Removes local or domain users from the local group.

.DESCRIPTION
Removes local or domain users from the local group. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [String[]]
    $users,

    [Parameter(Mandatory = $true)]
    [String]
    $group
)

Set-StrictMode -Version 5.0

$isWinServer2016OrNewer = [Environment]::OSVersion.Version.Major -ge 10;
# ADSI does NOT support 2016 Nano, meanwhile net localgroup do NOT support downlevel "net : System error 1312 has occurred."

$message = ""
Foreach ($user in $users) {
    if ($isWinServer2016OrNewer) {
        # If execute an external command, the following steps to be done to product correct format errors:
        # -	Use "2>&1" to store the error to the variable.
        # -	Watch $Error.Count to determine the execution result.
        # -	Concatinate the error message to single string and sprit out with Write-Error.
        $Error.Clear()
        $result = & 'net' localgroup $group $user /delete 2>&1
        # $LASTEXITCODE here does not return error code, have to use $Error
        if ($Error.Count -ne 0) {
            foreach($item in $result) {
                if ($item.Exception.Message.Length -gt 0) {
                    $message += $item.Exception.Message
                }
            }
            $Error.Clear()
            Write-Error $message
        }
    }
    else {
        if ($user -like '*\*') { # domain user
            $user = $user.Replace('\', '/')
        }
        $groupInstance = [ADSI]("WinNT://localhost/$group,group")
        $groupInstance.Remove("WinNT://$user,user")
    }
}

}
## [END] Remove-WACLGUsersFromLocalGroup ##
function Rename-WACLGLocalGroup {
 <#

 .SYNOPSIS
Renames a local group.

.DESCRIPTION
Renames a local group. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016 but not Nano.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $GroupName,

    [Parameter(Mandatory = $true)]
    [String]
    $NewGroupName
)

Set-StrictMode -Version 5.0

# ADSI does NOT support 2016 Nano, meanwhile Rename-LocalGroup does NOT support downlevel and also with known bug
$Error.Clear()
try {
    $adsiConnection = [ADSI]"WinNT://localhost"
    $group = $adsiConnection.Children.Find($GroupName, "Group")
    if ($group) {
        $group.psbase.rename($NewGroupName)
        $group.psbase.CommitChanges()
    }
}
catch [System.Management.Automation.RuntimeException]
{ # if above block failed due to no ADSI (say in 2016Nano), use another cmdlet
    if ($_.Exception.Message -ne 'Unable to find type [ADSI].') {
        Write-Error $_.Exception.Message
        return
    }
    # clear existing error info from try block
    $Error.Clear()
    Rename-LocalGroup -Name $GroupName -NewGroupName $NewGroupName
}

}
## [END] Rename-WACLGLocalGroup ##
function Set-WACLGLocalGroupProperties {
<#

.SYNOPSIS
Set local group properties.

.DESCRIPTION
Set local group properties. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $GroupName,

    [Parameter(Mandatory = $false)]
    [String]
    $Description
)

Set-StrictMode -Version 5.0

try {
    $group = [ADSI]("WinNT://localhost/$GroupName, group")

    if ($Description -ne $null) {
        $group.Description = $Description
    }
    
    $group.SetInfo()
}
catch [System.Management.Automation.RuntimeException]
{
     Write-Error $_.Exception.Message
}

return $true

}
## [END] Set-WACLGLocalGroupProperties ##
function Set-WACLGLocalUserPassword {
<#

.SYNOPSIS
Set local user password.

.DESCRIPTION
Set local user password. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $UserName,

    [Parameter(Mandatory = $false)]
    [String]
    $Password
)

Set-StrictMode -Version 5.0

if (-not $Password)
{
    $decryptedPassword = ""
}

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function DecryptDataWithJWKOnNode($encryptedData) {
  if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue) {
    $rsaProvider = (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedData), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
  }
  throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
}

try {
    $decryptedPassword = DecryptDataWithJWKOnNode $Password
    $user = [ADSI]("WinNT://localhost/$UserName, user")
    $user.psbase.invoke("SetPassword", "$decryptedPassword")
}
catch [System.InvalidOperationException] {
    Write-Error $_.Exception.Message
    return @{
        valid = $false
        reason = "Error"
        data = $_.Exception.Message
    }
}

}
## [END] Set-WACLGLocalUserPassword ##
function Set-WACLGLocalUserProperties {
<#

.SYNOPSIS
Set local user properties.

.DESCRIPTION
Set local user properties. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $UserName,

    [Parameter(Mandatory = $false)]
    [String]
    $FullName,

    [Parameter(Mandatory = $false)]
    [String]
    $Description
)

Set-StrictMode -Version 5.0

$user = [ADSI]("WinNT://localhost/$UserName, user")

if ($Description -ne $null) {
    $user.Description = $Description
}

if ($FullName -ne $null) {
    $user.FullName = $FullName
}

$user.SetInfo()

return $true

}
## [END] Set-WACLGLocalUserProperties ##
function Test-WACLGPasswordPolicy {
<#

.SYNOPSIS
Test given password.

.DESCRIPTION
Tests given password against system's password policies to see if it is valid.

.ROLE
Readers

#>

param (
    [Parameter(Mandatory = $false)]
    [String]
    $UserName,

    [Parameter(Mandatory = $false)]
    [String]
    $FullName,

    [Parameter(Mandatory = $false)]
    [String]
    $Password
)

Set-StrictMode -Version 5.0

Import-Module Microsoft.PowerShell.Management -ErrorAction SilentlyContinue

###############################################################################
# Constants
###############################################################################
Set-Variable -Name LogName -Option Constant -Value "Microsoft-ServerManagementExperience" -ErrorAction SilentlyContinue
Set-Variable -Name LogSource -Option Constant -Value "SMEScripts" -ErrorAction SilentlyContinue

function writeInfoLog($logMessage) {
    $message = "[Check-PasswordPolicy]: " + $logMessage
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
        -Message $message -ErrorAction SilentlyContinue
}

function writeErrorLog($errorMessage) {
    $message = "[Check-PasswordPolicy]: " + $logMessage
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
        -Message $message -ErrorAction SilentlyContinue
}

function GetMinimumPasswordLength() {
    $CustomCode = @"
namespace SME
{
    using System;
    using System.Runtime.InteropServices;

    public static class UserModalsInfo
    {
        [DllImport("netapi32.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, SetLastError = true)]
        static extern uint NetUserModalsGet(
            string server,
            int level,
            out IntPtr BufPtr);

        [DllImport("netapi32.dll", SetLastError = true)]
        static extern int NetApiBufferFree(IntPtr Buffer);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct USER_MODALS_INFO_0
        {
            public uint usrmod0_min_passwd_len;
            public uint usrmod0_max_passwd_age;
            public uint usrmod0_min_passwd_age;
            public uint usrmod0_force_logoff;
            public uint usrmod0_password_hist_len;
        };

        public static uint GetMinimumPasswordLength()
        {
            uint passwordLength = 0;
            USER_MODALS_INFO_0 objUserModalsInfo0 = new USER_MODALS_INFO_0();
            IntPtr bufPtr;
            uint exitCode = NetUserModalsGet("\\\\.", 0, out bufPtr);
            if (exitCode != 0)
            {
                throw new InvalidOperationException("Couldn't get MinimumPasswordLength.");
            }

            objUserModalsInfo0 = (USER_MODALS_INFO_0)Marshal.PtrToStructure(bufPtr, typeof(USER_MODALS_INFO_0));
            passwordLength = objUserModalsInfo0.usrmod0_min_passwd_len;
            NetApiBufferFree(bufPtr);
            bufPtr = IntPtr.Zero;

            return passwordLength;
        }
    }
}
"@

    try {
        Add-Type -TypeDefinition $CustomCode
        Remove-Variable CustomCode
        $minimumLength = [SME.UserModalsInfo]::GetMinimumPasswordLength()
    } catch [System.InvalidOperationException] {
        $err = $_.Exception.Message
        $message = "Error occured attempting to get minimum password length policy from system. Error: " + $err
        writeErrorLog $message

        return @{
            valid = $false
            reason = "Error"
            data = $err
        }
    }

    writeInfoLog "Successfully retrieved password policy of length " + $minimumLength + " from system."

    return $minimumLength
}

function CheckPasswordLength($password) {
    $minimumLength = GetMinimumPasswordLength

    if ($password.Length -lt $minimumLength) {
        return @{
            valid = $false
            reason = "LessThanMinimumLength"
            data = $minimumLength
        }
    }

    return @{
        valid = $true
        reason = "Valid"
    }
}

function InitializePositiveRegexes() {
    # Documentation on Windows Password Policy from which these regexes were derived:
    # https://docs.microsoft.com/windows/security/threat-protection/security-policy-settings/password-policy
    $lowerCaseRegex = "[a-z]"
    $upperCaseRegex = "[A-Z]"
    $numericalRegex = "[0-9]"
    $symbolsRegex = "[~!@#\$%\^&\*_\-\+=`\|\\\(\){}\[\]:;`"`'<>,\.?\/]"
    $remainingUnicodeRegex = "[^a-zA-Z0-9~!@#\$%\^&\*_\-\+=`\|\\\(\){}\[\]:;`"`'<>,\.?\/]"

    writeInfoLog "Initialized positive regexes."

    return @($lowerCaseRegex, $upperCaseRegex, $numericalRegex, $symbolsRegex, $remainingUnicodeRegex)
}

function BuildNegativeRegex($userName, $fullName) {
    $negativeRegex = $null
    if ($null -ne $userName -and $userName.Length -ge 3) {
        $negativeRegex = "(" + $userName + ")"
    } else {
        writeInfoLog "No user name provided or length is < 3 characters, not including in negative regex."
    }

    if ($null -ne $fullName) {
        $delimiterRegex = "[,\.\-_\s#]"
        $splitName = $fullName -split $delimiterRegex

        foreach ($name in $splitName) {
            if ($name.Length -ge 3) {
                if ($null -eq $negativeRegex) {
                    $negativeRegex = "(" + $name + ")"
                } else {
                    $negativeRegex = $negativeRegex + "|(" + $name + ")"
                }
            }
        }
    } else {
        writeInfoLog "No full name provided, not including in negative regex."
    }

    writeInfoLog "Built negative regex: $negativeRegex"

    return $negativeRegex
}

function CheckNegativeRegex($password, $negativeRegex) {
    if ($null -ne $negativeRegex -and $password -match $negativeRegex) {
        writeInfoLog "Password failed negative regex check."
        return @{
            valid = $false
            reason = "ContainsName"
        }
    }

    writeInfoLog "Password passed negative regex check."
    return @{
        valid = $true
        reason = "Valid"
    }
}

function CheckPositiveRegexes($password) {
    $positiveRegexes = InitializePositiveRegexes

    $count = 0
    foreach ($regex in $positiveRegexes) {
        if ($password -cmatch $regex) {
            $count = $count + 1
        }
    }

    if ($count -lt 3) {
        writeInfoLog "Password failed positive regex check."
        return @{
            valid = $false
            reason = "UnvariedCharacters"
        }
    }

    writeInfoLog "Password passed positive regex check."
    return @{
        valid = $true
        reason = "Valid"
    }
}

function CheckRegexes($password, $userName, $fullName) {
    $negativeRegex = BuildNegativeRegex $userName $fullName

    $negativeRegexResult = CheckNegativeRegex $password $negativeRegex
    if ($negativeRegexResult.valid -eq $false) {
        return $negativeRegexResult
    }

    $positiveRegexResult = CheckPositiveRegexes $password
    if ($positiveRegexResult.valid -eq $false) {
        return $positiveRegexResult
    }

    return @{
        valid = $true
        reason = "Valid"
    }
}

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function DecryptDataWithJWKOnNode($encryptedData) {
  if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue) {
    $rsaProvider = (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedData), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
  }
  throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
}

function Main($password, $userName, $fullName) {
    if ($null -eq $password) {
        writeInfoLog "Password was not provided, returning invalid."
        return @{
            valid = $false
            reason = "NotProvided"
        }
    }
    try {
        $decryptedPassword = DecryptDataWithJWKOnNode $password
        $minimumLengthResult = CheckPasswordLength $decryptedPassword
        if ($minimumLengthResult.valid -eq $false) {
            return $minimumLengthResult
        }

        return CheckRegexes $decryptedPassword $userName $fullName
    }
    catch [System.InvalidOperationException] {
        Write-Error $_.Exception.Message
        return @{
            valid = $false
            reason = "Error"
            data = $_.Exception.Message
        }
    }
}

###############################################################################
# Script execution starts here...
###############################################################################
if (-not ($env:pester)) {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    return Main $Password $UserName $FullName
}

}
## [END] Test-WACLGPasswordPolicy ##
function Get-WACLGCimWin32LogicalDisk {
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
## [END] Get-WACLGCimWin32LogicalDisk ##
function Get-WACLGCimWin32NetworkAdapter {
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
## [END] Get-WACLGCimWin32NetworkAdapter ##
function Get-WACLGCimWin32PhysicalMemory {
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
## [END] Get-WACLGCimWin32PhysicalMemory ##
function Get-WACLGCimWin32Processor {
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
## [END] Get-WACLGCimWin32Processor ##
function Get-WACLGClusterInventory {
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
## [END] Get-WACLGClusterInventory ##
function Get-WACLGClusterNodes {
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
## [END] Get-WACLGClusterNodes ##
function Get-WACLGDecryptedDataFromNode {
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
## [END] Get-WACLGDecryptedDataFromNode ##
function Get-WACLGEncryptionJWKOnNode {
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
## [END] Get-WACLGEncryptionJWKOnNode ##
function Get-WACLGServerInventory {
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
## [END] Get-WACLGServerInventory ##

# SIG # Begin signature block
# MIIoKgYJKoZIhvcNAQcCoIIoGzCCKBcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDr1wjNNDRxMFje
# o26j9X08D1OwbXFwsGYAkIpR+54kPKCCDXYwggX0MIID3KADAgECAhMzAAADTrU8
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
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIDXEsjJhAa4aCq6AVzzc6oE+
# swumfe9wQU7wzYZvtp35MEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEAvNqVRLUnI6dRLVUJSKOFQRDVCPAkidRG8U+JQUjw+7YHSaYvXsrM9qMJ
# 06g8+VqY1msA0lJSxHCcxFIpY9BoL5OLdBwvwU9KhreQSSq/dWjopFhJ0Geod4B/
# mEddZ4p9zN12OT0NiSuvgVn1RVOb+KQ0qQTBFm6gy+XX8Rn/JF6QZcxzBSh8c7rS
# EyRYU6HY2B/Le0XzuK3VlefmdEQnCTDJDkyM09uX0P2Ta68LgftYs4g3VtL40Ctg
# DxzuDjPVWIzSgmwmYjU9eQzrI+avwXCsLlF9oidYC9o1PiLW+FCz+iK7ExdiIO1L
# ocajYxP1HQy86CmhxnRROV2oxSGBYaGCF5QwgheQBgorBgEEAYI3AwMBMYIXgDCC
# F3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsq
# hkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCBdQ7WkNY+Zc6vw+b7GfFDXPCBgXxKzSbv75rR0pwSy4wIGZVbCx7iy
# GBMyMDIzMTIwNjIzNDEzNC42ODNaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQGEwJV
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
# IFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6Rr0jTAiGA8yMDIzMTIwNjEzMjQy
# OVoYDzIwMjMxMjA3MTMyNDI5WjB0MDoGCisGAQQBhFkKBAExLDAqMAoCBQDpGvSN
# AgEAMAcCAQACAiW7MAcCAQACAhPwMAoCBQDpHEYNAgEAMDYGCisGAQQBhFkKBAIx
# KDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZI
# hvcNAQELBQADggEBAGmcd8NPOQupwLSgzOFTxGWCQElpCUMjxBxxGIdhuYhdVDA9
# zaOLTMo0FGyyRUOyPOvIR7gWAjXUxTYb1mCuFjHLl6D+3EEotxGpCJ6Q6dtNtA/C
# M2t4RfxbUt3gLBlsFqXcNNw6Xsv6R/XeMqzp1c5nK02bXKNebNWyjyy86Ts8QtsI
# BklFGRmhs5rj+jQixrdVj/wGnBfj0P/d3kCnqpdBvR/RadGUjao74+CZ3yTV4Put
# 3OKyKbIN156Mbyd+HbXkdIrljyXdoNeafEtaCHnYU0uvP0jvsE+uIGTtiMnRbj8R
# 3OGbreh/VNQ/aGOxtn1f5fyKfkeY9UjUx+8n94MxggQNMIIECQIBATCBkzB8MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdYnaf9yLVbIrgABAAAB1jAN
# BglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8G
# CSqGSIb3DQEJBDEiBCAfrl2/lSVAcjr5PvlyNgI98ebhGOSW0gA9iQ9i67ShxjCB
# +gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EINbLTQ1XeNM+EBinOEJMjZd0jMND
# ur+AK+O8P12j5ST8MIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
# MTACEzMAAAHWJ2n/ci1WyK4AAQAAAdYwIgQgQLOvfvDOqqPid9EWY9iO5XSoScUy
# 2WrTytwAI+yxXt4wDQYJKoZIhvcNAQELBQAEggIAsYHSTHaBVj9dxd6K9Xo00ZVt
# /Kq7fvbLYxDJN24P9usZTIEtJGTlTh7Lt+dDJO2JE+K4eP8fnDSpFAi632Y3XfVe
# Jv13TNHghKq8wLd9DjTS1G352N3fj0u9ypfkC1keaxdE+gpoY6PoTr/9pr2dKP9h
# MnYimxlzMf3fyBMRGHY+/AGr3Bfg4i3Oktd4NLIqqbkPzfGTI6ZmOcyzV6Ldn9iq
# WE+1gWqID/JXckMhZQ97riHzAAKMC8Cg7UQczC4bJ9wgeyGxOoQ3nBNkKk3et5ih
# N2nqh7U3z73na96+uhlDYj74WjC3vU0OaVXitMM/9xxuftcfUjVSHYv0KwWFF/H+
# T1GyicFD0QnFCviQYZXkAmnp0dduyb+RPhdeLPgP40eBZMh9wy1lwDNEwp9fVASW
# pQ/RG/hA/7XyU82uBXSCi43WQh/jwPwNIQtlQkl02lMSDTo+5WmirpkWDOujgd/O
# 5U656T69/2woiZBv5vUwl/eW92FlDrF6XuyQ0sHhTHwsBZXsHQftH+2s/r6sotbi
# fFwIHqrAGP5AszINsG3sv7UViWuureCZ09DQnyrmLjqEHdYs2O7wXNhuzv93gsaD
# x/h38Jg2uUmuDFymd98WgdOeul576W99IV2tJUyVvuEkdv0TBSN6xf366rQ7mErL
# 2cDKA18Pjfios4ztvpI=
# SIG # End signature block
