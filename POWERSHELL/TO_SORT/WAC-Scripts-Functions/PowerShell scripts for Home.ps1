function Disable-WACSHCredSspClientRole {
<#

.SYNOPSIS
Disables CredSSP on this client/gateway.

.DESCRIPTION
Disables CredSSP on this client/gateway.

.ROLE
Administrators

.Notes
The feature(s) that use this script are still in development and should be considered as being "In Preview".
Therefore, those feature(s) and/or this script may change at any time.

#>

Set-StrictMode -Version 5.0
Import-Module  Microsoft.WSMan.Management -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Setup all necessary global variables, constants, etc.

.DESCRIPTION
Setup all necessary global variables, constants, etc.

#>

function setupScriptEnv() {
    Set-Variable -Name WsManApplication -Option ReadOnly -Scope Script -Value "wsman"
    Set-Variable -Name CredSSPClientAuthPath -Option ReadOnly -Scope Script -Value "localhost\Client\Auth\CredSSP"
}

<#

.SYNOPSIS
Clean up all added global variables, constants, etc.

.DESCRIPTION
Clean up all added global variables, constants, etc.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name WsManApplication -Scope Script -Force
    Remove-Variable -Name CredSSPClientAuthPath -Scope Script -Force
}

<#

.SYNOPSIS
Is CredSSP client role enabled on this server.

.DESCRIPTION
When the CredSSP client role is enabled on this server then return $true.

#>

function getCredSSPClientEnabled() {
    $path = "{0}:\{1}" -f $WsManApplication, $CredSSPClientAuthPath

    $credSSPClientEnabled = $false;

    $credSSPClientService = Get-Item $path -ErrorAction SilentlyContinue
    if ($credSSPClientService) {
        $credSSPClientEnabled = [System.Convert]::ToBoolean($credSSPClientService.Value)
    }

    return $credSSPClientEnabled
}

<#

.SYNOPSIS
Disable CredSSP

.DESCRIPTION
Attempt to disable the CredSSP Client role and return any error that occurs

#>

function disableCredSSP() {
    $err = $null

    # Catching the result so that we can discard it. Otherwise it get concatinated with $err and we don't want that!
    $result = Disable-WSManCredSSP -Role Client -ErrorAction SilentlyContinue -ErrorVariable +err

    return $err
}

<#

.SYNOPSIS
Main function.

.DESCRIPTION
Main function.

#>

function main() {
    setupScriptEnv

    try {
        # Only continue if the client role is not disabled.
        if (getCredSSPClientEnabled) {
            $err = disableCredSSP

            if ($err) {
                # Only throw error if there was an error and the client role is still enabled.
                if (getCredSSPClientEnabled) {
                    throw @($err)[0]
                }
            }
        }
    } catch {
        Write-Error $_.Exception.Message

        throw $_.Exception.Message
    } finally {
        cleanupScriptEnv
    }

    return $true
}

###############################################################################
# SCcript execution starts here...
###############################################################################
$results = $null

if (-not ($env:pester)) {
    $results = main
}


return $results
    
}
## [END] Disable-WACSHCredSspClientRole ##
function Disable-WACSHCredSspManagedServer {
<#

.SYNOPSIS
Disables CredSSP on this server.

.DESCRIPTION
Disables CredSSP on this server.

.ROLE
Administrators

.Notes
The feature(s) that use this script are still in development and should be considered as being "In Preview".
Therefore, those feature(s) and/or this script may change at any time.

#>

Set-StrictMode -Version 5.0
Import-Module  Microsoft.WSMan.Management -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Is CredSSP client role enabled on this server.

.DESCRIPTION
When the CredSSP client role is enabled on this server then return $true.

#>

function getCredSSPClientEnabled() {
    Set-Variable credSSPClientPath -Option Constant -Value "WSMan:\localhost\Client\Auth\CredSSP" -ErrorAction SilentlyContinue

    $credSSPClientEnabled = $false;

    $credSSPClientService = Get-Item $credSSPClientPath -ErrorAction SilentlyContinue
    if ($credSSPClientService) {
        $credSSPClientEnabled = [System.Convert]::ToBoolean($credSSPClientService.Value)
    }

    return $credSSPClientEnabled
}

<#

.SYNOPSIS
Disable CredSSP

.DESCRIPTION
Attempt to disable the CredSSP Client role and return any error that occurs

#>

function disableCredSSPClientRole() {
    $err = $null

    # Catching the result so that we can discard it. Otherwise it get concatinated with $err and we don't want that!
    $result = Disable-WSManCredSSP -Role Client -ErrorAction SilentlyContinue -ErrorVariable +err

    return $err
}

<#

.SYNOPSIS
Disable the CredSSP client role on this server.

.DESCRIPTION
Disable the CredSSP client role on this server.

#>

function disableCredSSPClient() {
    # If disabled then we can stop.
    if (-not (getCredSSPClientEnabled)) {
        return $null
    }

    $err = disableCredSSPClientRole

    # If there is an error and it is not enabled, then success
    if ($err) {
        if (-not (getCredSSPClientEnabled)) {
            return $null
        }

        return $err
    }

    return $null
}

<#

.SYNOPSIS
Is CredSSP server role enabled on this server.

.DESCRIPTION
When the CredSSP server role is enabled on this server then return $true.

#>

function getCredSSPServerEnabled() {
    Set-Variable credSSPServicePath -Option Constant -Value "WSMan:\localhost\Service\Auth\CredSSP" -ErrorAction SilentlyContinue

    $credSSPServerEnabled = $false;

    $credSSPServerService = Get-Item $credSSPServicePath -ErrorAction SilentlyContinue
    if ($credSSPServerService) {
        $credSSPServerEnabled = [System.Convert]::ToBoolean($credSSPServerService.Value)
    }

    return $credSSPServerEnabled
}

<#

.SYNOPSIS
Disable CredSSP

.DESCRIPTION
Attempt to disable the CredSSP Server role and return any error that occurs

#>

function disableCredSSPServerRole() {
    $err = $null

    # Catching the result so that we can discard it. Otherwise it get concatinated with $err and we don't want that!
    $result = Disable-WSManCredSSP -Role Server -ErrorAction SilentlyContinue -ErrorVariable +err

    return $err
}

function disableCredSSPServer() {
    # If not enabled then we can leave
    if (-not (getCredSSPServerEnabled)) {
        return $null
    }

    $err = disableCredSSPServerRole

    # If there is an error, but the requested functionality completed don't fail the operation.
    if ($err) {
        if (-not (getCredSSPServerEnabled)) {
            return $null
        }

        return $err
    }
    
    return $null
}

<#

.SYNOPSIS
Main function.

.DESCRIPTION
Main function.

#>

function main() {
    $err = disableCredSSPServer
    if ($err) {
        throw $err
    }

    $err = disableCredSSPClient
    if ($err) {
        throw $err
    }

    return $true
}

###############################################################################
# Script execution starts here...
###############################################################################

if (-not ($env:pester)) {
    return main
}

}
## [END] Disable-WACSHCredSspManagedServer ##
function Enable-WACSHCredSSPClientRole {
<#

.SYNOPSIS
Enables CredSSP on this computer as client role to the other computer.

.DESCRIPTION
Enables CredSSP on this computer as client role to the other computer.

.ROLE
Administrators

.PARAMETER serverNames
The names of the server to which this gateway can forward credentials.

.LINK
https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/CVE-2018-0886

.LINK
https://aka.ms/CredSSP-Updates

#>

param (
    [Parameter(Mandatory=$True)]
    [string[]]$serverNames
)

Set-StrictMode -Version 5.0
Import-Module  Microsoft.WSMan.Management -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Setup all necessary global variables, constants, etc.

.DESCRIPTION
Setup all necessary global variables, constants, etc.

#>

function setupScriptEnv() {
    Set-Variable -Name WsManApplication -Option ReadOnly -Scope Script -Value "wsman"
    Set-Variable -Name CredSSPClientAuthPath -Option ReadOnly -Scope Script -Value "localhost\Client\Auth\CredSSP"
    Set-Variable -Name CredentialsDelegationPolicyPath -Option ReadOnly -Scope Script -Value "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CredentialsDelegation"
    Set-Variable -Name AllowFreshCredentialsPath -Option ReadOnly -Scope Script -Value "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CredentialsDelegation\AllowFreshCredentials"
    Set-Variable -Name AllowFreshCredentialsWhenNTLMOnlyPath -Option ReadOnly -Scope Script -Value "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CredentialsDelegation\AllowFreshCredentialsWhenNTLMOnly"
    Set-Variable -Name AllowFreshCredentialsPropertyName -Option ReadOnly -Scope Script -Value "AllowFreshCredentials"
    Set-Variable -Name AllowFreshCredentialsWhenNTLMOnlyPropertyName -Option ReadOnly -Scope Script -Value "AllowFreshCredentialsWhenNTLMOnly"
    Set-Variable -Name TypeAlreadyExistsHResult -Option ReadOnly -Scope Script -Value -2146233088
    Set-Variable -Name NativeCode -Option ReadOnly -Scope Script -Value @"
    using Microsoft.Win32;
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading;
    
    namespace SME
    {
        public static class LocalGroupPolicy
        {
            [Guid("EA502722-A23D-11d1-A7D3-0000F87571E3")]
            [ComImport]
            [ClassInterface(ClassInterfaceType.None)]
            public class GPClass
            {
            }
    
            [ComImport, Guid("EA502723-A23D-11d1-A7D3-0000F87571E3"),
            InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IGroupPolicyObject
            {
                void New(
                    [MarshalAs(UnmanagedType.LPWStr)] string pszDomainName,
                    [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName,
                    uint dwFlags);
    
                void OpenDSGPO(
                    [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
                    uint dwFlags);
    
                void OpenLocalMachineGPO(uint dwFlags);
    
                void OpenRemoteMachineGPO(
                    [MarshalAs(UnmanagedType.LPWStr)] string pszComputerName,
                    uint dwFlags);
    
                void Save(
                    [MarshalAs(UnmanagedType.Bool)] bool bMachine,
                    [MarshalAs(UnmanagedType.Bool)] bool bAdd,
                    [MarshalAs(UnmanagedType.LPStruct)] Guid pGuidExtension,
                    [MarshalAs(UnmanagedType.LPStruct)] Guid pGuid);
    
                void Delete();
    
                void GetName(
                    [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName,
                    int cchMaxLength);
    
                void GetDisplayName(
                    [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName,
                    int cchMaxLength);
    
                void SetDisplayName(
                    [MarshalAs(UnmanagedType.LPWStr)] string pszName);
    
                void GetPath(
                    [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszPath,
                    int cchMaxPath);
    
                void GetDSPath(
                    uint dwSection,
                    [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszPath,
                    int cchMaxPath);
    
                void GetFileSysPath(
                    uint dwSection,
                    [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszPath,
                    int cchMaxPath);
    
                IntPtr GetRegistryKey(uint dwSection);
    
                uint GetOptions();
    
                void SetOptions(uint dwOptions, uint dwMask);
    
                void GetMachineName(
                    [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName,
                    int cchMaxLength);
    
                uint GetPropertySheetPages(out IntPtr hPages);
            }
    
            private const int GPO_OPEN_LOAD_REGISTRY = 1;
            private const int GPO_SECTION_MACHINE = 2;
            private const string ApplicationName = @"wsman";
            private const string AllowFreshCredentials = @"AllowFreshCredentials";
            private const string AllowFreshCredentialsWhenNTLMOnly = @"AllowFreshCredentialsWhenNTLMOnly";
            private const string ConcatenateDefaultsAllowFresh = @"ConcatenateDefaults_AllowFresh";
            private const string PathCredentialsDelegationPath = @"SOFTWARE\Policies\Microsoft\Windows";
            private const string GPOpath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy Objects";
            private const string Machine = @"Machine";
            private const string CredentialsDelegation = @"\CredentialsDelegation";
            private const string PoliciesPath = @"Software\Policies\Microsoft\Windows";
            private const string BackSlash = @"\";
    
            public static string EnableAllowFreshCredentialsPolicy(string[] serverNames, bool hasNTLMOnlyUpdateBeenMade)
            {
                if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                {
                    return EnableAllowFreshCredentialsPolicyImpl(serverNames, hasNTLMOnlyUpdateBeenMade);
                }
                else
                {
                    string value = null;
    
                    var thread = new Thread(() =>
                    {
                        value = EnableAllowFreshCredentialsPolicyImpl(serverNames, hasNTLMOnlyUpdateBeenMade);
                    });
    
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();
    
                    return value;
                }
            }
    
            public static string RemoveServersFromAllowFreshCredentialsPolicy(string[] serverNames)
            {
                if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                {
                    return RemoveServersFromAllowFreshCredentialsPolicyImpl(serverNames);
                }
                else
                {
                    string value = null;
    
                    var thread = new Thread(() =>
                    {
                        value = RemoveServersFromAllowFreshCredentialsPolicyImpl(serverNames);
                    });
    
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();
    
                    return value;
                }
            }
    
            private static string EnableAllowFreshCredentialsPolicyImpl(string[] serverNames, bool hasNTLMOnlyUpdateBeenMade)
            {
                IGroupPolicyObject gpo = (IGroupPolicyObject)new GPClass();
                gpo.OpenLocalMachineGPO(GPO_OPEN_LOAD_REGISTRY);
    
                var KeyHandle = gpo.GetRegistryKey(GPO_SECTION_MACHINE);
    
                try
                {
                    var rootKey = Registry.CurrentUser;
    
                    using (RegistryKey GPOKey = rootKey.OpenSubKey(GPOpath, true))
                    {
                        foreach (var keyName in GPOKey.GetSubKeyNames())
                        {
                            if (keyName.EndsWith(Machine, StringComparison.OrdinalIgnoreCase))
                            {
                                var key = GPOpath + BackSlash + keyName + BackSlash + PoliciesPath;
    
                                UpdateGpoRegistrySettingsAllowFreshCredentials(ApplicationName, serverNames, Registry.CurrentUser, key, hasNTLMOnlyUpdateBeenMade);
                            }
                        }
                    }
    
                    //saving gpo settings
                    gpo.Save(true, true, new Guid("35378EAC-683F-11D2-A89A-00C04FBBCFA2"), new Guid("7A9206BD-33AF-47af-B832-D4128730E990"));
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
                finally
                {
                    KeyHandle = IntPtr.Zero;
                }
    
                return null;
            }
    
            private static string RemoveServersFromAllowFreshCredentialsPolicyImpl(string[] serverNames)
            {
                IGroupPolicyObject gpo = (IGroupPolicyObject)new GPClass();
                gpo.OpenLocalMachineGPO(GPO_OPEN_LOAD_REGISTRY);
    
                var KeyHandle = gpo.GetRegistryKey(GPO_SECTION_MACHINE);
    
                try
                {
                    var rootKey = Registry.CurrentUser;
    
                    using (RegistryKey GPOKey = rootKey.OpenSubKey(GPOpath, true))
                    {
                        foreach (var keyName in GPOKey.GetSubKeyNames())
                        {
                            if (keyName.EndsWith(Machine, StringComparison.OrdinalIgnoreCase))
                            {
                                var key = GPOpath + BackSlash + keyName + BackSlash + PoliciesPath;
    
                                UpdateGpoRegistrySettingsRemoveServersFromFreshCredentials(ApplicationName, serverNames, Registry.CurrentUser, key);
                            }
                        }
                    }
    
                    //saving gpo settings
                    gpo.Save(true, true, new Guid("35378EAC-683F-11D2-A89A-00C04FBBCFA2"), new Guid("7A9206BD-33AF-47af-B832-D4128730E990"));
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
                finally
                {
                    KeyHandle = IntPtr.Zero;
                }
    
                return null;
            }
    
            private static void UpdateGpoRegistrySettingsAllowFreshCredentials(string applicationName, string[] serverNames, RegistryKey rootKey, string registryPath, bool hasNTLMOnlyUpdateBeenMade)
            {
                var registryPathCredentialsDelegation = registryPath + CredentialsDelegation;
                var credentialDelegationKey = rootKey.OpenSubKey(registryPathCredentialsDelegation, true);
    
                try
                {
                    if (credentialDelegationKey == null)
                    {
                        credentialDelegationKey = rootKey.CreateSubKey(registryPathCredentialsDelegation, RegistryKeyPermissionCheck.ReadWriteSubTree);
                    }
    
                    credentialDelegationKey.SetValue(AllowFreshCredentials, 1, RegistryValueKind.DWord);
                    credentialDelegationKey.SetValue(ConcatenateDefaultsAllowFresh, 1, RegistryValueKind.DWord);
                }
                finally
                {
                    credentialDelegationKey.Dispose();
                    credentialDelegationKey = null;
                }
    
                var allowFreshCredentialKey = rootKey.OpenSubKey(registryPathCredentialsDelegation + BackSlash + AllowFreshCredentials, true);
    
                try
                {

                    if (allowFreshCredentialKey == null)
                    {
                        allowFreshCredentialKey = rootKey.CreateSubKey(registryPathCredentialsDelegation + BackSlash + AllowFreshCredentials, RegistryKeyPermissionCheck.ReadWriteSubTree);
                    }

                    if (allowFreshCredentialKey != null)
                    {
                        var values = allowFreshCredentialKey.ValueCount;
                        var valuesToAdd = serverNames.ToDictionary(key => string.Format(CultureInfo.InvariantCulture, @"{0}/{1}", applicationName, key), value => value);
                        var valueNames = allowFreshCredentialKey.GetValueNames();
                        var existingValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    
                        foreach (var valueName in valueNames)
                        {
                            var value = allowFreshCredentialKey.GetValue(valueName).ToString();
    
                            if (!existingValues.ContainsKey(value))
                            {
                                existingValues.Add(value, value);
                            }
                        }
    
                        foreach (var key in valuesToAdd.Keys)
                        {
                            if (!existingValues.ContainsKey(key))
                            {
                                allowFreshCredentialKey.SetValue(Convert.ToString(values + 1, CultureInfo.InvariantCulture), key, RegistryValueKind.String);
                                values++;
                            }
                        }
                    }
                }
                finally
                {
                    allowFreshCredentialKey.Dispose();
                    allowFreshCredentialKey = null;
                }

                if (hasNTLMOnlyUpdateBeenMade) {
                    var allowFreshCredentialWhenNTLMOnlyKey = rootKey.OpenSubKey(registryPathCredentialsDelegation + BackSlash + AllowFreshCredentialsWhenNTLMOnly, true);
        
                    try
                    {

                        if (allowFreshCredentialWhenNTLMOnlyKey == null)
                        {
                            allowFreshCredentialWhenNTLMOnlyKey = rootKey.CreateSubKey(registryPathCredentialsDelegation + BackSlash + AllowFreshCredentialsWhenNTLMOnly, RegistryKeyPermissionCheck.ReadWriteSubTree);
                        }

                        if (allowFreshCredentialWhenNTLMOnlyKey != null)
                        {
                            var values = allowFreshCredentialWhenNTLMOnlyKey.ValueCount;
                            var valuesToAdd = serverNames.ToDictionary(key => string.Format(CultureInfo.InvariantCulture, @"{0}/{1}", applicationName, key), value => value);
                            var valueNames = allowFreshCredentialWhenNTLMOnlyKey.GetValueNames();
                            var existingValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        
                            foreach (var valueName in valueNames)
                            {
                                var value = allowFreshCredentialWhenNTLMOnlyKey.GetValue(valueName).ToString();
        
                                if (!existingValues.ContainsKey(value))
                                {
                                    existingValues.Add(value, value);
                                }
                            }
        
                            foreach (var key in valuesToAdd.Keys)
                            {
                                if (!existingValues.ContainsKey(key))
                                {
                                    allowFreshCredentialWhenNTLMOnlyKey.SetValue(Convert.ToString(values + 1, CultureInfo.InvariantCulture), key, RegistryValueKind.String);
                                    values++;
                                }
                            }
                        }
                    }
                    finally
                    {
                        allowFreshCredentialWhenNTLMOnlyKey.Dispose();
                        allowFreshCredentialWhenNTLMOnlyKey = null;
                    }
                }
            }
    
            private static void UpdateGpoRegistrySettingsRemoveServersFromFreshCredentials(string applicationName, string[] serverNames, RegistryKey rootKey, string registryPath)
            {
                var registryPathCredentialsDelegation = registryPath + CredentialsDelegation;
    
                using (var allowFreshCredentialKey = rootKey.OpenSubKey(registryPathCredentialsDelegation + BackSlash + AllowFreshCredentials, true))
                {
                    if (allowFreshCredentialKey != null)
                    {
                        var valuesToRemove = serverNames.ToDictionary(key => string.Format(CultureInfo.InvariantCulture, @"{0}/{1}", applicationName, key), value => value);
                        var valueNames = allowFreshCredentialKey.GetValueNames();
    
                        foreach (var valueName in valueNames)
                        {
                            var value = allowFreshCredentialKey.GetValue(valueName).ToString();
                            
                            if (valuesToRemove.ContainsKey(value))
                            {
                                allowFreshCredentialKey.DeleteValue(valueName);
                            }
                        }
                    }
                }

                using (var allowFreshCredentialWhenNTLMOnlyKey = rootKey.OpenSubKey(registryPathCredentialsDelegation + BackSlash + AllowFreshCredentialsWhenNTLMOnly, true))
                {
                    if (allowFreshCredentialWhenNTLMOnlyKey != null)
                    {
                        var valuesToRemove = serverNames.ToDictionary(key => string.Format(CultureInfo.InvariantCulture, @"{0}/{1}", applicationName, key), value => value);
                        var valueNames = allowFreshCredentialWhenNTLMOnlyKey.GetValueNames();
            
                        foreach (var valueName in valueNames)
                        {
                            var value = allowFreshCredentialWhenNTLMOnlyKey.GetValue(valueName).ToString();
                                    
                            if (valuesToRemove.ContainsKey(value))
                            {
                                 allowFreshCredentialWhenNTLMOnlyKey.DeleteValue(valueName);
                            }
                        }
                    }
                }
            }
        }
    }
"@  # Cannot have leading whitespace on this line!
}

<#

.SYNOPSIS
Clean up all added global variables, constants, etc.

.DESCRIPTION
Clean up all added global variables, constants, etc.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name WsManApplication -Scope Script -Force
    Remove-Variable -Name CredSSPClientAuthPath -Scope Script -Force
    Remove-Variable -Name CredentialsDelegationPolicyPath -Scope Script -Force
    Remove-Variable -Name AllowFreshCredentialsPath -Scope Script -Force
    Remove-Variable -Name AllowFreshCredentialsWhenNTLMOnlyPath -Scope Script -Force
    Remove-Variable -Name AllowFreshCredentialsPropertyName -Scope Script -Force
    Remove-Variable -Name AllowFreshCredentialsWhenNTLMOnlyPropertyName -Scope Script -Force
    Remove-Variable -Name TypeAlreadyExistsHResult -Scope Script -Force
    Remove-Variable -Name NativeCode -Scope Script -Force
}

<#

.SYNOPSIS
Enable CredSSP client role on this computer.

.DESCRIPTION
Enable the CredSSP client role on this computer.  This computer should be a 
Windows Admin Center gateway, desktop or service mode.

#>

function enableCredSSPClient() {
    $path = "{0}:\{1}" -f $WsManApplication, $CredSSPClientAuthPath

    Set-Item -Path $path True -Force -ErrorAction SilentlyContinue -ErrorVariable +err

    return $err
}

<#

.SYNOPSIS
Get if CredSSP is enabled on client.

.DESCRIPTION
Get if CredSSP is enabled on client for client role.

#>

function getCredSSPClientEnabled() {
    $credSSPClientEnabled = $false;

    $path = "{0}:\{1}" -f $WsManApplication, $CredSSPClientAuthPath
    $credSSPClientService = Get-Item $path -ErrorAction SilentlyContinue

    if ($credSSPClientService) {
        $credSSPClientEnabled = [System.Convert]::ToBoolean($credSSPClientService.Value)
    }

    return $credSSPClientEnabled
}

<#

.SYNOPSIS
Get the CredentialsDelegation container from the registry.

.DESCRIPTION
Get the CredentialsDelegation container from the registry.  If the container
does not exist then a new one will be created.

#>

function getCredentialsDelegationItem() {
    $credentialDelegationItem = Get-Item  $CredentialsDelegationPolicyPath -ErrorAction SilentlyContinue
    if (-not ($credentialDelegationItem)) {
        $credentialDelegationItem = New-Item  $CredentialsDelegationPolicyPath
    }

    return $credentialDelegationItem
}

<#

.SYNOPSIS
Creates the CredentialsDelegation\AllowFreshCredentials / CredentialsDelegation\AllowFreshCredentialsWhenNTLMOnly container from the registry.

.DESCRIPTION
Create the CredentialsDelegation\AllowFreshCredentials / CredentialsDelegation\AllowFreshCredentialsWhenNTLMOnly container from the registry.  If the container
does not exist then a new one will be created.

#>

function createCredentialsDelegationItem([string] $propertyPath) {
    $allowFreshCredentialsItem = Get-Item $propertyPath -ErrorAction SilentlyContinue
    if (-not ($allowFreshCredentialsItem)) {
        New-Item $propertyPath
    }
}

<#

.SYNOPSIS
Set the AllowFreshCredentials / AllowFreshCredentialsWhenNTLMOnly property value in the CredentialsDelegation container.

.DESCRIPTION
Set the AllowFreshCredentials / AllowFreshCredentialsWhenNTLMOnly property value in the CredentialsDelegation container.
If the value exists then it is not changed.

#>

function setCredentialsDelegationProperty([Object] $credentialDelegationItem, [string] $propertyName) {
    $credentialDelegationItem | New-ItemProperty -Name $propertyName -Value 1 -Type DWord -Force
}

<#

.SYNOPSIS
Add the passed in server(s) to the AllowFreshCredentials / AllowFreshCredentialsWhenNTLMOnly key/container.

.DESCRIPTION
Add the passed in server(s) to the AllowFreshCredentials / AllowFreshCredentialsWhenNTLMOnly key/container. 
If a given server is already present then do not add it again.

#>

function addServersToCredentialsDelegationItem([string[]] $serverNames, [string] $propertyName, [string] $propertyPath) {
    $valuesAdded = 0

    foreach ($serverName in $serverNames) {
        $newValue = "{0}/{1}" -f $WsManApplication, $serverName

        $hasValue = $false

        # Check if any registry-value nodes values of registry-key node have certain value.
        $key = Get-ChildItem $CredentialsDelegationPolicyPath | ? PSChildName -eq $propertyName

        if ($null -eq $key) {
            New-Item -Path $propertyPath
            $valueNames = @()
        } else {
            $valueNames = @($key.GetValueNames())

            foreach ($valueName in $valueNames) {
                $value = $key.GetValue($valueName)
    
                if ($value -eq $newValue) {
                    $hasValue = $true
                    break
                }
            }
        }

        if (-not ($hasValue)) {
            New-ItemProperty $propertyPath -Name ($valueNames.Length + 1) -Value $newValue -Force
            $valuesAdded++
        }
    }

    return $valuesAdded -gt 0
}

<#

.SYNOPSIS
Add the passed in server(s) to the delegation list in the registry.

.DESCRIPTION
Add the passed in server(s) to the delegation list in the registry.

#>

function addServersToDelegation([string[]] $serverNames) {
    # Get the CredentialsDelegation key/container
    $credentialDelegationItem = getCredentialsDelegationItem

    # Test, and create if needed, the AllowFreshCredentials property value
    setCredentialsDelegationProperty $credentialDelegationItem $AllowFreshCredentialsPropertyName
    # Test, and create if needed, the AllowFreshCredentialsWhenNTLMOnly property value
    setCredentialsDelegationProperty $credentialDelegationItem $AllowFreshCredentialsWhenNTLMOnlyPropertyName

    # Create the AllowFreshCredentials key/container
    createCredentialsDelegationItem $AllowFreshCredentialsPath
    # Create the AllowFreshCredentialsWhenNTLMOnly key/container
    createCredentialsDelegationItem $AllowFreshCredentialsWhenNTLMOnlyPath

    # Add the servers to the AllowFreshCredentials key/container, if not already present
    $updateFreshCredentialsGroupPolicy = addServersToCredentialsDelegationItem $serverNames $AllowFreshCredentialsPropertyName $AllowFreshCredentialsPath
    # Add the servers to the AllowFreshCredentialsWhenNTLMOnly key/container, if not already present
    $updateFreshCredentialsWhenNTLMOnlyGroupPolicy = addServersToCredentialsDelegationItem $serverNames $AllowFreshCredentialsWhenNTLMOnlyPropertyName $AllowFreshCredentialsWhenNTLMOnlyPath

    if ($updateFreshCredentialsGroupPolicy) {
        $hasNTLMOnlyUpdateBeenMade = $false
        if ($updateFreshCredentialsWhenNTLMOnlyGroupPolicy) {
            $hasNTLMOnlyUpdateBeenMade = $true
        }
        setLocalGroupPolicy $serverNames $hasNTLMOnlyUpdateBeenMade
    }
}

<#

.SYNOPSIS
Set the local group policy to match the settings that have already been made.

.DESCRIPTION
Local Group Policy must match the settings that were made by this script to
ensure that an older Local GP setting does not overwrite the thos settings.

#>

function setLocalGroupPolicy([string[]] $serverNames, [bool] $hasNTLMOnlyUpdateBeenMade) {
    try {
        Add-Type -TypeDefinition $NativeCode
    } catch {
        if ($_.Exception.HResult -ne $TypeAlreadyExistsHResult) {
            throw $_.Exception.Message
        }
    }

    $errorMessage = [SME.LocalGroupPolicy]::EnableAllowFreshCredentialsPolicy($serverNames, $hasNTLMOnlyUpdateBeenMade)

    if ($errorMessage) {
        throw $errorMessage
    }

    return $true
}

<#

.SYNOPSIS
Main function of this script.

.DESCRIPTION
Enable CredSSP client role and add the passed in servers to the list
of servers to which this client can delegate credentials.

#>
function main([string[]] $serverNames) {
    setupScriptEnv

    try {
        # If client role is enabled, skip to adding servers to delegation
        if (-not (getCredSSPClientEnabled)) {
            # If not enabled try to enable
            $err = enableCredSSPClient
            if ($err) {
                # If there was an error, and server role is not enabled return error.
                if (-not (getCredSSPClientEnabled)) {
                    throw $err
                }
            }
        }

        addServersToDelegation $serverNames
    } catch {
        Write-Error $_.Exception.Message

        throw $_.Exception.Message
    } finally {
        cleanupScriptEnv
    }

    return $true
}

###############################################################################
# Script execution starts here
###############################################################################
if (-not ($env:pester)) {
    return main $serverNames
}
}
## [END] Enable-WACSHCredSSPClientRole ##
function Enable-WACSHCredSspManagedServer {
<#

.SYNOPSIS
Enables CredSSP on this server.

.DESCRIPTION
Enables CredSSP server role on this server.

.ROLE
Administrators

.LINK
https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/CVE-2018-0886

.LINK
https://aka.ms/CredSSP-Updates


#>

Set-StrictMode -Version 5.0
Import-Module  Microsoft.WSMan.Management -ErrorAction SilentlyContinue

function setupScriptEnv() {
    Set-Variable CredSSPServicePath -Option ReadOnly -Scope Script -Value "WSMan:\localhost\Service\Auth\CredSSP"
}

function cleanupScriptEnv() {
    Remove-Variable CredSSPServicePath -Scope Script -Force
}

<#

.SYNOPSIS
Get if CredSSP is enabled on this server.

.DESCRIPTION
Get if CredSSP is enabled on this server for server role.

#>

function getCredSSPServerEnabled()
{
    $credSSPServerEnabled = $false;

    $credSSPServerService = Get-Item $CredSSPServicePath -ErrorAction SilentlyContinue
    if ($credSSPServerService) {
        $credSSPServerEnabled = [System.Convert]::ToBoolean($credSSPServerService.Value)
    }

    return $credSSPServerEnabled
}

<#

.SYNOPSIS
Enables CredSSP on this server.

.DESCRIPTION
Enables CredSSP on this server for server role.

#>

function enableCredSSP() {
    $err = $null

    # Catching the result so that we can discard it. Otherwise it get concatinated with $err and we don't want that!
    $result = Enable-WSManCredSSP -Role Server -Force -ErrorAction SilentlyContinue -ErrorVariable +err

    return $err
}

<#

.SYNOPSIS
Main function.

.DESCRIPTION
Main function.

#>

function main() {
    setupScriptEnv
    
    try {
        # If server role is enabled then return success.
        if (-not (getCredSSPServerEnabled)) {
            # If not enabled try to enable
            $err = enableCredSSP
            if ($err) {
                # If there was an error, and server role is not enabled return error.
                if (-not (getCredSSPServerEnabled)) {
                    throw $err
                }
            }
        }
    } catch {
        Write-Error $err

        throw $_.Exception.Message
    } finally {
        cleanupScriptEnv
    }

    return $true
}

###############################################################################
# Script execution starts here...
###############################################################################
if (-not ($env:pester)) {
    return main
}

}
## [END] Enable-WACSHCredSspManagedServer ##
function Get-WACSHArcStatus {
<#

.SYNOPSIS
Check for arc agent status on server
.DESCRIPTION
Check for arc agent status on server
.ROLE
Readers

#>

$LogName = "Microsoft-ServerManagementExperience"
$LogSource = "SMEScript"
$ScriptName = "Get-ArcStatus.ps1"

Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
    

Get-Service -Name himds -ErrorVariable Err -ErrorAction SilentlyContinue | Out-Null

if(!!$Err){

    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
    -Message "[$ScriptName]: $Err"  -ErrorAction SilentlyContinue
    
    $status = "NotInstalled"
}
else {
    $status = (azcmagent show --json | ConvertFrom-Json -ErrorAction Stop).status
}

$computerSummary = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object Name

@{
    azureArcStatus = $status
    computerName = $computerSummary.Name
}


}
## [END] Get-WACSHArcStatus ##
function Get-WACSHCimWin32LogicalDisk {
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
## [END] Get-WACSHCimWin32LogicalDisk ##
function Get-WACSHCimWin32NetworkAdapter {
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
## [END] Get-WACSHCimWin32NetworkAdapter ##
function Get-WACSHCimWin32OperatingSystem {
<#

.SYNOPSIS
Gets Win32_OperatingSystem object.

.DESCRIPTION
Gets Win32_OperatingSystem object.

.ROLE
Readers

#>

##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_OperatingSystem

}
## [END] Get-WACSHCimWin32OperatingSystem ##
function Get-WACSHCimWin32PhysicalMemory {
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
## [END] Get-WACSHCimWin32PhysicalMemory ##
function Get-WACSHCimWin32Processor {
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
## [END] Get-WACSHCimWin32Processor ##
function Get-WACSHClusterInventory {
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
## [END] Get-WACSHClusterInventory ##
function Get-WACSHClusterNodes {
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
## [END] Get-WACSHClusterNodes ##
function Get-WACSHCredSspClientConfigurationOnGateway {
<#

.SYNOPSIS
Confirm the CredSSP enabled state on this computer as client role to the other computer.

.DESCRIPTION
Confirm the expected CredSSP client configuration for the gateway machine.

.ROLE
Administrators

.PARAMETER serverNames
The names of the server(s) to which this gateway should be able to forward credentials.

.LINK
https://aka.ms/CredSSP-Updates

#>

param (
    [Parameter(Mandatory=$True)]
    [string[]]$serverNames
)

Set-StrictMode -Version 5.0
Import-Module  Microsoft.WSMan.Management -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Setup all necessary global variables, constants, etc.

#>
function setupScriptEnv() {
    Set-Variable -Name WsManApplication -Option ReadOnly -Scope Script -Value "wsman"
    Set-Variable -Name CredSSPClientAuthPath -Option ReadOnly -Scope Script -Value "localhost\Client\Auth\CredSSP"
    Set-Variable -Name CredentialsDelegationPolicyPath -Option ReadOnly -Scope Script -Value "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CredentialsDelegation"
    Set-Variable -Name AllowFreshCredentialsPropertyName -Option ReadOnly -Scope Script -Value "AllowFreshCredentials"
    Set-Variable -Name AllowFreshCredentialsWhenNTLMOnlyPropertyName -Option ReadOnly -Scope Script -Value "AllowFreshCredentialsWhenNTLMOnly"
}

<#

.SYNOPSIS
Clean up all added global variables, constants, etc.

#>
function cleanupScriptEnv() {
    Remove-Variable -Name WsManApplication -Scope Script -Force
    Remove-Variable -Name CredSSPClientAuthPath -Scope Script -Force
    Remove-Variable -Name CredentialsDelegationPolicyPath -Scope Script -Force
    Remove-Variable -Name AllowFreshCredentialsPropertyName -Scope Script -Force
    Remove-Variable -Name AllowFreshCredentialsWhenNTLMOnlyPropertyName -Scope Script -Force
}

<#

.SYNOPSIS
Test 1: Is CredSSP enabled on this gateway machine?

.DESCRIPTION
If CredSSP is configured, the value set at 'localhost\Client\Auth\CredSSP' should be $True

#>
function getCredSSPClientStatus() {
    $path = "{0}:\{1}" -f $WsManApplication, $CredSSPClientAuthPath

    $credSSPClientEnabled = $false;

    $credSSPClientService = Get-Item $path -ErrorAction SilentlyContinue
    if ($credSSPClientService) {
        $credSSPClientEnabled = [System.Convert]::ToBoolean($credSSPClientService.Value)
    }

    return $credSSPClientEnabled
}

<#

.SYNOPSIS
Determines if the required CredentialsDelegation containers are in the registry.

.DESCRIPTION
Get the CredentialsDelegation container from the registry.  If the container
does not exist then we can return false since CredSSP is not configured on this
client (gateway).

#>
function verifyCredentialsDelegationRegKeyPresent([string] $registryKeyName) {
    $credentialDelegationItem = Get-Item  $CredentialsDelegationPolicyPath -ErrorAction SilentlyContinue
    if ($credentialDelegationItem) {
        $key = Get-ChildItem $CredentialsDelegationPolicyPath | Where-Object PSChildName -eq $registryKeyName

        if ($key -and $key.GetValueNames()) {
            return $True
        }
    }

    return $False
}

<#

.SYNOPSIS
Get list of delegatable servers from registry.

#>
function getDelegatableServersListForRegKey([string]$registryKeyName) {
    $serversList = @()

    $key = Get-ChildItem $CredentialsDelegationPolicyPath | Where-Object PSChildName -eq $registryKeyName
    $valueNames = $key.GetValueNames()

    foreach ($valueName in $valueNames) {
        $serversList += $key.GetValue($valueName).ToLower()
    }

    return $serversList
}

<#

.SYNOPSIS
Main function of this script.

.DESCRIPTION
Return true if the gateway is already configured as a CredSSP client, and all of the servers provided
have already been configured to allow fresh credential delegation.

#>

function main([string[]] $serverNames) {
    setupScriptEnv
    $serverNames = $serverNames | ForEach-Object { $_.ToLower() }

    # Test 1: Is client role enabled on gateway machine
    $isGatewayConfiguredAsClientRole = getCredSSPClientStatus
    $allowFreshCredentialsKeyPresent = verifyCredentialsDelegationRegKeyPresent $AllowFreshCredentialsPropertyName
    $allowFreshCredentialsWhenNTLMOnlyKeyPresent = verifyCredentialsDelegationRegKeyPresent $AllowFreshCredentialsWhenNTLMOnlyPropertyName
    $allowFreshCredentialsDelegatableServers = @()
    $allowFreshCredentialsWhenNTLMOnlyDelegatableServers = @()
    $notDelegatableServers = @()

    # Test 2: Determines which servers could be delegated fresh credentials from gateway
    if ($allowFreshCredentialsKeyPresent) {
        $freshCredentialsServers = getDelegatableServersListForRegKey $AllowFreshCredentialsPropertyName
        [array]$allowFreshCredentialsDelegatableServers = $serverNames | Where-Object { $freshCredentialsServers.Contains("$WsManApplication/$($_)") }
    }
    if (-not $allowFreshCredentialsDelegatableServers) {
        $allowFreshCredentialsDelegatableServers = @()
    }

    if ($allowFreshCredentialsWhenNTLMOnlyKeyPresent) {
        $freshCredentialsWhenNTLMServers = getDelegatableServersListForRegKey $AllowFreshCredentialsWhenNTLMOnlyPropertyName
        [array]$allowFreshCredentialsWhenNTLMOnlyDelegatableServers = $serverNames | Where-Object { $freshCredentialsWhenNTLMServers.Contains("$WsManApplication/$($_)") }
    }
    if (-not $allowFreshCredentialsWhenNTLMOnlyDelegatableServers) {
        $allowFreshCredentialsWhenNTLMOnlyDelegatableServers = @()
    }

    [array]$notDelegatableServers = $serverNames | Where-Object { -not $allowFreshCredentialsDelegatableServers.Contains($_) -and -not $allowFreshCredentialsWhenNTLMOnlyDelegatableServers.Contains($_) }
    if (-not $notDelegatableServers) {
        $notDelegatableServers = @()
    }

    cleanupScriptEnv

    $result = [PSCustomObject]@{
        IsGatewayConfiguredAsClientRole = $isGatewayConfiguredAsClientRole
        AllowFreshCredentialsKeyPresent = $allowFreshCredentialsKeyPresent
        AllowFreshCredentialsWhenNTLMOnlyKeyPresent = $allowFreshCredentialsWhenNTLMOnlyKeyPresent
        AllowFreshCredsDelegatableServers = $allowFreshCredentialsDelegatableServers
        AllowFreshCredsWhenNTLMOnlyDelegatableServers = $allowFreshCredentialsWhenNTLMOnlyDelegatableServers
        NotDelegatableServers = $notDelegatableServers
    }

    return $result
}

###############################################################################
# Script execution starts here
###############################################################################

return main $serverNames

}
## [END] Get-WACSHCredSspClientConfigurationOnGateway ##
function Get-WACSHCredSspClientRole {
<#

.SYNOPSIS
Gets the CredSSP enabled state on this computer as client role to the other computer.

.DESCRIPTION
Gets the CredSSP enabled state on this computer as client role to the other computer.

.ROLE
Administrators

.PARAMETER serverNames
The names of the server to which this gateway can forward credentials.

.LINK
https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/CVE-2018-0886

.LINK
https://aka.ms/CredSSP-Updates

#>

param (
    [Parameter(Mandatory=$True)]
    [string[]]$serverNames
)

Set-StrictMode -Version 5.0
Import-Module  Microsoft.WSMan.Management -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Setup all necessary global variables, constants, etc.

.DESCRIPTION
Setup all necessary global variables, constants, etc.

#>

function setupScriptEnv() {
    Set-Variable -Name WsManApplication -Option ReadOnly -Scope Script -Value "wsman"
    Set-Variable -Name CredSSPClientAuthPath -Option ReadOnly -Scope Script -Value "localhost\Client\Auth\CredSSP"
    Set-Variable -Name CredentialsDelegationPolicyPath -Option ReadOnly -Scope Script -Value "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CredentialsDelegation"
    Set-Variable -Name AllowFreshCredentialsPropertyName -Option ReadOnly -Scope Script -Value "AllowFreshCredentials"
    Set-Variable -Name AllowFreshCredentialsWhenNTLMOnlyPropertyName -Option ReadOnly -Scope Script -Value "AllowFreshCredentialsWhenNTLMOnly"
}

<#

.SYNOPSIS
Clean up all added global variables, constants, etc.

.DESCRIPTION
Clean up all added global variables, constants, etc.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name WsManApplication -Scope Script -Force
    Remove-Variable -Name CredSSPClientAuthPath -Scope Script -Force
    Remove-Variable -Name CredentialsDelegationPolicyPath -Scope Script -Force
    Remove-Variable -Name AllowFreshCredentialsPropertyName -Scope Script -Force
    Remove-Variable -Name AllowFreshCredentialsWhenNTLMOnlyPropertyName -Scope Script -Force
}

<#

.SYNOPSIS
Is CredSSP client role enabled on this server.

.DESCRIPTION
When the CredSSP client role is enabled on this server then return $true.

#>

function getCredSSPClientEnabled() {
    $path = "{0}:\{1}" -f $WsManApplication, $CredSSPClientAuthPath

    $credSSPClientEnabled = $false;

    $credSSPClientService = Get-Item $path -ErrorAction SilentlyContinue
    if ($credSSPClientService) {
        $credSSPClientEnabled = [System.Convert]::ToBoolean($credSSPClientService.Value)
    }

    return $credSSPClientEnabled
}

<#

.SYNOPSIS
Are the servers already configure to delegate fresh credentials?

.DESCRIPTION
Are the servers already configure to delegate fresh credentials?

#>

function getServersDelegated([string[]] $serverNames, [string] $propertyName) {
    $valuesFound = 0

    foreach ($serverName in $serverNames) {
        $newValue = "{0}/{1}" -f $WsManApplication, $serverName

        # Check if any registry-value nodes values of registry-key node have certain value.
        $key = Get-ChildItem $CredentialsDelegationPolicyPath | ? PSChildName -eq $propertyName
        $valueNames = $key.GetValueNames()

        foreach ($valueName in $valueNames) {
            $value = $key.GetValue($valueName)

            if ($value -eq $newValue) {
                $valuesFound++
                break
            }
        }
    }

    return $valuesFound -eq $serverNames.Length
}

<#

.SYNOPSIS
Detemines if the required CredentialsDelegation containers are in the registry.

.DESCRIPTION
Get the CredentialsDelegation container from the registry.  If the container
does not exist then we can return false since CredSSP is not configured on this
client (gateway).

#>

function areCredentialsDelegationItemsPresent([string] $propertyName) {
    $credentialDelegationItem = Get-Item  $CredentialsDelegationPolicyPath -ErrorAction SilentlyContinue
    if ($credentialDelegationItem) {
        $key = Get-ChildItem $CredentialsDelegationPolicyPath | ? PSChildName -eq $propertyName

        if ($key) {
            $valueNames = $key.GetValueNames()
            if ($valueNames) {
                return $true
            }
        }
    }

    return $false
}

<#

.SYNOPSIS
Main function of this script.

.DESCRIPTION
Return true if the gateway is already configured as a CredSSP client, and all of the servers provided
have already been configured to allow fresh credential delegation.

#>

function main([string[]] $serverNames) {
    setupScriptEnv

    $isServersDelegatedToAllowFreshCredentials = $false
    $isServersDelegatedToAllowFreshCredentialsWhenNTLMOnly = $false

    $isClientEnabled = getCredSSPClientEnabled
    
    if (areCredentialsDelegationItemsPresent $AllowFreshCredentialsPropertyName) {
        $isServersDelegatedToAllowFreshCredentials = getServersDelegated $serverNames $AllowFreshCredentialsPropertyName
    }

    if (areCredentialsDelegationItemsPresent $AllowFreshCredentialsWhenNTLMOnlyPropertyName) {
        $isServersDelegatedToAllowFreshCredentialsWhenNTLMOnly = getServersDelegated $serverNames $AllowFreshCredentialsWhenNTLMOnlyPropertyName
    }

    cleanupScriptEnv

    return $isClientEnabled -and $isServersDelegatedToAllowFreshCredentials -and $isServersDelegatedToAllowFreshCredentialsWhenNTLMOnly
}

###############################################################################
# Script execution starts here
###############################################################################

return main $serverNames

}
## [END] Get-WACSHCredSspClientRole ##
function Get-WACSHCredSspManagedServer {
<#

.SYNOPSIS
Gets the CredSSP server role on this server.

.DESCRIPTION
Gets the CredSSP server role on this server.

.ROLE
Administrators

.Notes
The feature(s) that use this script are still in development and should be considered as being "In Preview".
Therefore, those feature(s) and/or this script may change at any time.

#>

Set-StrictMode -Version 5.0
Import-Module  Microsoft.WSMan.Management -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Setup all necessary global variables, constants, etc.

.DESCRIPTION
Setup all necessary global variables, constants, etc.

#>

function setupScriptEnv() {
    Set-Variable -Name WsManApplication -Option ReadOnly -Scope Script -Value "wsman"
    Set-Variable -Name CredSSPServiceAuthPath -Option ReadOnly -Scope Script -Value "localhost\Service\Auth\CredSSP"
}

<#

.SYNOPSIS
Clean up all added global variables, constants, etc.

.DESCRIPTION
Clean up all added global variables, constants, etc.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name WsManApplication -Scope Script -Force
    Remove-Variable -Name CredSSPServiceAuthPath -Scope Script -Force
}

<#

.SYNOPSIS
Is CredSSP server role enabled on this server.

.DESCRIPTION
When the CredSSP server role is enabled on this server then return $true.

#>

function getCredSSPServerEnabled() {
    $path = "{0}:\{1}" -f $WsManApplication, $CredSSPServiceAuthPath

    $credSSPServerEnabled = $false;

    $credSSPServerService = Get-Item $path -ErrorAction SilentlyContinue
    if ($credSSPServerService) {
        $credSSPServerEnabled = [System.Convert]::ToBoolean($credSSPServerService.Value)
    }

    return $credSSPServerEnabled
}

<#

.SYNOPSIS
Main function.

.DESCRIPTION
Main function.

#>

function main() {
    setupScriptEnv

    $result = getCredSSPServerEnabled

    cleanupScriptEnv

    return $result
}

###############################################################################
# Script execution starts here...
###############################################################################

if (-not ($env:pester)) {
    return main
}

}
## [END] Get-WACSHCredSspManagedServer ##
function Get-WACSHDecryptedDataFromNode {
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
## [END] Get-WACSHDecryptedDataFromNode ##
function Get-WACSHEncryptionJWKOnNode {
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
## [END] Get-WACSHEncryptionJWKOnNode ##
function Get-WACSHRebootPendingStatus {
<#

.SYNOPSIS
Gets information about the server pending reboot.

.DESCRIPTION
Gets information about the server pending reboot.

.ROLE
Readers

#>

import-module CimCmdlets

function Get-ComputerNameChangeStatus {
  $currentComputerName = (Get-ItemProperty "HKLM:SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName"-ErrorAction SilentlyContinue).ComputerName
  $activeComputerName = (Get-ItemProperty "HKLM:SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName"-ErrorAction SilentlyContinue).ComputerName
  return $currentComputerName -ne $activeComputerName
}

function Get-CCMRebootPending {
  $status = Invoke-WmiMethod -Namespace "ROOT\ccm\ClientSDK" -Class CCM_ClientUtilities -Name ` DetermineIfRebootPending  -ErrorAction SilentlyContinue
  return $status
}

function Test-PendingReboot {
  if ($value = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction Ignore) {
    return @{
      rebootRequired        = $true
      additionalInformation = $value
    }
  }

  if ($value = Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction Ignore) {
    return @{
      rebootRequired        = $true
      additionalInformation = $value
    }
  }

  if (Get-ComputerNameChangeStatus) {
    return @{
      rebootRequired        = $true
      additionalInformation = 'Computer ID changed'
    }
  }

  $status = Get-CCMRebootPending
  if (($status -ne $null) -and $status.RebootPending) {
    return @{
      rebootRequired        = $true
      additionalInformation = $status
    }
  }

  return @{
    rebootRequired        = $false
    additionalInformation = $null
  }
}

Test-PendingReboot

}
## [END] Get-WACSHRebootPendingStatus ##
function Get-WACSHScheduledTask {
<#

.SYNOPSIS
Gets details of scheduled task.
.DESCRIPTION

.ROLE
Administrators

#>

Import-Module ScheduledTasks -ErrorAction SilentlyContinue

Get-ScheduledTaskInfo -TaskName WACInstallUpdates

}
## [END] Get-WACSHScheduledTask ##
function Get-WACSHServerInventory {
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
## [END] Get-WACSHServerInventory ##
function Get-WACSHSystemLockdownPolicy {
<#

.SYNOPSIS
Gets an EnforcementMode that describes the system lockdown policy on this computer.

.DESCRIPTION
By checking the system lockdown policy, we can infer if PowerShell is in ConstrainedLanguage mode as a result of an enforced WDAC policy.
Note: $ExecutionContext.SessionState.LanguageMode should not be used within a trusted (by the WDAC policy) script context for this purpose because
the language mode returned would potentially not reflect the system-wide lockdown policy/language mode outside of the execution context.

.ROLE
Readers

#>

return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy().ToString()

}
## [END] Get-WACSHSystemLockdownPolicy ##
function Get-WACSHUpdateStatus {
<#

.SYNOPSIS
Check for failed update
.DESCRIPTION
check for failed update
.ROLE
Readers

#>
$fileContent = Get-Content $env:temp\InstallLog.txt -ErrorAction SilentlyContinue
if($null -eq $fileContent){
  return @{failedToUpdate = $false}
}
$file_data = Get-Content $env:temp\InstallLog.txt | Where-Object { $_ -like '*Windows Admin Center -- Installation failed*' }
if ( $file_data ) {
    return @{failedToUpdate = $true}
}
return @{failedToUpdate = $false}

}
## [END] Get-WACSHUpdateStatus ##
function New-WACSHWindowsAdminCenterUpdateSchedule {
<#

.SYNOPSIS
Create a scheduled task to run a powershell script file to installs all available WAC updates through

.DESCRIPTION
Schedule an installation for a WAC Update using the task scheduler

.ROLE
Administrators

.PARAMETER installTime
  The user-defined time to install the update

#>

param (
  [Parameter(Mandatory = $true)]
  [string]$installTime,
  [Parameter(Mandatory = $true)]
  [string]$downloadWACUrl,
  [Parameter(Mandatory = $true)]
  [bool]$installNow,
  [Parameter(Mandatory = $true)]
  [boolean]
  $fromTaskScheduler
)

Import-Module ScheduledTasks -ErrorAction SilentlyContinue

function newWindowsAdminCenterUpdateSchedule() {
  param (
    [string]
    $downloadWACUrl
  )
  function Get-CurrentSslCertificates([string]$portNumber) {
    Write-Host "Retrieving current certificate for $portNumber"
    $netshCommand = "netsh http show sslcert ipport=0.0.0.0:{0}" -f $portNumber
    $portCertificate = (Invoke-Expression $netshCommand) -Join ' '
    $pattern = "Certificate hash\s+: (\w+)"
    $thumbprint = $portCertificate | Microsoft.PowerShell.Utility\Select-String -Pattern $pattern | ForEach-Object { $_ -match $pattern > $null; $matches[1] };
    return @{ thumbprint = $thumbprint }
  }

  function Set-ErrorLogContent {
    Param(
      [Parameter(Mandatory = $True)]
      [bool] $failedToUpdate,
      [Parameter(Mandatory = $True)]
      [string] $errorCode,
      [Parameter(Mandatory = $True)]
      [string] $errorMessage
    )
    $errorLogFile = $env:temp + "\WindowsAdminCenterErrorLogFile.json"
    $fileContents = ""
    $errorContent = "{ ""failedToUpdate"": ""$($failedToUpdate)"", ""time"": ""$(Get-Date)"",  ""errorCode"": ""$($errorCode)"", ""errorMessage"": ""$($errorMessage)""}"

    if (-Not(Test-Path $errorLogFile)) {
      new-item $errorLogFile
      $json = '{
          "errors":[]
        }'
      $fileContents = ConvertFrom-Json -InputObject $json
    }
    else {
      $fileContents = ([System.IO.File]::ReadAllText($errorLogFile)  | ConvertFrom-Json)

      # ensuring correctness of error content format. It will error if fileContents.errors is null otherwise and not write to the file.
      if ($fileContents.errors -eq $null) {
        $json = '{
              "errors":[]
            }'
        $fileContents = ConvertFrom-Json -InputObject $json
      }
    }

    $fileContents.errors += $errorContent
    ConvertTo-Json -InputObject $fileContents | Out-File $errorLogFile
  }

  $serviceFound = Get-Service ServerManagementGateway -ErrorAction SilentlyContinue
  if ($null -eq $serviceFound) {
    $isServiceMode = $false
  }
  else {
    $isServiceMode = $true
  }
  $isHAEnabled = "false"
  if (Test-Path -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ServerManagementGateway\Ha") {
    $HARegistryResult = Get-ItemPropertyValue -path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ServerManagementGateway\HA" -Name isHAEnabled
    $isHAEnabled = $HARegistryResult
  }

  $registryResult = Get-ItemPropertyValue -path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ServerManagementGateway" -Name SmePort
  $port = $registryResult
  $directoryForDownload = "$($env:Temp)"
  $tempDlPath = $directoryForDownload + "\WAC.msi"

  $privacyType = Get-ItemPropertyValue -path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ServerManagementGateway" -Name SmeTelemetryPrivacyType
  $devMode = Get-ItemPropertyValue -path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ServerManagementGateway" -Name DevMode

  if ($null -eq $devMode -or '' -eq $devMode) {
    $devMode = '0'
  }

  $testConnection = Test-NetConnection
  if ($testConnection.pingSucceeded -eq $false) {
    $errorMessage = "InternetConnectionFailed"
    Set-ErrorLogContent $true "503" $errorMessage
    return $errorMessage
  }

  if ($isServiceMode -eq $true) {
    Stop-Service ServerManagementGateway
  }
  else {
    Stop-Process -Name SmeDesktop
  }

  for ($i = 1; $i -le 5 ; $i++) {
    try {
      Invoke-WebRequest $downloadWACUrl -OutFile $tempDlPath -ErrorVariable +err
      Break
    }
    catch {
      if ($i -ge 5) {
        $statusMessage = $Error[0].Exception.Message
        $errorMessage = "WACDownloadNotFound " + $statusMessage
        Set-ErrorLogContent $true "404" $errorMessage
        return @{err = $err }
      }
      else {
        Start-Sleep -s 1
      }
    }
  }

  $logPath = "$($env:Temp)" + "\installLog.txt"

  $certResults = Get-CurrentSslCertificates($port)
  $certThumbprint = $certResults.thumbprint

  if ($isHAEnabled -eq "true") {
    $zipFilePath = $directoryForDownload + "Install-WindowsAdminCenterHA-1907.zip"
    Invoke-WebRequest "http://aka.ms/WACHAScript" -OutFile $zipFilePath
    Expand-Archive -Path $zipFilePath -DestinationPath $directoryForDownload -Verbose -Force
    Remove-Item $zipFilePath
    Set-Location $directoryForDownload # so we can locate and run the install-WindowsAdminCenterHA.ps1 script
    .\Install-WindowsAdminCenterHA.ps1 -msiPath $tempDlPath -Verbose
  }
  else {
    $exitCode = 0;
    $updateFailed = $false
    # Installation happens quietly in the background.
    if ($null -eq $certThumbprint) {
      $process = Start-Process -NoNewWindow -Wait -FilePath 'msiexec.exe' `
        -ArgumentList "/i $tempDlPath /qn /L*v ""$($logPath)"" SME_PORT=$port SSL_CERTIFICATE_OPTION=generate DEV_MODE=$devMode SME_TELEMETRY_PRIVACY_TYPE=$privacyType" `
        -Passthru
      $exitCode = $process.exitCode
    }
    else {
      $process = Start-Process -NoNewWindow -Wait -FilePath "msiexec.exe" `
        -ArgumentList "/i $tempDlPath /qn /L*v ""$($logPath)"" SME_PORT=$port SME_THUMBPRINT=$certThumprint SSL_CERTIFICATE_OPTION=installed DEV_MODE=$devMode SME_TELEMETRY_PRIVACY_TYPE=$privacyType" `
        -Passthru
      $exitCode = $process.exitCode
    }

    $errorMessage = "WAC_InstallSuccessful"
    if ( $exitCode -ne 0 ) {
      $errorMessage = "WAC_InstallFailed"
      $updateFailed = $true
    }
    Set-ErrorLogContent $updateFailed $exitCode $errorMessage

    if ($isServiceMode -eq $true) {
      Restart-Service ServerManagementGateway
    }
    else {
      Start-Process "$env:ProgramFiles\Windows Admin Center\SmeDesktop.exe"
    }
  }
  # Deleting the WAC.msi file after the install is done
  Remove-item $tempDlPath
}

enum UpdateScheduleErrorCode {
  noError = 0
  scriptNotCreated = 1
  notAnAdmin = 2
  schedulerNoConnection = 3
}

#---- Script execution starts here ----
$isWdacEnforced = $ExecutionContext.SessionState.LanguageMode -eq 'ConstrainedLanguage';

#In WDAC environment script file will already be available on the machine
#In WDAC mode the same script is executed - once normally and once through task Scheduler
if ($isWdacEnforced) {
  if ($fromTaskScheduler) {
    newWindowsAdminCenterUpdateSchedule $downloadWACUrl;
    return;
  }
}
else {
  #In non-WDAC environment script file will not be available on the machine
  #Hence, a dynamic script is created which is executed through the task Scheduler
  $ScriptFile = $env:Temp + "\WACInstall-Updates.ps1"
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
      Write-EventLog -LogName Application -Source "SME Windows Updates Install Updates" -EntryType Error -EventID 1 -Message "Can't connect to Schedule service"
      Write-Error "Can't connect to Schedule service" -ErrorAction Stop
      return @{
        IsSuccess = $false
        ErrorCode = [UpdateScheduleErrorCode]::schedulerNoConnection
      }
    }
    else {
      Start-Sleep -s 1
    }
  }
}

$cmdInstallNowValue = if ($installNow -eq $true) { "`$true" } else { "`$false" }

if ($isWdacEnforced) {
  $arg = "-WindowStyle Hidden -command ""&{Import-Module Microsoft.SME.Shell; New-WACSHWindowsAdminCenterUpdateSchedule -fromTaskScheduler `$true -installTime $installTime -installNow $cmdInstallNowValue -downloadWACUrl $downloadWACUrl}"""
}
else {
  (Get-Command newWindowsAdminCenterUpdateSchedule).ScriptBlock | Set-Content -path $ScriptFile
  if (-Not(Test-Path $ScriptFile)) {
    $message = "Failed to create file:" + $ScriptFile
    Write-Error $message
    #If failed to create script file, no need continue just return here
    return @{
      IsSuccess = $false
      ErrorCode = [UpdateScheduleErrorCode]::scriptNotCreated
    }
  }
  $arg = "-WindowStyle Hidden -File $ScriptFile -downloadWACUrl $downloadWACUrl"
}


$RootFolder = $Scheduler.GetFolder("\")
#Create a scheduled task
$taskName = "WACInstallUpdates"
#Delete existing task
if ($RootFolder.GetTasks(0) | Where-Object { $_.Name -eq $taskName }) {
  Write-Debug("Deleting existing task" + $taskName)
  Unregister-ScheduledTask -TaskName 'WACInstallUpdates' -Confirm:$false
}

$taskTrigger = New-ScheduledTaskTrigger -Once -At $installTime
$taskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -WakeToRun -RunOnlyIfNetworkAvailable
$taskAction = New-ScheduledTaskAction `
  -Execute 'powershell.exe' `
  -Argument $arg

$description = "Update WAC to the latest version at the defined time"

$id = [System.Security.Principal.WellKnownSidtype]::BuiltinAdministratorsSid
$x = new-object System.Security.Principal.SecurityIdentifier -ArgumentList $id, $null
$translatedID = $x.Translate([System.Security.Principal.NTAccount])

$taskPrincipal = New-ScheduledTaskPrincipal -GroupId $translatedID.value -RunLevel Highest

# Register the scheduled task
Register-ScheduledTask `
  -TaskName $taskName `
  -Action $taskAction `
  -Trigger $taskTrigger `
  -Description $description `
  -Settings $taskSettings `
  -Principal $taskPrincipal | Out-Null

if ($installNow -eq $true) {
  Start-ScheduledTask -TaskName "WACInstallUpdates"
}

return @{
  IsSuccess = $true
  ErrorCode = [UpdateScheduleErrorCode]::noError
}

}
## [END] New-WACSHWindowsAdminCenterUpdateSchedule ##
function Remove-WACSHInstallUpdateScheduledTask {
<#

.SYNOPSIS
Deletes an already created scheduled task.
.DESCRIPTION

.ROLE
Administrators

#>

Import-Module ScheduledTasks -ErrorAction SilentlyContinue

Unregister-ScheduledTask -TaskName 'WACInstallUpdates' -Confirm:$false

}
## [END] Remove-WACSHInstallUpdateScheduledTask ##
function Set-WACSHAzureHybridManagement {
<#

.SYNOPSIS
Onboards a machine for hybrid management.

.DESCRIPTION
Sets up a non-Azure machine to be used as a resource in Azure
The supported Operating Systems are Windows Server 2012 R2 and above.

.ROLE
Administrators

.PARAMETER subscriptionId
    The GUID that identifies subscription to Azure services

.PARAMETER resourceGroup
    The container that holds related resources for an Azure solution

.PARAMETER tenantId
    The GUID that identifies a tenant in AAD

.PARAMETER azureRegion
    The region in Azure where the service is to be deployed

.PARAMETER useProxyServer
    The flag to determine whether to use proxy server or not

.PARAMETER proxyServerIpAddress
    The IP address of the proxy server

.PARAMETER proxyServerIpPort
    The IP port of the proxy server

.PARAMETER authToken
    The authentication token for connection

.PARAMETER correlationId
    The correlation ID for the connection

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $subscriptionId,
    [Parameter(Mandatory = $true)]
    [String]
    $resourceGroup,
    [Parameter(Mandatory = $true)]
    [String]
    $tenantId,
    [Parameter(Mandatory = $true)]
    [String]
    $azureRegion,
    [Parameter(Mandatory = $true)]
    [boolean]
    $useProxyServer,
    [Parameter(Mandatory = $false)]
    [String]
    $proxyServerIpAddress,
    [Parameter(Mandatory = $false)]
    [String]
    $proxyServerIpPort,
    [Parameter(Mandatory = $true)]
    [string]
    $authToken,
    [Parameter(Mandatory = $true)]
    [string]
    $correlationId
)

Set-StrictMode -Version 5.0

<#

.SYNOPSIS
Setup script runtime environment.

.DESCRIPTION
Setup script runtime environment.

#>

function setupScriptEnv() {
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScript" -Scope Script
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Set-HybridManagement.ps1" -Scope Script
    Set-Variable -Name Machine -Option ReadOnly -Value "Machine" -Scope Script
    Set-Variable -Name HybridAgentFile -Option ReadOnly -Value "AzureConnectedMachineAgent.msi" -Scope Script
    Set-Variable -Name HybridAgentPackageLink -Option ReadOnly -Value "https://aka.ms/AzureConnectedMachineAgent" -Scope Script
    Set-Variable -Name HybridAgentExecutable -Option ReadOnly -Value "$env:ProgramFiles\AzureConnectedMachineAgent\azcmagent.exe" -Scope Script
    Set-Variable -Name HttpsProxy -Option ReadOnly -Value "https_proxy" -Scope Script
}

<#

.SYNOPSIS
Cleanup script runtime environment.

.DESCRIPTION
Cleanup script runtime environment.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name Machine -Scope Script -Force
    Remove-Variable -Name HybridAgentFile -Scope Script -Force
    Remove-Variable -Name HybridAgentPackageLink -Scope Script -Force
    Remove-Variable -Name HybridAgentExecutable -Scope Script -Force
    Remove-Variable -Name HttpsProxy -Scope Script -Force
}

<#

.SYNOPSIS
The main function.

.DESCRIPTION
Export the passed in virtual machine on this server.

#>

function main(
    [string]$subscriptionId,
    [string]$resourceGroup,
    [string]$tenantId,
    [string]$azureRegion,
    [boolean]$useProxyServer,
    [string]$proxyServerIpAddress,
    [string]$proxyServerIpPort,
    [string]$authToken,
    [string]$correlationId
) {
    $err = $null
    $args = @{}

    # Download the package
    Invoke-WebRequest -Uri $HybridAgentPackageLink -OutFile $HybridAgentFile -ErrorAction SilentlyContinue -ErrorVariable +err
    if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Couldn't download the hybrid management package. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return @()
    }

    # Install the package
    msiexec /i $HybridAgentFile /l*v installationlog.txt /qn | Out-String -ErrorAction SilentlyContinue -ErrorVariable +err
    if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Error while installing the hybrid agent package. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return @()
    }

    # Set the proxy environment variable. Note that authenticated proxies are not supported for Private Preview.
    if ($useProxyServer) {
        [System.Environment]::SetEnvironmentVariable($HttpsProxy, $proxyServerIpAddress+':'+$proxyServerIpPort, $Machine)
        $env:https_proxy = [System.Environment]::GetEnvironmentVariable($HttpsProxy, $Machine)
    }

    # Run connect command
    & $HybridAgentExecutable connect --resource-group $resourceGroup --tenant-id $tenantId --location $azureRegion `
                                     --subscription-id $subscriptionId --access-token $authToken --correlation-id $correlationId

    # Restart himds service
    Restart-Service -Name himds -ErrorAction SilentlyContinue -ErrorVariable +err
    if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Couldn't restart the himds service. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return $err
    }
}


###############################################################################
# Script execution starts here
###############################################################################
setupScriptEnv

try {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    return main $subscriptionId $resourceGroup $tenantId $azureRegion $useProxyServer $proxyServerIpAddress $proxyServerIpPort $authToken $correlationId

} finally {
    cleanupScriptEnv
}

}
## [END] Set-WACSHAzureHybridManagement ##
function Start-WACSHScheduledReboot {
<#

.SYNOPSIS
Schedule system restart.

.DESCRIPTION
This script is used to schedule system reboot

.PARAMETER restartlater
(Boolean) Flag to determine if to restart immediately or later

.PARAMETER restartTime
Time to schedule restart if restartLater flag is set to true

.PARAMETER restartReason
The reason for the restart or shutdown. p indicates that the restart or shutdown is planned.
u indicates that the reason is user defined. If neither p nor u is specified the restart or shutdown is unplanned

.PARAMETER reasonNumberMajor
The major reason number (positive integer less than 256).

.PARAMETER reasonNumberMajor
The minor reason number (positive integer less than 65536).

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [bool]$restartlater,
  [Parameter(Mandatory = $false)]
  [String]$restartTime,
  [Parameter(Mandatory = $false)]
  [String]$restartReason,
  [Parameter(Mandatory = $false)]
  [System.Int16]$reasonNumberMajor,
  [Parameter(Mandatory = $false)]
  [System.Int16]$reasonNumberMinor
)

$waitTime = 5
if ($restartlater -and -not($restartTime)) {
  # Default to 30 seconds
  $waitTime = 30
} elseif ($restartTime) {
  $waitTime = [decimal]::round(((Get-Date $restartTime) - (Get-Date)).TotalSeconds);

  # Validate timeout
  # The valid range is 0-315360000 (10 years), with a default of 30.
  # If the timeout period is greater than 0, the /f parameter is implied.
  # -30 to accommodate delays
  if ($waitTime -gt (315360000 - 30)) {
    THROW "Invalid restart time $restartTime. The valid range is 0-315360000s (10 years), with a default of 30"
  }
}

if ($waitTime -lt 5 ) {
  # Restart almost immediately, given some seconds for this PSSession to complete
  $waitTime = 5
}

$command = "Shutdown /r /t $waitTime"
if ($restartReason -and ($reasonNumberMajor -ne $null) -and ($reasonNumberMinor -ne $null)) {
  $command += " /d ${restartReason}:${reasonNumberMajor}:${reasonNumberMinor}"
}

# Reboot/ schdeule system reboot
Invoke-Expression -Command $command

}
## [END] Start-WACSHScheduledReboot ##
function Stop-WACSHReboot {
<#

.SYNOPSIS
Cancel system restart.

.DESCRIPTION
This script is used to cancel scheduled system reboot

.ROLE
Administrators
#>

# To avoid shutdown : A system shutdown has already been scheduled.(1190)
$command = "Shutdown /a"

# Cancel schdeule system reboot
Invoke-Expression -Command $command

}
## [END] Stop-WACSHReboot ##
function Test-WACSHCredSsp {
<#

.SYNOPSIS
Test CredSSP

.DESCRIPTION
Tests CredSSP

.EXAMPLE
./Test-CredSsp.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [string]
    $ServerName,
    [Parameter(Mandatory = $true)]
    [string]
    $Username,
    [Parameter(Mandatory = $true)]
    [string]
    $Password
)

$secure_password = ConvertTo-SecureString $Password -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($Username, $secure_password)
Test-WSMan -computerName $ServerName -Authentication Credssp -Credential $cred
}
## [END] Test-WACSHCredSsp ##
function Test-WACSHScriptCanBeRun {
<#

.SYNOPSIS
Test that a basic script can be run on a server.

.DESCRIPTION
Test that a basic script can be run on a server.

.ROLE
Readers

#>

param(
    [Parameter(Mandatory = $True)]
    [string]
    $nodeName
)

return @{ ScriptCanRun = $true }

}
## [END] Test-WACSHScriptCanBeRun ##

# SIG # Begin signature block
# MIIoOQYJKoZIhvcNAQcCoIIoKjCCKCYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDd0k/LehVczCwY
# CcEfXRBjmlzRjVjSC6JwO1YNdUYE/qCCDYUwggYDMIID66ADAgECAhMzAAADTU6R
# phoosHiPAAAAAANNMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI4WhcNMjQwMzE0MTg0MzI4WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDUKPcKGVa6cboGQU03ONbUKyl4WpH6Q2Xo9cP3RhXTOa6C6THltd2RfnjlUQG+
# Mwoy93iGmGKEMF/jyO2XdiwMP427j90C/PMY/d5vY31sx+udtbif7GCJ7jJ1vLzd
# j28zV4r0FGG6yEv+tUNelTIsFmmSb0FUiJtU4r5sfCThvg8dI/F9Hh6xMZoVti+k
# bVla+hlG8bf4s00VTw4uAZhjGTFCYFRytKJ3/mteg2qnwvHDOgV7QSdV5dWdd0+x
# zcuG0qgd3oCCAjH8ZmjmowkHUe4dUmbcZfXsgWlOfc6DG7JS+DeJak1DvabamYqH
# g1AUeZ0+skpkwrKwXTFwBRltAgMBAAGjggGCMIIBfjAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUId2Img2Sp05U6XI04jli2KohL+8w
# VAYDVR0RBE0wS6RJMEcxLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJh
# dGlvbnMgTGltaXRlZDEWMBQGA1UEBRMNMjMwMDEyKzUwMDUxNzAfBgNVHSMEGDAW
# gBRIbmTlUAXTgqoXNzcitW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIw
# MTEtMDctMDguY3JsMGEGCCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDEx
# XzIwMTEtMDctMDguY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIB
# ACMET8WuzLrDwexuTUZe9v2xrW8WGUPRQVmyJ1b/BzKYBZ5aU4Qvh5LzZe9jOExD
# YUlKb/Y73lqIIfUcEO/6W3b+7t1P9m9M1xPrZv5cfnSCguooPDq4rQe/iCdNDwHT
# 6XYW6yetxTJMOo4tUDbSS0YiZr7Mab2wkjgNFa0jRFheS9daTS1oJ/z5bNlGinxq
# 2v8azSP/GcH/t8eTrHQfcax3WbPELoGHIbryrSUaOCphsnCNUqUN5FbEMlat5MuY
# 94rGMJnq1IEd6S8ngK6C8E9SWpGEO3NDa0NlAViorpGfI0NYIbdynyOB846aWAjN
# fgThIcdzdWFvAl/6ktWXLETn8u/lYQyWGmul3yz+w06puIPD9p4KPiWBkCesKDHv
# XLrT3BbLZ8dKqSOV8DtzLFAfc9qAsNiG8EoathluJBsbyFbpebadKlErFidAX8KE
# usk8htHqiSkNxydamL/tKfx3V/vDAoQE59ysv4r3pE+zdyfMairvkFNNw7cPn1kH
# Gcww9dFSY2QwAxhMzmoM0G+M+YvBnBu5wjfxNrMRilRbxM6Cj9hKFh0YTwba6M7z
# ntHHpX3d+nabjFm/TnMRROOgIXJzYbzKKaO2g1kWeyG2QtvIR147zlrbQD4X10Ab
# rRg9CpwW7xYxywezj+iNAc+QmFzR94dzJkEPUSCJPsTFMIIHejCCBWKgAwIBAgIK
# YQ6Q0gAAAAAAAzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEw
# OTA5WjB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYD
# VQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+la
# UKq4BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc
# 6Whe0t+bU7IKLMOv2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4D
# dato88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+
# lD3v++MrWhAfTVYoonpy4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nk
# kDstrjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6
# A4aN91/w0FK/jJSHvMAhdCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmd
# X4jiJV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL
# 5zmhD+kjSbwYuER8ReTBw3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zd
# sGbiwZeBe+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3
# T8HhhUSJxAlMxdSlQy90lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS
# 4NaIjAsCAwEAAaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRI
# bmTlUAXTgqoXNzcitW2oynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTAL
# BgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBD
# uRQFTuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jv
# c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEF
# BQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1h
# cnljcHMuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkA
# YwB5AF8AcwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn
# 8oalmOBUeRou09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7
# v0epo/Np22O/IjWll11lhJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0b
# pdS1HXeUOeLpZMlEPXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/
# KmtYSWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvy
# CInWH8MyGOLwxS3OW560STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBp
# mLJZiWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJi
# hsMdYzaXht/a8/jyFqGaJ+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYb
# BL7fQccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbS
# oqKfenoi+kiVH6v7RyOA9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sL
# gOppO6/8MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtX
# cVZOSEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01p
# Y3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAANNTpGmGiiweI8AAAAA
# A00wDQYJYIZIAWUDBAIBBQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIEa7
# 1BwKP8JE6e1z72gIoyFVDbkvKcOKzx9LGPIFsS//MEIGCisGAQQBgjcCAQwxNDAy
# oBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20wDQYJKoZIhvcNAQEBBQAEggEASgbNA/DyeBPTJym4LtZ+qazDszKRKLQfDiud
# w5Bv1WZN4tsW4YZANiGr4Z8+OtIAMkSTsxGwNR1WVaj6gU/jp1tmQdmshVoQnas3
# d6UCdKjkSo+YlLbf2aPMSy5QZ98WEqUSo16LeW5JEQuzF/vF7mMbRokfBBjD4SVx
# Yp8DxAsoM28yBFd/XPkIe+CRbV49OMuQxRTuEIrHk/uK4kTH4m5q9ytTcc5AR7NY
# vhNQsikOi0kSxyk82U0Bw/a129+RfDDV3cRQsyCC0ecCUZIGf4/KG0TCv4ixbXc6
# 8rHnlvhTsyJz6E9nuhkGjOKlvSveQ4+yuV/9X5ZG/qv+7FfphqGCF5QwgheQBgor
# BgEEAYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZI
# AWUDBAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGE
# WQoDATAxMA0GCWCGSAFlAwQCAQUABCCcXhp3700vCBlXqKVl8bTUuT27rLb0IhG1
# iCOURzlVPQIGZVbIYQ4/GBMyMDIzMTIwODE3NDc0OS43NjRaMASAAgH0oIHRpIHO
# MIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQL
# ExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxk
# IFRTUyBFU046QTkzNS0wM0UwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1l
# LVN0YW1wIFNlcnZpY2WgghHqMIIHIDCCBQigAwIBAgITMwAAAdGyW0AobC7SRQAB
# AAAB0TANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAx
# MDAeFw0yMzA1MjUxOTEyMThaFw0yNDAyMDExOTEyMThaMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTkzNS0w
# M0UwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Uw
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCZTNo0OeGz2XFd2gLg5nTl
# Bm8XOpuwJIiXsMU61rwq1ZKDpa443RrSG/pH8Gz6XNnFQKGnCqNCtmvoKULApwrT
# /s7/e1X0lNFKmj7U7X4p00S0uQbW6LwSn/zWHaG2c54ZXsGY+BYfhWDgbFpCTxRz
# TnRCG62bkWPp6ZHbZPg4Ht1CRCAMhhOGTR8wI4G7wwWZwdMc6UvUUlq0ql9AxAfz
# kYRpi2tRvDHMdmZ3vyXpqhFwvRG8cgCH/TTCjW5q6aNbdqKL3BFDPzUtuCNsPXL3
# /E0dR2bDMqa0aNH+iIfhGC4/vcwuteOMCPUIDVSqDCNfIaPDEwYci1fd9gu1zVw+
# HEhDZM7Ea3nxIUrzt+Rfp5ToMMj4QAmJ6Uadm+TPbDbo8kFIK70ShmW8wn8fJk9R
# eQQEpTtIN43eRv9QmXy3Ued80osOBE+WkdMvSCFh+qgCsKdzQxQJG62cTeoU2eqN
# hH3oppXmyfVUwbsefQzMPtbinCZd0FUlmlM/dH+4OniqQyaHvrtYy3wqIafY3zeF
# ITlVAoP9q9vF4W7KHR/uF0mvTpAL5NaTDN1plQS0MdjMkgzZK5gtwqOe/3rTlqBz
# xwa7YYp3urP5yWkTzISGnhNWIZOxOyQIOxZfbiIbAHbm3M8hj73KQWcCR5Javgkw
# UmncFHESaQf4Drqs+/1L1QIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFAuO8UzF7DcH
# 0mmsF4XQxxHQvS2jMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8G
# A1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBs
# BggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUy
# MDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
# AwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQCbu9rTAHV24mY0
# qoG5eEnImz5akGXTviBwKp2Y51s26w8oDrWor+m00R4/3BcDmYlUK8Nrx/auYFYi
# dZddcUjw42QxSStmv/qWnCQi/2OnH32KVHQ+kMOZPABQTG1XkcnYPUOOEEor6f/3
# Js1uj4wjHzE4V4aumYXBAsr4L5KR8vKes5tFxhMkWND/O7W/RaHYwJMjMkxVosBo
# k7V21sJAlxScEXxfJa+/qkqUr7CZgw3R4jCHRkPqQhMWibXPMYar/iF0ZuLB9O89
# DMJNhjK9BSf6iqgZoMuzIVt+EBoTzpv/9p4wQ6xoBCs29mkj/EIWFdc+5a30kuCQ
# OSEOj07+WI29A4k6QIRB5w+eMmZ0Jec0sSyeQB5KjxE51iYMhtlMrUKcr06nBqCs
# SKPYsSAITAzgssJD+Z/cTS7Cu35fJrWhM9NYX24uAxYLAW0ipNtWptIeV6akuZEe
# EV6BNtM3VTk+mAlV5/eC/0Y17aVSjK5/gyDoLNmrgVwv5TAaBmq/wgRRFHmW9UJ3
# zv8Lmk6mIoAyTpqBbuUjMLyrtajuSsA/m2DnKMO0Qiz1v+FSVbqM38J/PTlhCTUb
# FOx0kLT7Y/7+ZyrilVCzyAYfFIinDIjWlM85tDeU8ZfJCjFKwq3DsRxV4JY18xww
# 8TTmod3lkr9NqGQ54LmyPVc+5ibNrjCCB3EwggVZoAMCAQICEzMAAAAVxedrngKb
# SZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmlj
# YXRlIEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIy
# NVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UE
# AxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXI
# yjVX9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjo
# YH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1y
# aa8dq6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v
# 3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pG
# ve2krnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viS
# kR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYr
# bqgSUei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlM
# jgK8QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSL
# W6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AF
# emzFER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIu
# rQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIE
# FgQUKqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWn
# G1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEW
# M2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5
# Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBi
# AEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV
# 9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3Js
# Lm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAx
# MC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2
# LTIzLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv
# 6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZn
# OlNN3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1
# bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4
# rPf5KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU
# 6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDF
# NLB62FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/
# HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdU
# CbFpAUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKi
# excdFYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTm
# dHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZq
# ELQdVTNYs6FwZvKhggNNMIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJp
# Y2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkE5MzUtMDNF
# MC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMK
# AQEwBwYFKw4DAhoDFQBHJY2Fv+GhLQtRDR2vIzBaSv/7LKCBgzCBgKR+MHwxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
# c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6R2dAzAi
# GA8yMDIzMTIwODEzNDc0N1oYDzIwMjMxMjA5MTM0NzQ3WjB0MDoGCisGAQQBhFkK
# BAExLDAqMAoCBQDpHZ0DAgEAMAcCAQACAitNMAcCAQACAhLxMAoCBQDpHu6DAgEA
# MDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAI
# AgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBAArQYkJ7iBXybEvzu7Oz/O/4UHjL
# mjjeQkZco+0OQnam4gcvPJZw7O/3Ei10KNRf5Fsk4OpVbC1fBzCMj8EPYQaorjgW
# tzKlYgW8UNwmn1QsApxW0A7FdiLh/5sxBgqQWciPylZd/ixF361NcaSMpEcwswpX
# MJ0EMp26MPjOzNneKGhd1c17sagkMyxaETtEkSqT1jTNtaSFviWaYPbVmFGR7hSr
# dTflfGG7EnLt8pT2LMQK2f5yL4akLnuf8yZpZkXInZu00i45aXEaw0NOjxWHd0yr
# T70M0GR3eKxz9EfMNG8ufh15RPxzOX5Q3jQQ51I6cS5+DtIQmgnm7ljWmpkxggQN
# MIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdGy
# W0AobC7SRQABAAAB0TANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0G
# CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCDlxHIcktHw/YbODe+Zv/xlEPVy
# OGMnKp3V0hjn7+IqTDCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIMy8YXkC
# ALv57c5sRhrPTub1q4TwJ6oVA36k8IiI/AcMMIGYMIGApH4wfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAHRsltAKGwu0kUAAQAAAdEwIgQgAWM65sxy
# Nd4lrlY9fc53fvrPBDmG/1nUFqPf2BVCOQEwDQYJKoZIhvcNAQELBQAEggIAbg1K
# tQdTmVoTDdb6120AYjDDcBnErZUgZMg8EzfTlHCd9KGf0jcQ48z/aK1jSCbS6Jl/
# sWuj/c4zHPyZN7Wl6+Kquy1K2mQJuhybqfBX2ig16i0WXEwOieCTMZZ1SajoCdHW
# y27xdJ5qph39Wd3eiSVkWvXBmpFau+tdP5KMnmNeW3is/1NAN7zcChyq6NFSLVhp
# E95rRbmOJXDlxYNVWxW+DJ9ifpE4NXfJcjFSqJnThO7oOmD14gdkp6+6b2v1jggA
# MF5TVpScROkDr4N5P88Xge6DMU4NnTj4+hIEkzjGD6D1xhh712Zwg9HVV54fY77M
# +rbqUusSqltlR3gMx7wiThNT1Y4oKk7eFnwiugEpVgQNMvBkn2p9imHfzfVoJdlC
# TRPgiV2h0Y9Gdl0Cwsx/vr6pDZrYMAc2dF1RgVo9UyqeyNC+4wn9XPhaxgfaphl6
# pQrZBcWuPEnjfgLU5dfbJ/+tQtSTS+Iy7JLNbw3BOFmpPWLp87rwSOQuL9pdnLB+
# XdZCvIrDV7BJ3V5+EnGSr1sdC5Gxj3ZYAmlntqBIg45lv+gny3wMr34NRcooz2eH
# oaBEfIIqIx4RJczBFlD788cAs58hzLkpunmJmHQRQUtsLRBtshavdlEOwjCOIs5L
# STj40PborZi7lbyJUVlfOU0MK9U1YboUUBS4Sz4=
# SIG # End signature block
