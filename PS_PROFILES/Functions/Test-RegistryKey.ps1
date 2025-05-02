function Test-RegistryKey {
    <#
    .SYNOPSIS
    Tests if a registry Key and/or value exists. Can also test if the value is of a certain type.
    .DESCRIPTION
    Provides a reliable way to check registry values even when they are empty or null.
    .PARAMETER Key
    Path of the registry key (Required).
    .PARAMETER Name
    The value name (optional).
    .PARAMETER Value
    Value to compare against (optional).
    .PARAMETER Type
    The expected type of the registry value. Options: 'Binary', 'DWord', 'MultiString', 'QWord', 'String'.
    .EXAMPLE
    Test-RegistryKey -Key "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name Install
    # Returns $true if the value exists.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Key,

        [Parameter(Mandatory=$false)]
        [string]$Name,

        [Parameter(Mandatory=$false)]
        $Value,

        [Parameter(Mandatory=$false)]
        [ValidateSet('Binary','DWord','MultiString','QWord','String')]
        [Microsoft.Win32.RegistryValueKind]$Type
    )

    Begin {
        $CmdletName = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-Verbose "Running function: $CmdletName"
    }

    Process {
        $Key = Convert-RegistryPath -Key $Key

        if (-not (Test-Path -Path $Key -PathType Container)) {
            Write-Verbose "Key [$Key] does NOT exist."
            return $false
        }

        $AllKeyValues = Get-ItemProperty -Path $Key

        if ($Name) {
            if (-not ($AllKeyValues.PSObject.Properties.Name -contains $Name)) {
                Write-Verbose "Key [$Key] exists, but value [$Name] does not."
                return $false
            }

            $ValDataRead = $AllKeyValues.$Name
            Write-Verbose "Value [$Name] exists with data [$ValDataRead]."

            if ($Value -and $Value -eq $ValDataRead) {
                return $true
            }

            if ($Type) {
                $ActualType = switch ($ValDataRead.GetType().Name) {
                    "String"   { 'String' }
                    "Int32"    { 'DWord' }
                    "Int64"    { 'QWord' }
                    "String[]" { 'MultiString' }
                    "Byte[]"   { 'Binary' }
                    default    { 'Unknown' }
                }

                Write-Verbose "Value [$Name] is of type [$ActualType]."
                return ($ActualType -eq $Type)
            }

            return $true
        }

        Write-Verbose "Key [$Key] exists."
        return $true
    }

    End {
        Write-Verbose "Function [$CmdletName] execution completed."
    }
}