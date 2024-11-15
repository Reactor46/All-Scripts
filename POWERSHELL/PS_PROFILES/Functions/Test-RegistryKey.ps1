function Test-RegistryKey {
<#
.SYNOPSIS
Tests if a registry Key and/or value exists. Can also test if the value is of certain type
.DESCRIPTION
Based on PSADT’s Set-RegistryKey function
Since registry values are treated like properties to a registry key so they have no specific path and you cannot use Test-Path to check whether a given registry value exists.
This provides a reliable way to check registry values even when it is empty or null.
Checks for registry keys
Can also do simple Value content -eq compare to save you from writing another If statement
Uses Write-Log to aid troubleshooting
.PARAMETER Key
Path of the registry key (Required).
.PARAMETER Name
The value name (optional).
.PARAMETER Value
Value to compare against (optional).
This is to save you from writing another If statement if doing a simple -eq test
If comparing ‘ExpandString’, Pre-Expand the value first
If comparing ‘Binary’, see example
If comparing ‘MultiString’, please report how you did it if successful
.PARAMETER Type
The type of registry value to create or set. Options: ‘Binary’,’DWord’,’MultiString’,’QWord’,’String’. (optional).
CAVEAT: Unable to test for ExpandString. Current test will say a value of type ExpandString is a String
UN-supported Types: ‘ExpandString’,’None’,’Unknown’

.EXAMPLE
Key exists tests
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full”	#Key exists in win7	(returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP”	#(returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Do Not Exist”	#(returns $false)

Value exist tests
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full” -Name Install	#(returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4” -Name ‘DoesNotExist’ #(returns $false)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4” -Name ‘IMblank’ #=<blank string>(returns $true)

Value Type tests
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full” -Name Install -Type DWord	#REG_SZ (returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4” -Name ‘IMblank’ -Type String	#=<blank string>(returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4” -Name ‘IMblank’ -Type MultiString	#=<blank string>(returns $false)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment” -Name Temp -Type DWord	#REG_EXP_SZ (returns $false)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment” -Name Temp -Type String	#REG_EXP_SZ (returns $true)

(Default) Value tests
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4” -Name ‘(Default)’ #=Not set (returns $false)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4.0” -Name ‘(Default)’	#=deprecated	(returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4.0” -Name ‘(Default)’ -Type ‘String’	#=deprecated (returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4.0” -Name ‘(Default)’ -Type ‘MultiString’	#=deprecated (returns $false)

Value Content tests
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full” -Name Install -Value 1	#REG_DWORD=1 returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full” -Name Install -Value “1”	#REG_DWORD=1 returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5” -Name InstallPath -Value “C:\Windows\Microsoft.NET\Framework64\v3.5\” #REG_SZ (returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5” -Name Version -Value “3.5.30729.5420”	#REG_SZ (returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment” -Name ComSpec -Value “%SystemRoot%\system32\cmd.exe”	#REG_EXP_SZ (returns $false)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment” -Name ComSpec -Value “C:\Windows\system32\cmd.exe”	#REG_SZ (returns $true)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management” -Name “ExistingPageFiles” -Value “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management” #REG_MULTI_SZ (returns $false)
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power” -Name SystemPowerPolicy -Value “010101” #REZ_BINARY (returns $false)
$test=”1 0 0 0 2 0 0 0 0 0 0 0 0 0 0 0 2 0 0 0 0 0 0 0 0 0 0 0 2 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 0 0 2 0 0 0 0 0 0 0 0 0 0 0 16 14 0 0 90 0 0 0 4 0 0 0 4 0 0 0 1 0 0 0 1 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 0 0 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 0 0 10 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 0 0 0 176 4 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 176 4 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0″
Test-RegistryKey -Key “HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power” -Name SystemPowerPolicy -Value $test #returns $true, but machine specific
.NOTES
Limits:
a (default) value that is “Not Set” = “does not exist” and will return $False
Cannot test for both ValueData and Type in the same function call
#>
[CmdletBinding()]
param(
[Parameter(Mandatory=$true,HelpMessage = “Path of the registry key”)]
[ValidateNotNullorEmpty()]
[string]$Key,
[Parameter(Mandatory=$false,HelpMessage = “The value name (optional). For a key’s default value use (default)”)]
[ValidateNotNull()]
[String]$Name,
[Parameter(Mandatory=$false,HelpMessage = “Value Type (optional). Use ‘String’ for ‘ExpandString'”)]
[ValidateNotNullOrEmpty()]
[ValidateSet(‘Binary’,’DWord’,’MultiString’,’QWord’,’String’)]
[Microsoft.Win32.RegistryValueKind]$Type,
[Parameter(Mandatory=$false)]
$Value	#do NOT cast data type! $null or empty could also be given!
)

Begin {
## Get the name of this function and write header
[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
}
Process {
$Key = Convert-RegistryPath -Key $Key

if( -not (Test-Path -Path $Key -PathType Container) ) {
Write-Host “Key [$Key] does NOT Exists” -Source ${CmdletName}
return $false #no point testing for the value
}

$AllKeyValues = Get-ItemProperty -Path $Key
If ($PSBoundParameters.ContainsKey(‘ValueName’) ) {	#	If ($Name) {
if ( -not $AllKeyValues ) {
Write-log “Key [$Key] Exists but has no values!” -Source ${CmdletName}
return $false
}

$ValDataRead = Get-Member -InputObject $AllKeyValues -Name $Name	#CAVEAT: converts REG_SZ to Int32 if it’s a number
Write-log “Value [$Name] Exists and contains [$ValDataRead]” -Source ${CmdletName}
if( $ValDataRead ) {
$ValDataRead = $($AllKeyValues.$Name)	#CAVEAT: converts REG_SZ to Int32 if it’s a number
Write-log “Value [$Name] Exists and contains [$ValDataRead]” -Source ${CmdletName}

If ($Value) { # do a data compare (Why not, we have everything else)
#	$ValDataRead = Get-ItemProperty -Path $Key -Name $Name	#CAVEAT: returns hash array @{Install=1}
#	Write-log “Value [$Name] Exists and contains [$ValDataRead]”
If ($Value -eq $ValDataRead) {	#If there is no way to
Return $true
} Else { Return $false }
} ElseIf ($Type) {	# do a Type compare (Why not, we have everything else)
$ValTypeRead = switch ($Name.gettype().Name) {
“String”{‘String’}	# {“REG_SZ”; } # also REG_EXP_SZ [ hex(2) ]
“Int32” {‘DWord’}	# {“REG_DWORD”; }
“Int64” {‘QWord’}	# {“REG_QWORD”; }
“String[]” {‘MultiString’}# {“REG_MULTI_SZ”; }
“Byte[]” {‘Binary’}	# {“REG_BINARY”}
default {‘Unknown’}
}
Write-log “Value [$Name] is of type [$ValTypeRead]” -Source ${CmdletName}
If ($ValTypeRead -eq $Type) {
Return $true
} else {return $false }

} Else {
return $true
} # If ($Value) {
} else {
Write-log “Key [$Key] exist but [$Name] does not exist” -Source ${CmdletName}
return $false
}
} Else { #The Key exists but we don’t care about its values
Write-log “Key [$Key] Exists” -Source ${CmdletName}
return $true
}
}
End {
Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
}

}	# Test-RegistryKey