#==========================================================================
#
# CREATE SECURE PASSWORD FILES
#
# AUTHOR: Dennis Span (https://dennisspan.com)
# DATE : 05.04.2017
#
# COMMENT:
# -This script generates a 256-bit AES key file and a password file
# -In order to use this PowerShell script, start it interactively (select this file
# in Windows Explorer. With a right-mouse click select 'Run with PowerShell')
#
#==========================================================================
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
# Define variables
$Directory = "C:\Scripts\" # change this location to your needs
$KeyFile = Join-Path $Directory "BBP_OPT_AES_KEY_FILE.key" # change this file name to your needs
$PasswordFile = Join-Path $Directory "BBP_OPT_AES_PASSWORD_FILE.txt" # change this file name to your needs
# Text for the console
Write-Host "CREATE SECURE PASSWORD FILE"
Write-Host ""
Write-Host "Comments:"
Write-Host "This script creates a 256-bit AES key file and a password file"
Write-Host "containing the password you enter below."
Write-Host ""
Write-Host "Two files will be generated in the directory $($Directory):"
Write-Host "-$($KeyFile)"
Write-Host "-$($PasswordFile)"
Write-Host ""
Write-Host "Enter password and press ENTER:"
$Password = Read-Host -AsSecureString
Write-Host ""
# Create the AES key file
try {
$Key = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Key)
$Key | out-file $KeyFile
$KeyFileCreated = $True
Write-Host "The key file $KeyFile was created successfully"
} catch {
write-Host "An error occurred trying to create the key file $KeyFile (error: $($Error[0])"
}
Start-Sleep 2
# Add the plaintext password to the password file (and encrypt it based on the AES key file)
If ( $KeyFileCreated -eq $True ) {
try {
$Key = Get-Content $KeyFile
$Password | ConvertFrom-SecureString -key $Key | Out-File $PasswordFile
Write-Host "The key file $PasswordFile was created successfully"
} catch {
write-Host "An error occurred trying to create the password file $PasswordFile (error: $($Error[0])"
}
}
Write-Host ""
write-Host "End of script (press any key to quit)"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

##
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 # ensures correct TLS settings.
# Define variables
$Account = "<domain>/<username>" #
# Credentials
$KeyFile = Join-Path $Directory "BBP_OPT_AES_KEY_FILE.key" # change this to the file name you used in creating the secure password
$PasswordFile = Join-Path $Directory "BBP_OPT_AES_PASSWORD_FILE.txt" # change this to the file name you used in creating the secure password
# Read the secure password from a password file and decrypt it to a normal readable string
$SecurePassword = ( (Get-Content $PasswordFile) | ConvertTo-SecureString -Key (Get-Content $KeyFile) )
# Convert the standard encrypted password stored in the password file to a secure string using the AES key file
$SecurePasswordInMemory = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword);
# Write the secure password to unmanaged memory (specifically to a binary or basic string)
$PasswordAsString = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($SecurePasswordInMemory);
# Read the plain-text password from memory and store it in a variable
[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($SecurePasswordInMemory);
# Example of Mapping a drive with Account and SecurePassword
$Map = (New-SmbMapping -RemotePath "\\10.251.68.61\NextGenRoot" -LocalPath "X:" -UserName $Account -Password $PasswordAsString)