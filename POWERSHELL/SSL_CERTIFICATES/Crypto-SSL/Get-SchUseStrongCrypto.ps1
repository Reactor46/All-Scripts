<#
In Summery
This script gets the list of all cipher suites available in the current user context by querying the SchUseStrongCrypto
registry value under HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319. It then filters out any null values and converts
the cipher suites to a human-readable format using the ToString() method of the System.Security.Cryptography.SslProtocols
enum. Finally, the script displays the list of cipher suites in the console.
#>

# Get the list of all cipher suites available in the current user context
$cipherSuites = Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319' -Recurse |
  Get-ItemProperty -Name 'SchUseStrongCrypto' -EA 0 |
  Select-Object -ExpandProperty 'SchUseStrongCrypto'

# Filter out null values and convert the cipher suites to a human-readable format
$cipherSuites = $cipherSuites | Where-Object { $_ } | ForEach-Object {
  $cipher = [System.Security.Cryptography.SslProtocols]$_
  $cipher.ToString()
}

# Display the list of cipher suites
Write-Host "Available cipher suites:" -ForegroundColor Yellow
$cipherSuites | ForEach-Object { Write-Host "- $_" }