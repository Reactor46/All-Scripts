# Profile Extensions

Get-ChildItem ( Join-Path $home "Documents\WindowsPowerShell\Profile_Extensions" ) -File *".ps1" | ForEach-Object { . $_.FullName; Set-Variable -Name $_.basename -Value $_.FullName }

# Startup

Import-Module Hyper-V -RequiredVersion 1.1
Import-Module misExchange -DisableNameChecking
Get-VerseoftheDay
Cal

# Chocolatey Profile
$ChocolateyProfile = "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"
if (Test-Path($ChocolateyProfile)) {
  Import-Module "$ChocolateyProfile"
}
