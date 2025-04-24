# Test_Youtube.ps1

$result = Test-Connection -ComputerName youtube.com -Count 1

if ($result -ne $null) {
    Write-Output "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): Connection to YouTube is successful."
} else {
    Write-Output "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): Unable to connect to YouTube."
}
