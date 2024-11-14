try {
    $basePath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\"
    $paths = "Triple DES 168"
    $key = "Enabled"
    $value = "0"

    ForEach ($reg in $paths) {
        $writable = $true
        $key = (Get-Item $basePath).OpenSubKey("Ciphers", $writable).CreateSubKey($reg)
        $key.SetValue("Enabled", 0)
    }
}
catch {
    Write-Error "Error adding items to registry."
    Write-Output "Error adding items to registry."

    $Host.UI.WriteErrorLine('Error adding items to registry..')
    [Environment]::Exit(1024)
}

