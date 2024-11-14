try {
    $path = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\Schannel\Protocols\SSL 3.0"
    $key = "DisabledByDefault"
    $value = "1"

    If(!(Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
        New-ItemProperty -Path $path -Name $key -Value $value -PropertyType DWORD -Force | Out-Null
    }
    else {
        New-ItemProperty -Path $path -Name $key -Value $value -PropertyType DWORD -Force | Out-Null
    }
    
    If(Test-Path $path) {
        Write-Output "Registry $path added successfully."
    }
}
catch {
    Write-Error "Error adding items to registry."
    Write-Output "Error adding items to registry."

    $Host.UI.WriteErrorLine('Error adding items to registry..')
    [Environment]::Exit(1024)
}

