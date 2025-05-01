Function Get-OSInstallDate {
    (Get-CimInstance Win32_OperatingSystem).InstallDate
}

Function Get-OSVersion {
    (Get-CimInstance Win32_OperatingSystem).Version
}

Function Get-OSRegisteredUser {
    (Get-CimInstance Win32_OperatingSystem).RegisteredUser
}

Function Get-OSOrganization {
    (Get-CimInstance Win32_OperatingSystem).Organization
}

Function Get-OSBuildNumber {
    (Get-CimInstance Win32_OperatingSystem).BuildNumber
}

Function Test-PendingReboot {
    $PendingReboot = $false

    Push-Location "HKLM:\Software\Microsoft\Windows\CurrentVersion\"

    if (Get-ChildItem "Component Based Servicing\RebootPending" -EA Ignore) {
        $PendingReboot = $true
    }

    if (Get-Item "WindowsUpdate\Auto Update\RebootRequired" -EA Ignore) { 
        $PendingReboot = $true
    }

    Pop-Location

    Push-Location "HKLM:\SYSTEM\CurrentControlSet\Control"

    if (Get-ItemProperty "Session Manager" -Name PendingFileRenameOperations -EA Ignore) { 
        $PendingReboot = $true 
    }

    Pop-Location

    return $PendingReboot
}

Function New-RemoteDesktopShortcut {
    param (
        [string]$Path = "$ComputerName.rdp",
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $ComputerName,
        [string]$UserName = "${env:USERDOMAIN}\${env:USERNAME}"
    )

    if (Test-Path $Path) {
        $choice = Select-Item -Caption "RDP File Exists" `
            -Message "Do you want to replace the file?" `
            -choiceList "&Yes", "&No" -default 1

        if ($choice -eq 0) {
            Remove-Item $Path -Force -Confirm:$false
        }
    }

    Set-Content -Path $Path -Value @"
full address:s:$ComputerName
redirectprinters:i:0
redirectcomports:i:0
redirectsmartcards:i:0
redirectclipboard:i:1
redirectposdevices:i:0
username:s:$UserName
screen mode id:i:1
use multimon:i:0
desktopwidth:i:1366
desktopheight:i:768
session bpp:i:32
compression:i:1
"@
}

##############################################################################

Export-ModuleMember Get-OSInstallDate
Export-ModuleMember Get-OSVersion
Export-ModuleMember Get-OSRegisteredUser
Export-ModuleMember Get-OSInstallDate
Export-ModuleMember Get-OSBuildNumber

Export-ModuleMember Test-PendingReboot

Export-ModuleMember New-RemoteDesktopShortcut
