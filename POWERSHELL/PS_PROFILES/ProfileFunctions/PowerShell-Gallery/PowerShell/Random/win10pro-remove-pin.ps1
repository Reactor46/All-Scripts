# Script created by Pete Gypps
# Website: www.corsolutions.co.uk

# Prevents Windows 10 and Windows 11 from prompting to set up a PIN after being added to Azure AD

# Instructions:
# 1. Run this script in PowerShell with administrative privileges.
#    To do this, right-click on the PowerShell icon and choose "Run as administrator."

# Disable PIN requirement
$path = "HKLM:\SOFTWARE\Policies\Microsoft"
$key = "PassportForWork"
$name = "Enabled"
$value = "0"

try {
    New-Item -Path $path -Name $key –Force
    New-ItemProperty -Path $path\$key -Name $name -Value $value -PropertyType DWORD -Force
}
catch {
    Write-Host "Error disabling PIN requirement: $_"
}

# Delete existing PINs
$passportFolder = "C:\Windows\ServiceProfiles\LocalService\AppData\Local\Microsoft\Ngc"

if (Test-Path -Path $passportFolder) {
    $confirmation = Read-Host "Are you sure you want to delete existing PINs? (Y/N)"
    if ($confirmation -eq "Y") {
        try {
            Takeown /f $passportFolder /r /d "Y"
            ICACLS $passportFolder /reset /T /C /L /Q
            Remove-Item –path $passportFolder –recurse -force
        }
        catch {
            Write-Host "Error deleting existing PINs: $_"
        }
    }
    else {
        Write-Host "Operation cancelled."
    }
}
