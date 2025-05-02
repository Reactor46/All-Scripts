# Function to check software
function CheckSoftware {
    param (
        [string]$SoftwareName,
        [string]$FilePath
    )

    # Check if the file exists at the specified path
    if (Test-Path $FilePath) {
        try {
            # Attempt to get the version information
            $versionInfo = (Get-Item $FilePath).VersionInfo
            # Display the product version
            Write-Host "$SoftwareName`;$($versionInfo.ProductVersion)" -ForegroundColor Green
	        #Write-Host "$SoftwareName $($versionInfo.ProductVersion)" -ForegroundColor Green
	        #Write-Host "$SoftwareName`t$($versionInfo.ProductVersion)" -ForegroundColor Green
        } catch {
            # If an error occurs while retrieving version information, display an error message
            Write-Host "Error: Unable to retrieve version information for $SoftwareName." -ForegroundColor Red
        }
    } else {
        # If the file does not exist at the specified path, display a message indicating that the software is not installed
        Write-Host "$SoftwareName is not installed or the installed location was changed during installation/upgrade." -ForegroundColor Red
    }
}

function CheckFileZillaClient
{
$filezillaKey = Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\FileZilla Client'

    if ($filezillaKey -ne $null) {
        $version = $filezillaKey.Version
        if ($version -ne $null) {
            Write-Host "FileZilla Client version`;$version" -ForegroundColor Green
        } else {
            Write-Host "FileZilla version information not found." -ForegroundColor Yellow
        }
    } else {
        Write-Host "FileZilla Client is not installed." -ForegroundColor Red
    }
     #$FileZillaExe = "C:\Program Files\FileZilla FTP Client\filezilla.exe"
     #
    #if (Test-Path $FileZillaExe) {
    #    $Version = (Get-Item $FileZillaExe).VersionInfo.ProductVersion
    #    Write-Host "FileZilla FTP Client is installed. Version: $Version" -ForegroundColor Yellow
    #} else {
    #    Write-Host "FileZilla FTP Client is not installed or the installation path has changed." -ForegroundColor Yellow
    #}
}

function CheckWindowsOSVersion {
    #$majorVersion = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentMajorVersionNumber
    #$minorVersion = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentMinorVersionNumber
    #$releaseId = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').ReleaseId
    #Write-Host "Windows OS Version`;$majorVersion.$minorVersion.$releaseId" -ForegroundColor Green
    $version = Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty Version
    Write-Host "Windows OS Version`;$version" -ForegroundColor Green
}

#Set-ExecutionPolicy -Scope Process

# Call the function to check Adobe Reader
CheckSoftware -SoftwareName "Adobe Reader" -FilePath "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"

# Call the function to check Mozilla Firefox
CheckSoftware -SoftwareName "Mozilla Firefox" -FilePath "C:\Program Files\Mozilla Firefox\firefox.exe"


# Call the function to check Google Chrome
CheckSoftware -SoftwareName "Google Chrome" -FilePath "C:\Program Files\Google\Chrome\Application\chrome.exe"

# Call the function to check Outlook
CheckSoftware -SoftwareName "Outlook" -FilePath "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
CheckSoftware -SoftwareName "Outlook" -FilePath "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"

# Call the function to check 7-Zip
CheckSoftware -SoftwareName "7-Zip" -FilePath "C:\Program Files\7-Zip\7zFM.exe"

# Call the function to check Notepad ++
CheckSoftware -SoftwareName "Notepad ++" -FilePath "C:\Program Files\Notepad++\notepad++.exe"

# Call the function to check FileZilla Client
CheckFileZillaClient -SoftwareName "FileZilla Client" #-FilePath "C:\Program Files\FileZilla FTP Client\filezilla.exe"

# Call the function to check Windows OS Version
CheckWindowsOSVersion -SoftwareName "Windows OS Version"

# Call the function to check DNS Filter
CheckSoftware -SoftwareName "DNS Filter" -FilePath "C:\Program Files\DNSFilter Agent\DNSFilter Agent.exe"