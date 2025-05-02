# Define the list of servers

$servers = @("FBV-SCORINT-D01.ksnet.com",
"FBV-SCORDEV-D12.ksnet.com",
"FBV-SCRCM10-T01.ksnet.com",
"FBV-SCRCD10-T01.ksnet.com",
"FBV-SCORINT-D02.ksnet.com",
"FBV-SCRCD10-P02.ksnet.com",
"FBV-SCRCM10-P01.ksnet.com",
"FBV-SCRCD10-T02.ksnet.com",
"FBV-SCRCD10-P01.ksnet.com",
"FBV-SCORDEV-D08.ksnet.com",
"FBV-SCRXC10-T01.ksnet.com",
"FBV-SCORDEV-D01.ksnet.com") # Replace with your server names

#$servers = @("FBV-WBDV20-D01.ksnet.com")

# Define the 7-Zip installer URL or path
#$installerPath = "https://www.7-zip.org/a/7z2408-x64.msi" # Example URL for version 24.8.0
#$installerPath = "\\fbv-wbdv20-d01\D$\Apps\7z2408-x64.msi"
# Alternatively, if you have a local installer, use a path like "C:\Path\To\7z2408-x64.msi"

foreach ($server in $servers) {
    try {
        # Check if the server is reachable
        if (Test-Connection -ComputerName $server -Count 2 -Quiet) {
            # Check if 7-Zip is installed
            #$sevenZip = Get-WmiObject -Class Win32_Product -ComputerName $server | Where-Object { $_.Name -like "7-Zip*" }
            $sevenZip = Get-CimInstance -ClassName Win32_Product -ComputerName $server | Where-Object { $_.Name -like "7-Zip*" }
            
            if ($sevenZip) {
                # Get the version of 7-Zip
                $version = [version]$sevenZip.Version
                
                # Define the minimum version to check against
                $minVersion = [version]"24.08.00.0"
                
                if ($version -lt $minVersion) {
                    Write-Host "Updating 7-Zip on $server (current version: $version)"
                    # Execute the installer
                   # Let's go directly to the website and see what it lists as the current version
                $BaseUri = "https://www.7-zip.org/"
                $BasePage = Invoke-WebRequest -Uri ( $BaseUri + 'download.html' ) -UseBasicParsing
                    # Determine bit-ness of O/S and download accordingly
                    if ( [System.Environment]::Is64BitOperatingSystem ) {
                    # The most recent 'current' (non-beta/alpha) is listed at the top, so we only need the first.
                    $ChildPath = $BasePage.Links | Where-Object { $_.href -like '*7z*-x64.msi' } | Select-Object -First 1 | Select-Object -ExpandProperty href
                    } else {
                    # The most recent 'current' (non-beta/alpha) is listed at the top, so we only need the first.
                    $ChildPath = $BasePage.Links | Where-Object { $_.href -like '*7z*.msi' } | Select-Object -First 1 | Select-Object -ExpandProperty href
            }
 
                # Let's build the required download link
                $DownloadUrl = $BaseUri + $ChildPath
 
                    Write-Host "Downloading the latest 7-Zip to the temp folder"
                    Invoke-WebRequest -Uri $DownloadUrl -OutFile "$env:TEMP\$( Split-Path -Path $DownloadUrl -Leaf )" | Out-Null
                    Write-Host "Installing the latest 7-Zip"
                    Start-Process -FilePath "$env:SystemRoot\system32\msiexec.exe" -ArgumentList "/package", "$env:TEMP\$( Split-Path -Path $DownloadUrl -Leaf )", "/passive" -Wait

                } else {

                    Write-Host "7-Zip on $server is up to date (version: $version)"
                }

            } else {
                
                Write-Host "7-Zip is not installed on $server. Skipping update."
            }

        } else {

            Write-Host "$server is not reachable."

        }

    } catch {

        Write-Host "Error connecting to $server : $_"

    }
}
