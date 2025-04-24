# Define the list of remote servers
$ldapFilter = "(&(operatingSystem=*Windows Server*)(|(managedBy=CN=resManagedBy-AppTech-WebSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)(managedBy=CN=resManagedBy-AppTech-SiteCoreSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)))"
$servers = (Get-ADComputer -LDAPFilter $ldapFilter).Name

# Define the output file paths
$csvOutputPath = "JavaVersionReport.csv"
$htmlOutputPath = "JavaVersionReport.html"

# Initialize an array to hold the results
$results = @()

# Function to get Java version and path
function Get-JavaInfo {
    param (
        [string]$path
    )

    try {
        $javaExe = Get-Command -Name "$path\java.exe" -ErrorAction Stop
        $versionInfo = & $javaExe -version 2>&1 | Select-Object -First 1
        return [PSCustomObject]@{
            Path = $javaExe.Path
            Version = $versionInfo
        }
    } catch {
        return $null
    }
}

# Script block to execute on remote servers
$scriptBlock = {
    # Search common installation paths in the filesystem
    $systemPaths = @(
        "C:\Program Files\Java\jre*\bin",
        "C:\Program Files\Java\jdk*\bin",
        "C:\Program Files (x86)\Java\jre*\bin",
        "C:\Program Files (x86)\Java\jdk*\bin",
        "C:\Program Files\OpenJDK\jdk*\bin"
    )

    $userProfile = [System.Environment]::GetFolderPath('UserProfile')
    $userPaths = @(
        "$userProfile\AppData\Local\Programs\Java\jre*\bin",
        "$userProfile\AppData\Local\Programs\Java\jdk*\bin"
    )

    $results = @()

    # Check system paths
    foreach ($path in $systemPaths) {
        $javaInfo = Get-JavaInfo -path $path
        if ($javaInfo) {
            $results += $javaInfo
        }
    }

    # Check user paths
    foreach ($path in $userPaths) {
        $javaInfo = Get-JavaInfo -path $path
        if ($javaInfo) {
            $results += $javaInfo
        }
    }

    # Check the registry for Java installations
    $registryPaths = @(
        "HKLM:\SOFTWARE\JavaSoft\Java Runtime Environment",
        "HKLM:\SOFTWARE\JavaSoft\JDK",
        "HKCU:\SOFTWARE\JavaSoft\Java Runtime Environment",
        "HKCU:\SOFTWARE\JavaSoft\JDK"
    )

    foreach ($regPath in $registryPaths) {
        $javaVersions = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
        foreach ($version in $javaVersions) {
            $javaPath = Get-ItemProperty -Path $version.PSPath -Name "JavaHome" -ErrorAction SilentlyContinue
            if ($javaPath) {
                $javaInfo = Get-JavaInfo -path "$($javaPath.JavaHome)\bin"
                if ($javaInfo) {
                    $results += $javaInfo
                }
            }
        }
    }

    return $results
}

# Execute the command on each server
foreach ($server in $servers) {
    try {
        $javaResults = Invoke-Command -ComputerName $server -ScriptBlock $scriptBlock
        if ($javaResults) {
            foreach ($result in $javaResults) {
                $results += [PSCustomObject]@{
                    Server = $server
                    Path = $result.Path
                    Version = $result.Version
                }
            }
        } else {
            $results += [PSCustomObject]@{
                Server = $server
                Path = "N/A"
                Version = "Java not found"
            }
        }
    } catch {
        $results += [PSCustomObject]@{
            Server = $server
            Path = "N/A"
            Version = "Error accessing server: $_"
        }
    }
}

# Export results to CSV
$results | Export-Csv -Path $csvOutputPath -NoTypeInformation

# Export results to HTML
$results | ConvertTo-Html -Property Server, Path, Version -Title "Java Version Report" | Out-File -FilePath $htmlOutputPath

# Output to console
$results | Format-Table -AutoSize
