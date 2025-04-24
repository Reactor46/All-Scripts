# Define variables
$installPath = "C:\Program Files\OpenJDK"
$zipPath = "\\fbv-wbdv20-d01\D$\Apps\OpenJDK\openjdk-23_windows-x64_bin.zip"
$msiPath = "\\fbv-wbdv20-d01\D$\Apps\OpenJDK\jdk-23_windows-x64_bin.msi"
$tempZipFile = "$env:TEMP\openjdk.zip"
$jdkUrl = "https://download.java.net/java/GA/jdk23/36/binaries/openjdk-23_windows-x64_bin.zip"  # URL for latest OpenJDK
$msijdkUrl = "https://download.oracle.com/java/23/latest/jdk-23_windows-x64_bin.msi"
$solrServiceName = "solr-8983"

# Function to get the current Java version
function Get-JavaVersion {
    try {
        $javaVersion = & "$installPath\bin\java.exe" -version 2>&1
        if ($javaVersion -match 'version "(\d+)\.(\d+)\.(\d+)_\d+"') {
            return [int]$matches[1]  # Return the major version
        }
    } catch {
        return 0  # If Java is not found, return 0
    }
}

# Function to stop Solr service if running
function Stop-SolrService {
    if (Get-Service -Name $solrServiceName -ErrorAction SilentlyContinue) {
        Write-Host "Stopping Solr service..."
        Stop-Service -Name $solrServiceName -Force
        Start-Sleep -Seconds 5  # Wait for a few seconds to ensure it stops
    }
}

# Function to remove existing environment variables
function Remove-EnvironmentVariables {
    $varsToRemove = @('JAVA_HOME', 'PATH')

    foreach ($var in $varsToRemove) {
        if ([System.Environment]::GetEnvironmentVariable($var, 'Machine')) {
            Write-Host "Removing environment variable: $var"
            [System.Environment]::SetEnvironmentVariable($var, $null, 'Machine')
        }
    }
}

# Locate existing OpenJDK version
$currentVersion = Get-JavaVersion
Write-Host "Current OpenJDK Version: $currentVersion"

# Verify if the current version is less than 23
if ($currentVersion -ge 23) {
    Write-Host "OpenJDK version is $currentVersion, which is 23 or greater. No update needed."
    exit
}

# Check for the latest ZIP version
if (-Not (Test-Path $zipPath)) {
    Write-Host "Local ZIP not found at $zipPath. Downloading the latest OpenJDK ZIP..."
    Invoke-WebRequest -Uri $jdkUrl -OutFile $tempZipFile
    $zipToInstall = $tempZipFile
} else {
    Write-Host "Using local ZIP file: $zipPath"
    $zipToInstall = $zipPath
}

# Stop Solr service if running
Stop-SolrService

# Remove the existing OpenJDK installation if it exists
if (Test-Path $installPath) {
    Write-Host "Removing existing OpenJDK installation at $installPath..."
    Remove-Item -Path $installPath -Recurse -Force
}

# Create the installation directory
New-Item -ItemType Directory -Path $installPath -Force

# Extract the ZIP file to the installation path
Expand-Archive -Path $zipToInstall -DestinationPath $installPath -Force

# Remove existing environment variables
Remove-EnvironmentVariables

# Set new environment variables
[System.Environment]::SetEnvironmentVariable('JAVA_HOME', "$($installPath)\jdk-23", 'Machine')
$newPath = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + "%JAVA_HOME%\bin\"
[System.Environment]::SetEnvironmentVariable('Path', $newPath, 'Machine')

$env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') 

# Verify the new installation
$newVersion = Get-JavaVersion
Write-Host "New OpenJDK Version: $newVersion"

# Check if the new version is 23 or greater
if ($newVersion -ge 23) {
    Write-Host "OpenJDK updated successfully to version $newVersion."
} else {
    Write-Host "OpenJDK update failed. Current version is $newVersion."
    exit
}

# Verify version using java.exe
$versionCheck = & "$installPath\java.exe" -version
Write-Host "Java Version Check Output: $versionCheck"

# Start Solr service if it exists
if (Get-Service -Name $solrServiceName -ErrorAction SilentlyContinue) {
    Write-Host "Starting Solr service..."
    Start-Service -Name $solrServiceName
    Start-Sleep -Seconds 5  # Wait for a few seconds to ensure it starts
}

# Verify Solr service is running
if ((Get-Service -Name $solrServiceName).Status -eq 'Running') {
    Write-Host "Solr service is running."
} else {
    Write-Host "Solr service failed to start."
}
