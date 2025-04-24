# Define variables
$installPath = "C:\Program Files\OpenJDK"  # Existing installation path
$localZipPath = "\\fbv-wbdv20-d01\D$\Apps\OpenJDK\openjdk-23_windows-x64_bin.zip"  # Local path to check for the ZIP
$tempZipFile = "$env:TEMP\openjdk.zip"
$jdkUrl = "https://download.java.net/java/GA/jdk23/3c5b90190c68498b986a97f276efd28a/37/GPL/openjdk-23_windows-x64_bin.zip"  # Update with the latest URL
$solrServiceName = "solr-8983"

# Function to get the current Java version
function Get-JavaVersion {
    try {
        $javaVersion = & java -version 2>&1
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
function Get-JavaHome {
# Run the command to show Java settings
$javaSettings = & java -XshowSettings:properties -version 2>&1

# Initialize a variable to hold JAVA_HOME
$javaHome = $null

# Search for the java.home line in the output
foreach ($line in $javaSettings) {
    if ($line -match 'java\.home\s+=\s+(.*)') {
        $javaHome = $matches[1].Trim()  # Get the value and trim whitespace
        break
    }
}
$javahome
# Check and output the JAVA_HOME
#if ($javaHome) {
#    Write-Host "JAVA_HOME is set to: $javaHome"
#} else {
#    Write-Host "JAVA_HOME could not be found."
}

# Function to remove existing environment variables
function Remove-EnvironmentVariables {
    $varsToRemove = @('JAVA_HOME', 'JRE_HOME', 'JDK_HOME', 'CLASSPATH')

    foreach ($var in $varsToRemove) {
        if ([System.Environment]::GetEnvironmentVariable($var, 'Machine')) {
            Write-Host "Removing environment variable: $var"
            [System.Environment]::SetEnvironmentVariable($var, $null, 'Machine')
        }
    }

    # Remove the specific PATH entry for OpenJDK
    $currentPath = [System.Environment]::GetEnvironmentVariable('Path', 'Machine')
    $pathToRemove = $javahome
    
    if ($currentPath -like "*$pathToRemove*") {
        Write-Host "Removing path entry: $pathToRemove"
        $newPath = $currentPath -replace [regex]::Escape($pathToRemove + ";?"), ''
        [System.Environment]::SetEnvironmentVariable('Path', $newPath, 'Machine')
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
if (-Not (Test-Path $localZipPath)) {
    Write-Host "Local ZIP not found at $localZipPath. Downloading the latest OpenJDK ZIP..."
    Invoke-WebRequest -Uri $jdkUrl -OutFile $tempZipFile
    $zipToInstall = $tempZipFile
} else {
    Write-Host "Using local ZIP file: $localZipPath"
    $zipToInstall = $localZipPath
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
$env:JAVA_HOME = "$installPath\jdk-23"
[System.Environment]::SetEnvironmentVariable('JAVA_HOME', "$installPath\jdk-23", 'Machine')
[System.Environment]::SetEnvironmentVariable('Path', [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + "$installPath\jdk-23\bin", 'Machine')

# Clean up if downloaded
if ($zipToInstall -eq $tempZipFile) {
    Remove-Item $tempZipFile
}

# Verify the installation
$newVersion = Get-JavaVersion
Write-Host "New OpenJDK Version: $newVersion"

# Check if the new version is 23 or greater
if ($newVersion -ge 23) {
    Write-Host "OpenJDK updated successfully to version $newVersion."
} else {
    Write-Host "OpenJDK update failed. Current version is $newVersion."
    exit
}

# Start Solr service
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
