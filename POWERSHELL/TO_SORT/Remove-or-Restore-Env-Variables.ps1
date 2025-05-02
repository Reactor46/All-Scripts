# Define backup file paths
$machinePathBackupFile = "$env:TEMP\MachinePathBackup.txt"
$userPathBackupFile = "$env:TEMP\UserPathBackup.txt"
$machineRegistryBackupFile = "$env:TEMP\MachineRegistryBackup.reg"
$userRegistryBackupFile = "$env:TEMP\UserRegistryBackup.reg"

# Function to back up PATH and registry settings
function Backup-PathAndRegistry {
    param (
        [string]$scope = 'Machine'
    )

    # Backup the PATH variable
    if ($scope -eq 'Machine') {
        $currentPath = [System.Environment]::GetEnvironmentVariable('Path', 'Machine')
        [System.IO.File]::WriteAllText($machinePathBackupFile, $currentPath)
        Write-Host "Machine PATH backed up to $machinePathBackupFile"

        # Backup registry settings
        $registryPath = "HKLM\SYSTEM\ControlSet002\Control\Session Manager\Environment"
        reg export $registryPath $machineRegistryBackupFile /y
        Write-Host "Machine registry backed up to $machineRegistryBackupFile"
    } else {
        $currentPath = [System.Environment]::GetEnvironmentVariable('Path', 'User')
        [System.IO.File]::WriteAllText($userPathBackupFile, $currentPath)
        Write-Host "User PATH backed up to $userPathBackupFile"

        # Backup registry settings for User
        $registryPath = "HKCU\Environment"
        reg export $registryPath $userRegistryBackupFile /y
        Write-Host "User registry backed up to $userRegistryBackupFile"
    }
}

# Function to remove entries from the PATH environment variable
function Remove-PathEntries {
    param (
        [string]$searchTerm,
        [string]$scope = 'Machine'  # Options: 'Machine' or 'User'
    )

    # Backup current PATH and registry settings before modification
    Backup-PathAndRegistry -scope $scope

    # Get the current PATH variable
    if ($scope -eq 'Machine') {
        $currentPath = [System.Environment]::GetEnvironmentVariable('Path', 'Machine')
    } else {
        $currentPath = [System.Environment]::GetEnvironmentVariable('Path', 'User')
    }

    # Split PATH into an array
    $pathEntries = $currentPath -split ';'

    # Filter out entries containing the search term
    $filteredPathEntries = $pathEntries | Where-Object { $_ -notlike "*$searchTerm*" }

    # Join the filtered entries back into a single string
    $newPath = ($filteredPathEntries -join ';')

    # Set the new PATH variable
    if ($scope -eq 'Machine') {
        [System.Environment]::SetEnvironmentVariable('Path', $newPath, 'Machine')
    } else {
        [System.Environment]::SetEnvironmentVariable('Path', $newPath, 'User')
    }

    Write-Host "Removed entries containing '$searchTerm' from the $scope PATH."
}

# Function to restore PATH and registry settings
function Restore-PathAndRegistry {
    param (
        [string]$scope = 'Machine'
    )

    # Restore PATH variable
    if ($scope -eq 'Machine') {
        if (Test-Path $machinePathBackupFile) {
            $backupPath = Get-Content $machinePathBackupFile -Raw
            [System.Environment]::SetEnvironmentVariable('Path', $backupPath, 'Machine')
            Write-Host "Machine PATH restored from $machinePathBackupFile"
        }

        # Restore registry settings
        if (Test-Path $machineRegistryBackupFile) {
            reg import $machineRegistryBackupFile
            Write-Host "Machine registry restored from $machineRegistryBackupFile"
        }
    } else {
        if (Test-Path $userPathBackupFile) {
            $backupPath = Get-Content $userPathBackupFile -Raw
            [System.Environment]::SetEnvironmentVariable('Path', $backupPath, 'User')
            Write-Host "User PATH restored from $userPathBackupFile"
        }

        # Restore registry settings for User
        if (Test-Path $userRegistryBackupFile) {
            reg import $userRegistryBackupFile
            Write-Host "User registry restored from $userRegistryBackupFile"
        }
    }
}

# Main script execution
# Remove entries containing "java" or "openjdk"
Remove-PathEntries -searchTerm "java" -scope 'Machine'
Remove-PathEntries -searchTerm "openjdk" -scope 'Machine'
Remove-PathEntries -searchTerm "java" -scope 'User'
Remove-PathEntries -searchTerm "openjdk" -scope 'User'

# Uncomment the following lines to restore the backups
# Restore-PathAndRegistry -scope 'Machine'
# Restore-PathAndRegistry -scope 'User'
