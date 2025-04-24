param (
    [string]$XML = "rewriteMaps.config",
    [string]$CSV = "redirects-to-remove.csv",
    [switch]$Unattended
)

# Ensure UTF8 encoding for all file operations
[System.Text.Encoding]::UTF8

# Load CSV data
$csvData = Import-Csv -Path $CSV -Encoding UTF8

# Load and parse XML data
[xml]$redirectXML = Get-Content -Path $XML -Encoding UTF8
$rdList = $redirectXML.rewriteMaps.rewriteMap.add

# Backup the original XML file
$backupFile = "$XML.bak"
Copy-Item -Path $XML -Destination $backupFile

# Track changes for event logging
$changes = @()

foreach ($csvEntry in $csvData) {
    $key = $csvEntry.key
    $value = $csvEntry.value
    
    $xmlEntry = $rdList | Where-Object { $_.key -eq $key -and $_.value -eq $value }
    
    if ($xmlEntry) {
        foreach ($entry in $xmlEntry) {
            # Comment out the entry
            $entry.OuterXml = "<!--$($entry.OuterXml)-->"

            $changes += "Commented out entry with key: $key and value: $value"
        }
    }
}

# Save the updated XML
$redirectXML.Save($XML)

# Log summary of changes
foreach ($change in $changes) {
    Write-Host $change
}

# Additional logging for unattended mode
if ($Unattended) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $logMessage = "$timestamp - Changes applied by $user. Backup file: $backupFile"

    foreach ($change in $changes) {
        $logMessage += "`n$change"
    }

    # Function to log events
    function Log-Event {
        param (
            [string]$EventType,
            [string]$Message
        )
        Write-EventLog -LogName $EventLogName -Source $EventType -EventId 1 -EntryType Information -Message $Message
    }

    # Create EventLog if it doesn't exist
    $EventLogName = "IIS-ReWriteMapsMonitor"
    $sources = @("DisabledEntry", "UpdatedEntry", "NewEntry")

    if (-not (Get-EventLog -LogName $EventLogName -ErrorAction SilentlyContinue)) {
        New-EventLog -LogName $EventLogName -Source $sources
    }

    Log-Event -EventType "DisabledEntry" -Message $logMessage
}

#.\ReWriteMapsEditor.ps1 -XML "rewriteMaps.config" -CSV "redirects-to-remove.csv" -Unattended
