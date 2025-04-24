# Define constants
$webRoot = "C:\inetpub\wwwroot"
$configDir = "$webRoot\Config_Backups"
$changesFile = "C:\path\to\rewriteMaps-Changes.config"
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

# Ensure Config_Backups directory exists
if (-not (Test-Path $configDir)) {
    New-Item -ItemType Directory -Path $configDir
}

# Backup configuration files
$filesToBackup = @("$webRoot\web.config", "$webRoot\rewriteMaps.config", "$webRoot\rewriteRules.config")
foreach ($file in $filesToBackup) {
    if (Test-Path $file) {
        Copy-Item -Path $file -Destination "$configDir\$(Split-Path -Leaf $file).$timestamp.bak"
    }
}

# Load the rewriteMaps and changes
[xml]$rewriteMaps = Get-Content "$webRoot\rewriteMaps.config"
[xml]$changes = Get-Content $changesFile

# Process changes
$report = @()
foreach ($site in $changes.ChildNodes) {
    $siteUrl = switch ($site.Name) {
        "KSC" { "https://www.kelsey-seybold.com" }
        "KCA" { "https://www.kelseycareadvantage.com" }
        default { continue }
    }

    $rewriteMap = $rewriteMaps.rewriteMaps.rewriteMap | Where-Object { $_.name -eq "PermRedirects" }

    # Process new entries
    foreach ($entry in $site.New.add) {
        $newEntry = $rewriteMap.add | Where-Object { $_.key -eq $entry.key }
        if (-not $newEntry) {
            $newNode = $rewriteMaps.CreateElement("add")
            $newNode.SetAttribute("key", $entry.key)
            $newNode.SetAttribute("value", $entry.value)
            $rewriteMap.AppendChild($newNode)
            $report += "Added new entry for $siteUrl : $($entry.key) -> $($entry.value)"
        }
    }

    # Process updated entries
    foreach ($entry in $site.Update.add) {
        $updateEntry = $rewriteMap.add | Where-Object { $_.key -eq $entry.key }
        if ($updateEntry) {
            $updateEntry.SetAttribute("value", $entry.value)
            $report += "Updated entry for $siteUrl : $($entry.key) -> $($entry.value)"
        }
    }

    # Process removed entries
    foreach ($entry in $site.Remove.add) {
        $removeEntry = $rewriteMap.add | Where-Object { $_.key -eq $entry.key }
        if ($removeEntry) {
            $rewriteMap.RemoveChild($removeEntry)
            $report += "Removed entry for $siteUrl : $($entry.key)"
        }
    }
}

# Save updated rewriteMaps.config
$rewriteMaps.Save("$webRoot\rewriteMaps.config")

# Recycle the application pool
$appPoolName = "YourAppPoolName"  # Replace with your actual app pool name
Restart-WebAppPool -Name $appPoolName

# Validate redirects
foreach ($site in $changes.ChildNodes) {
    $siteUrl = switch ($site.Name) {
        "KSC" { "https://www.kelsey-seybold.com" }
        "KCA" { "https://www.kelseycareadvantage.com" }
        default { continue }
    }

    foreach ($entry in $site.New.add) {
        $response = Invoke-WebRequest -Uri "$siteUrl$($entry.key)" -Method Head
        if ($response.StatusCode -eq 301 -or $response.StatusCode -eq 302) {
            $report += "Redirect validation passed for $siteUrl : $($entry.key)"
        } else {
            $report += "Redirect validation failed for $siteUrl : $($entry.key)"
        }
    }
}

# Generate HTML report
$htmlReport = "<html><body><h1>Rewrite Map Changes Report</h1><ul>"
foreach ($line in $report) {
    $htmlReport += "<li>$line</li>"
}
$htmlReport += "</ul></body></html>"
Set-Content -Path "$webRoot\rwChanges.html" -Value $htmlReport

Write-Output "Process completed. Report generated at $webRoot\rwChanges.html"
