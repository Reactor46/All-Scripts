# Configuration Variables
#$AppPool = "ksprod-new-cd.ksnet.com"  # Application Pool Name
$AppPool = 'test'
$site = "https://www.kelseycareadvantage.com"  # Base Site URL to test redirects
$rewriteMapsFile = "C:\inetpub\wwwroot\ksprod-new-cd.ksnet.com\rewriteMaps.config"  # Path to rewriteMaps.config
$rewriteBackups = "C:\inetpub\wwwroot\ksprod-new-cd.ksnet.com\Rewrite BAKs"  # Backup directory
$webSiteURL = $site  # Base URL for testing redirects

# New Redirect Entries (add your own redirects here)
$newRedirects = @(
    @{Key = "oldurl1"; Value = "https://www.newsite.com/newurl1"},
    @{Key = "oldurl2"; Value = "https://www.newsite.com/newurl2"},
    @{Key = "oldurl3"; Value = "https://www.newsite.com/newurl3"}
)

# Step 1: Backup the current rewriteMaps.config
if (Test-Path $rewriteMapsFile) {
    $backupDir = Join-Path $rewriteBackups (Get-Date -Format "yyyyMMdd_HHmmss")
    New-Item -ItemType Directory -Path $backupDir -Force
    Copy-Item -Path $rewriteMapsFile -Destination $backupDir
    Write-Host "Backup created at $backupDir"
} else {
    Write-Host "rewriteMaps.config not found at $rewriteMapsFile"
    exit
}

# Step 2: Load the rewriteMaps.config file and check if 'rewriteMap' exists
if (Test-Path $rewriteMapsFile) {
    [xml]$xml = Get-Content $rewriteMapsFile

    # Find or create the 'rewriteMap' element (named ExampleRedirects)
    $rewriteMap = $xml.rewriteMaps.rewriteMap | Where-Object { $_.name -eq "ExampleRedirects" }

    # If not found, create the rewriteMap
    if (-not $rewriteMap) {
        $rewriteMap = $xml.rewriteMaps.AppendChild($xml.CreateElement("rewriteMap"))
        $rewriteMap.SetAttribute("name", "ExampleRedirects")
    }

    # Step 3: Add the new redirects to the rewriteMap
    foreach ($redirect in $newRedirects) {
        $addElement = $xml.CreateElement("add")
        $addElement.SetAttribute("key", $redirect.Key)
        $addElement.SetAttribute("value", $redirect.Value)
        $rewriteMap.AppendChild($addElement)
    }

    # Save the updated XML back to the file
    $xml.Save($rewriteMapsFile)
    Write-Host "New redirects added to rewriteMaps.config."
} else {
    Write-Host "rewriteMaps.config file does not exist."
    exit
}

# Step 4: Recycle the Application Pool to apply changes
Write-Host "Recycling Application Pool: $AppPool"
Restart-WebAppPool -Name $AppPool

# Step 5: Verify the Redirects are working (Return 301 or 302)
# Test each redirect URL and check if it's returning a 301 or 302 HTTP status code

foreach ($redirect in $newRedirects) {
    $oldUrl = "$webSiteURL/$($redirect.Key)"
    try {
        $response = Invoke-WebRequest -Uri $oldUrl -Method Head -ErrorAction Stop
        $statusCode = $response.StatusCode

        if ($statusCode -eq 301 -or $statusCode -eq 302) {
            Write-Host "Redirect for $oldUrl is working. Status code: $statusCode"
        } else {
            Write-Host "Unexpected status code for $oldUrl: $statusCode"
        }
    } catch {
        Write-Host "Error with URL $oldUrl: $_"
    }
}

Write-Host "Redirect verification complete."
