# Install the required module
if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber -Scope CurrentUser
}
#Import-Module -Name PnP.PowerShell 
Import-Module -Name Microsoft.Online.SharePoint.PowerShell
# Connect to SharePoint Online Admin Center
#$adminUrl = "https://ksclinic-admin.sharepoint.com"
#Setup usercredential
#$username = "k24696@ksnet.com"
#$password = ""
#$cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $userName, $(convertto-securestring $Password -asplaintext -force)

Connect-SPOService -Url "https://ksclinic.sharepoint.com" -AuthenticationUrl "https://ksclinic-admin.sharepoint.com" -ModernAuth $true -Verbose
#Connect-PnPOnline -Url "https://ksclinic-admin.sharepoint.com"

# Get all site collections
$siteCollections = Get-SPOSite -Limit All
$totalSites = $siteCollections.Count
$siteIndex = 0
 
# Array to store results
$results = @()
 
foreach ($site in $siteCollections) {
    $siteIndex++
    Write-Progress -Activity "Checking sites" -Status "Processing $siteIndex out of $totalSites" -PercentComplete (($siteIndex / $totalSites) * 100)
    Write-Host "Checking site: $($site.Url)"
    # Get the site users
    $siteUsers = Get-SPOUser -Site $site.Url | Where-Object { $_.LoginName -ne "Everyone except external users" }
 
    foreach ($user in $siteUsers) {
        # Get the user's permission level
        $permissionLevels = $user.Roles -join ", "
 
        $results += [PSCustomObject]@{
            Username = $user.LoginName
            SiteUrl = $site.Url
            Permissions = $permissionLevels
        }
    }
}
 
# Export results to CSV
$outputPath = "D:\Reports\SPO-Heather\UserAccessReport.csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation
 
Write-Host "Results exported to $outputPath"