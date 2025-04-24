###########################################################################
# IIS Wildcard Certificate Quick Swap
###########################################################################
## Remove Comments From Variables Below to Add Friendly Names to Script to
## Bypass User Input Requirement
# $oldcertFriendlyName = ""
# $newcertFriendlyName = ""

## Clear Old Variable Settings
Clear-Variable -Name newcert
Clear-Variable -Name oldcert

## Retrieve Old Certificate Friendly Name From Console if Not Pre Declared
if ($oldcertFriendlyName -eq $null) {
    Write-Host "Old Certificate Friendly Name:"
    $oldcertFriendlyName = Read-Host
} else {
    Write-Output "If the OLD Certificate Friendly Name below is correct, press any Enter to continue."
    Write-Output $oldcertFriendlyName
    Pause
}

## Retrieve Old Certificate Friendly Name From Console if Not Pre Declared
if ($newcertFriendlyName -eq $null) {
    Write-Host "New Certificate Friendly Name"
    $newcertFriendlyName = Read-Host
} else {
    Write-Output "If the NEW Certificate Friendly Name below is correct, press any Enter to continue."
    Write-Output $newcertFriendlyName
    Pause
}

## Set the Certificates to Appropriate Variables
$newcert = Get-ChildItem Cert:\LocalMachine\My\ | where{$_.FriendlyName -eq $newcertFriendlyName}
$oldcert = Get-ChildItem Cert:\LocalMachine\My\ | where{$_.FriendlyName -eq $oldcertFriendlyName}

## Retrieve Sites Using Old Certificate (Based on Thumbprint)
$sites = Get-WebBinding -port 443 | where {$_.certificateHash -eq $oldcert.Thumbprint} 


if ($newcert -eq $null) {
    # Make Sure newcert Variable is not Null Before Proceeding
    Write-Output "New Certificate Cannot Be Located With..."
    Write-Output $newcertFriendlyName
} elseif ($oldcert -eq $null) {
    # Make Sure oldert Variable is not Null Before Proceeding
    Write-Output "Old Certificate Cannot Be Located With..."
    Write-Output $oldcertFriendlyName
} else {
    # Update Each Site to Use New Certificate
    foreach ($site in $sites) {
        # Set Location to IIS SSL Bindings
        cd IIS:\sslBindings
    
        # Update Site to Use New Certificate
        $site.AddSslCertificate($newcert.getcerthashstring(), "my")

        # List Sites As Changes Occur
        Write-Output $site
    }
}