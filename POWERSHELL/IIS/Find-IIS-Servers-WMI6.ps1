# LDAP filter to find the appropriate servers
$ldapFilter = "(&(operatingSystem=*Windows Server*)(|(managedBy=CN=resManagedBy-AppTech-WebSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)(managedBy=CN=resManagedBy-AppTech-SiteCoreSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)))"

# Path for saving the HTML report
$htmlReportPath = "D:\Reports\IIS_Installed_Report.html"

# Function to get server list based on LDAP filter
function Get-ServersFromLDAP {
    param (
        [string]$ldapFilter
    )
    
    $domain = "ksnet.com"
    $searchBase = "LDAP://DC=ksnet,DC=com"
    $searcher = New-Object DirectoryServices.DirectorySearcher([ADSI]$searchBase, $ldapFilter)

    $results = $searcher.FindAll()
    $servers = @()

    foreach ($result in $results) {
        $serverName = $result.Properties["name"]
        $servers += $serverName
    }
    
    return $servers
}

# Function to check if IIS is installed on the remote server
function Is-IISInstalled {
    param (
        [string]$serverName
    )
    
    $iisInstalled = Invoke-Command -ComputerName $serverName -ScriptBlock {
        Get-WindowsFeature | Where-Object { $_.Name -eq 'Web-Server' -and $_.Installed }
    } -ErrorAction SilentlyContinue
    
    return $iisInstalled
}

# Function to check if IIS6 WMI Compatibility is installed
function Is-IIS6WMIInstalled {
    param (
        [string]$serverName
    )
    
    $iis6WMIInstalled = Invoke-Command -ComputerName $serverName -ScriptBlock {
        Get-WindowsFeature | Where-Object { $_.Name -eq 'IIS-WMI' -and $_.Installed }
    } -ErrorAction SilentlyContinue
    
    return $iis6WMIInstalled
}

# Function to install IIS6 WMI Compatibility
function Install-IIS6WMI {
    param (
        [string]$serverName
    )
    
    $confirmation = Read-Host "IIS6 WMI Compatibility is not installed on $serverName. Would you like to install it? (Y/N)"
    
    if ($confirmation -eq 'Y') {
        Invoke-Command -ComputerName $serverName -ScriptBlock {
            Install-WindowsFeature -Name IIS-WMI
        }
        Write-Host "IIS6 WMI Compatibility has been installed on $serverName."
    } else {
        Write-Host "Skipping installation of IIS6 WMI Compatibility on $serverName."
    }
}

# Function to check server online status
function Is-ServerOnline {
    param (
        [string]$serverName
    )
    
    $pingResult = Test-Connection -ComputerName $serverName -Count 1 -Quiet
    return $pingResult
}

# Get the list of servers from LDAP
$servers = Get-ServersFromLDAP -ldapFilter $ldapFilter

# Initialize an array to store server details for the HTML report
$serverReport = @()

# Loop through each server and gather details
foreach ($server in $servers) {
    $serverStatus = New-Object PSObject -property @{
        ServerName           = $server
        IISInstalled         = "Not Checked"
        IIS6WMIInstalled     = "Not Checked"
        IsOnline             = "Not Checked"
    }

    # Check if the server is online
    if (Is-ServerOnline -serverName $server) {
        $serverStatus.IsOnline = "Online"
    } else {
        $serverStatus.IsOnline = "Offline"
        $serverReport += $serverStatus
        continue
    }

    # Check if IIS is installed
    if (Is-IISInstalled -serverName $server) {
        $serverStatus.IISInstalled = "Installed"
    } else {
        $serverStatus.IISInstalled = "Not Installed"
    }

    # Check if IIS6 WMI Compatibility is installed
    if (Is-IIS6WMIInstalled -serverName $server) {
        $serverStatus.IIS6WMIInstalled = "Installed"
    } else {
        $serverStatus.IIS6WMIInstalled = "Not Installed"
        
        # Ask for installation if not installed
        Install-IIS6WMI -serverName $server
    }

    # Add the server status to the report array
    $serverReport += $serverStatus
}

# Generate the HTML report
$reportHtml = $serverReport | ConvertTo-Html -Property ServerName, IISInstalled, IIS6WMIInstalled, IsOnline -Head "<title>IIS and IIS6 WMI Compatibility Report</title>" -PreContent "<h1>IIS and IIS6 WMI Compatibility Report</h1>" -PostContent "<footer><p>Generated on: $(Get-Date)</p></footer>"

# Save the HTML report to the specified path
$reportHtml | Out-File -FilePath $htmlReportPath

Write-Host "HTML report generated at: $htmlReportPath"
