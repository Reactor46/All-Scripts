# LDAP filter to find the appropriate servers
$ldapFilter = "(&(operatingSystem=*Windows Server*)(|(managedBy=CN=resManagedBy-AppTech-WebSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)(managedBy=CN=resManagedBy-AppTech-SiteCoreSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)))"

# Path for saving the final HTML report
$htmlReportPath = "D:\Reports\Server_IIS_WebWMI_Report.html"

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

# Function to check if IIS6 WMI Compatibility (Web-WMI) is installed
function Is-IIS6WMIInstalled {
    param (
        [string]$serverName
    )
    
    $iis6WMIInstalled = Invoke-Command -ComputerName $serverName -ScriptBlock {
        Get-WindowsFeature | Where-Object { $_.Name -eq 'IIS-WMI' -and $_.Installed }
    } -ErrorAction SilentlyContinue
    
    return $iis6WMIInstalled
}

# Function to get the uptime of a server
function Get-Uptime {
    param (
        [string]$serverName
    )
    
    $uptime = Invoke-Command -ComputerName $serverName -ScriptBlock {
        (Get-Date) - (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    } -ErrorAction SilentlyContinue
    
    return $uptime
}

# Function to get IP address of the server
function Get-IP {
    param (
        [string]$serverName
    )
    
    $ipAddress = Invoke-Command -ComputerName $serverName -ScriptBlock {
        (Test-Connection -ComputerName $env:COMPUTERNAME -Count 1).Address
    } -ErrorAction SilentlyContinue
    
    return $ipAddress
}

# Function to get the 'ManagedBy' attribute from AD
function Get-ManagedBy {
    param (
        [string]$serverName
    )
    
    $managedBy = (Get-ADComputer $serverName -Properties ManagedBy).ManagedBy
    return $managedBy
}

# Function to install IIS6 WMI Compatibility (Web-WMI)
function Install-IIS6WMI {
    param (
        [string]$serverName
    )
    
    $confirmation = Read-Host "Web-WMI is not installed on $serverName. Would you like to install it? (Y/N)"
    
    if ($confirmation -eq 'Y') {
        Invoke-Command -ComputerName $serverName -ScriptBlock {
            Install-WindowsFeature -Name IIS-WMI
        }
        return "Approved"
    } else {
        return "Declined"
    }
}

# Get the list of servers from LDAP
$servers = Get-ServersFromLDAP -ldapFilter $ldapFilter

# Initialize an array to store server details for the HTML report
$serverReport = @()

# Prompt if report is required
$reportRequired = Read-Host "Do you want to generate a full report of all servers? (Y/N)"

# If report is required, generate and collect the necessary data
if ($reportRequired -eq 'Y') {
    foreach ($server in $servers) {
        Write-Host "Generating report for server: $server"

        $serverDetails = New-Object PSObject -property @{
            ServerName         = $server
            IPAddress          = (Get-IP -serverName $server)
            WindowsFeatures    = ""
            WebWMIInstalled    = "Not Checked"
            WebWMIPrompt       = "Not Checked"
            ManagedBy          = (Get-ManagedBy -serverName $server)
            Uptime             = (Get-Uptime -serverName $server)
        }

        # Get installed Windows features
        $features = Invoke-Command -ComputerName $server -ScriptBlock {
            Get-WindowsFeature | Where-Object { $_.Installed } | Select-Object -ExpandProperty Name
        } -ErrorAction SilentlyContinue
        $serverDetails.WindowsFeatures = $features -join ', '

        # Check if Web-WMI is installed
        if (Is-IIS6WMIInstalled -serverName $server) {
            $serverDetails.WebWMIInstalled = "Installed"
        } else {
            $serverDetails.WebWMIInstalled = "Not Installed"
        }

        # Add the server details to the report array
        $serverReport += $serverDetails
    }
}

# Ask if Web-WMI should be installed on servers where it is not installed
foreach ($server in $servers) {
    if ((Is-IIS6WMIInstalled -serverName $server) -eq $false) {
        $actionTaken = Install-IIS6WMI -serverName $server
        # Log the decision
        $serverDetails.WebWMIPrompt = $actionTaken
        $serverDetails.WebWMIPrompt += " by $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    }
}

# Generate the final HTML report with all the data
$finalReportHtml = $serverReport | ConvertTo-Html -Property ServerName, IPAddress, WebWMIInstalled, WebWMIPrompt, ManagedBy, Uptime -Head "<title>Server IIS and Web-WMI Report</title>" -PreContent "<h1>Server IIS and Web-WMI Report</h1>" -PostContent "<footer><p>Generated on: $(Get-Date)</p></footer>"

# Save the final HTML report to the specified path
$finalReportHtml | Out-File -FilePath $htmlReportPath

Write-Host "HTML report generated at: $htmlReportPath"
