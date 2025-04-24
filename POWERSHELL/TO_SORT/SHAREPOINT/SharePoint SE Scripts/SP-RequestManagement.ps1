# SharePoint Server Request Management
# Configure Request Manager in SharePoint Server
# https://learn.microsoft.com/en-us/sharepoint/security-for-sharepoint-server/configure-request-manager-in-sharepoint-server

# Enable routing and throttling for all web applications	
Get-SPWebApplication | Set-SPRequestManagementSettings -RoutingEnabled $true -ThrottlingEnabled $true

# Enable routing with static weighting for all web applications	
Get-SPWebApplication | Get-SPRequestManagementSettings | Set-SPRequestManagementSettings -RoutingEnabled $true -ThrottlingEnabled $false -RoutingWeightScheme Static

# Return a list of routing targets for all available web applications	
Get-SPWebApplication | Get-SPRequestManagementSettings | Get-SPRoutingMachineInfo -Availability Available

# Add a new routing target for a specified web application.	
$web=Get-SPWebApplication -Identity "URL of web application"
$rm=Get-SPRequestManagementSettings -Identity $web
Add-SPRoutingMachineInfo -RequestManagementSettings $rm -Name "MachineName" -Availability Available

# Edit an existing routing target's availability and static weight for a specified web application.
$web=Get-SPWebApplication -Identity "URL of web application"
$rm=Get-SPRequestManagementSettings -Identity $web
$m=Get-SPRoutingMachineInfo -RequestManagementSettings $rm -Name "MachineName"
Set-SPRoutingMachineInfo -Identity $m -Availability Unavailable

# Remove a routing target from a specified web application.
$web=Get-SPWebApplication -Identity "URL of web application"
$rm=Get-SPRequestManagementSettings -Identity $web
$m=Get-SPRoutingMachineInfo -RequestManagementSettings $rm -Name "MachineName"
Remove-SPRoutingMachineInfo -Identity $M

# Along with creating a performance monitor log file, the verbose logging level can be enabled by using the following Microsoft PowerShell syntax:
Set-SPLogLevel "Request Management" -TraceSeverity Verbose


