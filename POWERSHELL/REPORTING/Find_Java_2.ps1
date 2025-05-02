# Define the list of remote servers
$ldapFilter = "(&(operatingSystem=*Windows Server*)(|(managedBy=CN=resManagedBy-AppTech-WebSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)(managedBy=CN=resManagedBy-AppTech-SiteCoreSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)))"
$servers = (Get-ADComputer -LDAPFilter $ldapFilter).Name

# Command to get Java version
$scriptBlock = {
    try {
        # Check if Java is installed and get the version
        & "java" -version 2>&1 | Select-Object -First 1
    } catch {
        "Java is not installed"
    }
}

# Execute the command on each server
foreach ($server in $servers) {
    $javaVersion = Invoke-Command -ComputerName $server -ScriptBlock $scriptBlock
    Write-Output "$server : $javaVersion"
}
