# Define the list of remote servers
# Define the list of remote servers
$ldapFilter = "(&(operatingSystem=*Windows Server*)(|(managedBy=CN=resManagedBy-AppTech-WebSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)(managedBy=CN=resManagedBy-AppTech-SiteCoreSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)))"
#$servers = (Get-ADComputer -LDAPFilter $ldapFilter).Name
$servers = @("fbv-scslr10-t01","fbv-scslr10-t02","fbv-scslr10-t03","fbv-scslr10-p01","fbv-scslr10-p02","fbv-scslr10-p03","fbv-scordev-d08")
# Command to search for java.exe and get its version
$scriptBlock = {
    $javaPaths = @(
        "C:\Program Files\Java\jre*\bin\java.exe",
        "C:\Program Files\Java\jdk*\bin\java.exe",
        "C:\Program Files (x86)\Java\jre*\bin\java.exe",
        "C:\Program Files (x86)\Java\jdk*\bin\java.exe",
        "C:\Program Files\OpenJDK\jdk*\bin\java.exe",
        "C:\Program Files\Solr\java\bin\java.exe"

    )

    $javaVersion = $null

    foreach ($path in $javaPaths) {
        $javaExe = Get-ChildItem -Path $path -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($javaExe) {
            $javaVersion = & $javaExe -version 2>&1 | Select-Object -First 1
            break
        }
    }

    if (-not $javaVersion) {
        "Java is not installed."
    } else {
        $javaVersion
    }
}

# Execute the command on each server
foreach ($server in $servers) {
    try {
        $javaVersion = Invoke-Command -ComputerName $server -ScriptBlock $scriptBlock
        Write-Output "$server : $javaVersion"
    } catch {
        Write-Output "$server : Error accessing server. $_"
    }
}
