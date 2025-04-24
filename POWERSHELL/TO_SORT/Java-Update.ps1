# Get the current machine PATH
$currentPath = [System.Environment]::GetEnvironmentVariable("PATH", "Machine")
# Define the old and new paths
$oldPath = "C:\Program Files\OpenJDK\jdk-19.0.2\bin"
$newPath = "C:\Program Files\OpenJDK\jdk-23\bin"

# Replace the old path with the new path
$updatedPath = $currentPath -replace [regex]::Escape($oldPath), $newPath
# Set the updated PATH variable
[System.Environment]::SetEnvironmentVariable("PATH", $updatedPath, "Machine")
# Get the updated machine PATH
$updatedMachinePath = [System.Environment]::GetEnvironmentVariable("PATH", "Machine")
$updatedMachinePath

[System.Environment]::GetEnvironmentVariables([System.EnvironmentVariableTarget]::Machine)
[System.Environment]::SetEnvironmentVariable("JAVA_HOME","C:\Program Files\OpenJDK\jdk-23","Machine")

cmd /c java -version

Get-Service -Name Solr-8983
Restart-Service -Name Solr-8983

